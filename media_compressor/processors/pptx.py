import io
import tempfile
import traceback
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from ..constants import AUDIO_EXTS, CONTENT_TYPES, IMAGE_EXTS, VIDEO_EXTS
from ..deps import HAS_FFMPEG, HAS_PILLOW, Image
from ..utils.image_utils import _has_alpha, _resize_if_needed, _to_rgb
from .video import _compress_video_in_memory


def _pptx_build_dpi_limit_map(all_entries: dict, max_dpi: int) -> Dict[str, int]:
    from posixpath import dirname, join, normpath

    EMU_PER_INCH = 914400
    R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

    limit_map: Dict[str, int] = {}

    for name in all_entries:
        if "/_rels/" not in name or not name.endswith(".rels"):
            continue
        if "/slides/_rels/" not in name:
            continue

        try:
            rels_root = ET.fromstring(all_entries[name])
        except ET.ParseError:
            continue

        slide_dir = dirname(name.replace("/_rels", ""))

        rid_to_media: Dict[str, str] = {}
        for rel in rels_root:
            if "image" not in rel.get("Type", ""):
                continue
            rid = rel.get("Id", "")
            target = rel.get("Target", "")
            if not rid or not target:
                continue
            if target.startswith("/"):
                full = target.lstrip("/")
            else:
                full = normpath(join(slide_dir, target)).replace("\\", "/")
            rid_to_media[rid] = full

        if not rid_to_media:
            continue

        slide_xml = name.replace("/_rels/", "/").removesuffix(".rels")
        if slide_xml not in all_entries:
            continue

        try:
            slide_root = ET.fromstring(all_entries[slide_xml])
        except ET.ParseError:
            continue

        parent_map = {child: parent for parent in slide_root.iter() for child in parent}

        for elem in slide_root.iter():
            local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if local != "blip":
                continue

            rid = elem.get(R_EMBED)
            if not rid or rid not in rid_to_media:
                continue

            media_path = rid_to_media[rid]

            cx = cy = None
            node = elem
            for _ in range(10):
                node = parent_map.get(node)
                if node is None:
                    break
                for child in node.iter():
                    child_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if child_local == "ext" and "cx" in child.attrib:
                        try:
                            cx = int(child.get("cx", 0))
                            cy = int(child.get("cy", 0))
                        except ValueError:
                            pass
                        break
                if cx is not None:
                    break

            if not cx or cx <= 0:
                continue

            limit_px = int(max(cx, cy) / EMU_PER_INCH * max_dpi)
            limit_map[media_path] = max(limit_map.get(media_path, 0), limit_px)

    return limit_map


def _pptx_compress_image_entry(
    name: str, data: bytes, quality: int, max_dim: int, dpi_limit: Optional[int] = None
) -> Tuple[str, bytes]:
    img = Image.open(io.BytesIO(data))
    img.load()

    effective_max = max_dim
    if dpi_limit and dpi_limit > 0:
        effective_max = min(effective_max, dpi_limit)

    img = _resize_if_needed(img, effective_max)
    alpha = _has_alpha(img)

    candidates: List[Tuple[str, bytes]] = [(name, data)]

    png_buf = io.BytesIO()
    if alpha:
        img.save(png_buf, "PNG", optimize=True)
    else:
        img.convert("RGB").save(png_buf, "PNG", optimize=True)
    candidates.append((str(Path(name).with_suffix(".png")), png_buf.getvalue()))

    if not alpha:
        jpg_buf = io.BytesIO()
        _to_rgb(img).save(jpg_buf, "JPEG", quality=quality, optimize=True, progressive=True)
        candidates.append((str(Path(name).with_suffix(".jpg")), jpg_buf.getvalue()))

    best_name, best_data = min(candidates, key=lambda it: (len(it[1]), 0 if it[0] == name else 1))

    if len(best_data) >= len(data):
        return name, data
    return best_name, best_data


def _pptx_make_media_placeholder(ext: str) -> bytes:
    ext = ext.lower()

    if HAS_PILLOW and ext in IMAGE_EXTS:
        rgb = Image.new("RGB", (1, 1), (255, 255, 255))
        rgba = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        buf = io.BytesIO()

        if ext in (".jpg", ".jpeg"):
            rgb.save(buf, "JPEG", quality=40, optimize=True)
        elif ext == ".png":
            rgba.save(buf, "PNG", optimize=True)
        elif ext == ".gif":
            gif = rgba.convert("P", palette=Image.ADAPTIVE)
            gif.info["transparency"] = 0
            gif.save(buf, "GIF", transparency=0)
        elif ext in (".bmp",):
            rgb.save(buf, "BMP")
        elif ext in (".tif", ".tiff"):
            rgba.save(buf, "TIFF", compression="tiff_lzw")
        elif ext == ".webp":
            rgba.save(buf, "WEBP", quality=20, method=6)
        else:
            rgba.save(buf, "PNG", optimize=True)

        return buf.getvalue()

    return b"0"


def _pptx_fix_content_types(all_entries: Dict[str, bytes], name_map: Dict[str, str]) -> None:
    ct_name = "[Content_Types].xml"
    if ct_name not in all_entries:
        return

    base_name_map = {Path(old).name: Path(new).name for old, new in name_map.items() if old != new}
    if not base_name_map:
        return

    ct_text = all_entries[ct_name].decode("utf-8", errors="replace")
    changed = False

    for old_base, new_base in base_name_map.items():
        if old_base in ct_text:
            ct_text = ct_text.replace(old_base, new_base)
            changed = True

    introduced_new_exts: set = set()
    for old, new in name_map.items():
        if old == new:
            continue
        old_e = Path(old).suffix.lower()
        new_e = Path(new).suffix.lower()
        if old_e != new_e:
            introduced_new_exts.add(new_e)

    for ext in introduced_new_exts:
        ext_bare = ext.lstrip(".")
        ct_for_ext = CONTENT_TYPES.get(ext)
        if not ct_for_ext:
            continue
        if f'Extension="{ext_bare}"' not in ct_text:
            tag = f'<Default Extension="{ext_bare}" ContentType="{ct_for_ext}"/>'
            ct_text = ct_text.replace("</Types>", f"  {tag}\n</Types>")
            changed = True

    if changed:
        all_entries[ct_name] = ct_text.encode("utf-8")


def _pptx_super_dry(all_entries: Dict[str, bytes]) -> None:
    name_map: Dict[str, str] = {}

    for name in list(all_entries.keys()):
        if not name.startswith("ppt/media/"):
            continue
        ext = Path(name).suffix.lower()
        if ext not in (IMAGE_EXTS | AUDIO_EXTS | VIDEO_EXTS):
            continue

        if ext in (".jpg", ".jpeg", ".bmp"):
            cand = str(Path(name).with_suffix(".png")).replace("\\", "/")
            if cand in all_entries and cand != name:
                p = Path(name)
                cand = str(p.with_name(p.stem + "_superdry.png")).replace("\\", "/")
            all_entries.pop(name, None)
            all_entries[cand] = _pptx_make_media_placeholder(".png")
            name_map[name] = cand
        else:
            all_entries[name] = _pptx_make_media_placeholder(ext)

    if name_map:
        base_name_map = {Path(old).name: Path(new).name for old, new in name_map.items() if old != new}
        for entry_name, data in list(all_entries.items()):
            if not (entry_name.endswith(".xml") or entry_name.endswith(".rels")):
                continue
            if entry_name == "[Content_Types].xml":
                continue
            try:
                text = data.decode("utf-8")
            except Exception:
                continue
            changed = False
            for old_base, new_base in base_name_map.items():
                if old_base in text:
                    text = text.replace(old_base, new_base)
                    changed = True
            if changed:
                all_entries[entry_name] = text.encode("utf-8")

        _pptx_fix_content_types(all_entries, name_map)


def compress_pptx_file(src: Path, dst: Path, preset: dict) -> bool:
    if not HAS_PILLOW:
        print(f"\n    ✗ 需要 Pillow 才能压缩 PPTX: {src.name}")
        return False

    quality = preset["doc_quality"]
    max_dim = preset["doc_max_dim"]
    max_dpi = preset.get("doc_max_dpi")

    try:
        with zipfile.ZipFile(src, "r") as zin:
            all_entries = {name: zin.read(name) for name in zin.namelist()}

        if preset.get("super_dry", False):
            _pptx_super_dry(all_entries)
            dst.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
                for name, data in all_entries.items():
                    zout.writestr(name, data)
            return True

        dpi_limit_map: Dict[str, int] = {}
        if max_dpi:
            dpi_limit_map = _pptx_build_dpi_limit_map(all_entries, max_dpi)

        name_map: Dict[str, str] = {}
        new_contents: Dict[str, bytes] = {}

        for name, data in all_entries.items():
            parts = name.split("/")
            in_media = len(parts) >= 2 and parts[-2] == "media"
            if not in_media:
                continue
            ext = Path(name).suffix.lower()

            if ext in VIDEO_EXTS and HAS_FFMPEG:
                try:
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
                        tmp_in = Path(tf.name)
                    tmp_out = tmp_in.with_name(tmp_in.stem + "_compressed.mp4")
                    tmp_in.write_bytes(data)

                    ok = _compress_video_in_memory(tmp_in, tmp_out, preset)
                    if ok and tmp_out.exists():
                        new_data = tmp_out.read_bytes()
                        if len(new_data) < len(data):
                            new_contents[name] = new_data
                            name_map[name] = name
                        tmp_out.unlink(missing_ok=True)
                    tmp_in.unlink(missing_ok=True)
                except Exception as e:
                    print(f"\n      跳过视频 [{name}]: {e}")
                continue

            if ext not in IMAGE_EXTS or ext in (".emf", ".wmf"):
                continue
            try:
                new_name, new_data = _pptx_compress_image_entry(
                    name, data, quality, max_dim, dpi_limit=dpi_limit_map.get(name)
                )
                new_contents[new_name] = new_data
                name_map[name] = new_name
            except Exception as e:
                print(f"\n      跳过图片 [{name}]: {e}")

        _pptx_fix_content_types(all_entries, name_map)

        base_name_map = {Path(old).name: Path(new).name for old, new in name_map.items() if old != new}

        updated_texts: Dict[str, bytes] = {}
        for name, data in all_entries.items():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            if name == "[Content_Types].xml":
                continue
            try:
                text = data.decode("utf-8")
                changed = False
                for old_base, new_base in base_name_map.items():
                    if old_base in text:
                        text = text.replace(old_base, new_base)
                        changed = True
                if changed:
                    updated_texts[name] = text.encode("utf-8")
            except Exception:
                pass

        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for name, data in all_entries.items():
                if name in name_map and name_map[name] != name:
                    continue
                if name in updated_texts:
                    zout.writestr(name, updated_texts[name])
                elif name in new_contents:
                    zout.writestr(name, new_contents[name])
                else:
                    zout.writestr(name, data)
            for new_name, new_data in new_contents.items():
                if new_name not in all_entries:
                    zout.writestr(new_name, new_data)

        return True

    except Exception as e:
        print(f"\n    ✗ PPTX 压缩失败 [{src.name}]: {e}")
        traceback.print_exc()
        return False
