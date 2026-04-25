import io
import shutil
import subprocess
import tempfile
import traceback
import zipfile
from pathlib import Path
from typing import Dict, Optional

from ..constants import AUDIO_EXTS, IMAGE_EXTS, VIDEO_EXTS
from ..deps import HAS_FFMPEG, HAS_PILLOW, Image
from ..utils.image_utils import _resize_if_needed


def _iter_media_entries(entries: Dict[str, bytes]):
    for name, data in entries.items():
        if name.endswith("/"):
            continue
        ext = Path(name).suffix.lower()
        if ext not in (IMAGE_EXTS | AUDIO_EXTS | VIDEO_EXTS):
            continue
        parts = Path(name).parts
        if parts and parts[0] in {"Data", "Metadata"}:
            yield name, ext, data


def _compress_key_image_bytes(data: bytes, ext: str, preset: dict) -> Optional[bytes]:
    if not HAS_PILLOW:
        return None

    img = Image.open(io.BytesIO(data))
    img.load()

    # 使用文件实际格式（magic bytes 决定），而非扩展名
    # 避免 PNG 存为 .jpg 或 JPEG 存为 .png 时格式被错误转换
    actual_format = (img.format or "").upper()
    if actual_format == "JPEG":
        use_ext = ".jpg"
    elif actual_format == "PNG":
        use_ext = ".png"
    elif actual_format in ("TIFF", "TIFF"):
        use_ext = ".tiff"
    elif actual_format == "GIF":
        use_ext = ".gif"
    elif actual_format == "BMP":
        use_ext = ".bmp"
    elif actual_format == "WEBP":
        use_ext = ".webp"
    else:
        use_ext = ext  # 回退到扩展名

    img = _resize_if_needed(img, preset["doc_max_dim"])

    out = io.BytesIO()
    if use_ext in {".jpg", ".jpeg"}:
        img = img.convert("RGB")
        img.save(out, "JPEG", quality=preset["doc_quality"], optimize=True, progressive=True)
    elif use_ext == ".png":
        img.save(out, "PNG", optimize=True)
    elif use_ext == ".gif":
        img.save(out, "GIF", optimize=True)
    elif use_ext == ".bmp":
        img = img.convert("RGB")
        img.save(out, "BMP")
    elif use_ext in {".tif", ".tiff"}:
        img.save(out, "TIFF", compression="tiff_lzw")
    elif use_ext == ".webp":
        img.save(out, "WEBP", quality=preset["doc_quality"], method=6)
    else:
        return None

    new_data = out.getvalue()
    return new_data if len(new_data) < len(data) else None


def _compress_media_same_ext_bytes(data: bytes, ext: str, preset: dict, media_kind: str) -> Optional[bytes]:
    if not HAS_FFMPEG:
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            tmp_in = tmp / f"in{ext}"
            tmp_out = tmp / f"out{ext}"
            tmp_in.write_bytes(data)

            if media_kind == "audio":
                cmd = [
                    "ffmpeg", "-y", "-i", str(tmp_in),
                    "-c:a", preset["audio_codec"],
                    "-b:a", preset["audio_bitrate"],
                    "-ar", "44100",
                    "-ac", "1",
                    str(tmp_out),
                ]
            else:
                max_dim = int(preset.get("image_max_dim", 1920))
                scale_filter = (
                    f"scale='if(gt(iw,ih),if(gt(iw,{max_dim}),{max_dim},iw),-2)':"
                    f"'if(gt(ih,iw),if(gt(ih,{max_dim}),{max_dim},ih),-2)':"
                    "force_original_aspect_ratio=decrease,"
                    "scale=trunc(iw/2)*2:trunc(ih/2)*2,fps=24"
                )
                cmd = [
                    "ffmpeg", "-y", "-i", str(tmp_in),
                    "-c:v", preset["video_codec"],
                    "-crf", str(preset["video_crf"]),
                    "-preset", preset["video_preset"],
                    "-vf", scale_filter,
                    "-c:a", "aac",
                    "-b:a", preset["video_audio_bitrate"],
                    "-ar", "44100",
                    "-ac", "1",
                ]
                if preset.get("video_codec") == "libx265":
                    cmd += ["-tag:v", "hvc1", "-x265-params", "log-level=error"]
                cmd.append(str(tmp_out))

            result = subprocess.run(cmd, capture_output=True, timeout=3600)
            if result.returncode != 0 or not tmp_out.exists():
                return None

            new_data = tmp_out.read_bytes()
            return new_data if len(new_data) < len(data) else None
    except Exception:
        return None


def _make_placeholder(ext: str) -> bytes:
    ext = ext.lower()

    if HAS_PILLOW and ext in IMAGE_EXTS:
        rgb = Image.new("RGB", (1, 1), (255, 255, 255))
        rgba = Image.new("RGBA", (1, 1), (0, 0, 0, 0))
        out = io.BytesIO()

        if ext in {".jpg", ".jpeg"}:
            rgb.save(out, "JPEG", quality=35, optimize=True)
        elif ext == ".png":
            rgba.save(out, "PNG", optimize=True)
        elif ext == ".gif":
            gif = rgba.convert("P", palette=Image.ADAPTIVE)
            gif.info["transparency"] = 0
            gif.save(out, "GIF", transparency=0)
        elif ext == ".bmp":
            rgb.save(out, "BMP")
        elif ext in {".tif", ".tiff"}:
            rgba.save(out, "TIFF", compression="tiff_lzw")
        elif ext == ".webp":
            rgba.save(out, "WEBP", quality=20, method=6)
        else:
            rgba.save(out, "PNG", optimize=True)
        return out.getvalue()

    return b"0"


def _process_keynote_entries(entries: Dict[str, bytes], preset: dict) -> Dict[str, bytes]:
    out_entries = dict(entries)

    for name, ext, data in _iter_media_entries(entries):
        try:
            if preset.get("super_dry", False):
                out_entries[name] = _make_placeholder(ext)
                continue

            if ext in IMAGE_EXTS:
                new_data = _compress_key_image_bytes(data, ext, preset)
                if new_data is not None:
                    out_entries[name] = new_data
                continue

            if ext in AUDIO_EXTS:
                new_data = _compress_media_same_ext_bytes(data, ext, preset, media_kind="audio")
                if new_data is not None:
                    out_entries[name] = new_data
                continue

            if ext in VIDEO_EXTS:
                new_data = _compress_media_same_ext_bytes(data, ext, preset, media_kind="video")
                if new_data is not None:
                    out_entries[name] = new_data
        except Exception as e:
            print(f"\n      跳过 Keynote 媒体 [{name}]: {e}")

    return out_entries


def _fix_zip_filename(info: zipfile.ZipInfo) -> str:
    """
    Keynote 创建的 ZIP 文件中，部分中文文件名以 UTF-8 字节存储，
    但未设 UTF-8 flag（flag_bits=0）。Python zipfile 会将这些字节
    按 CP437 解码成乱码字符串。
    写回时 Python 会将乱码字符串重新按 UTF-8 编码，导致文件名字节
    完全不同，Keynote 无法找到对应文件，报"已损坏"。
    本函数将 CP437 乱码还原为正确的 UTF-8 文件名。
    """
    if info.flag_bits & 0x800:
        # 已有 UTF-8 flag，Python 已正确解码，无需修复
        return info.filename
    if not any(ord(c) > 127 for c in info.filename):
        # 纯 ASCII，无问题
        return info.filename
    try:
        # 将 CP437 解码结果重新编码为 CP437 字节，再按 UTF-8 解码
        return info.filename.encode("cp437").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return info.filename


def _compress_keynote_zip(src: Path, dst: Path, preset: dict) -> bool:
    with zipfile.ZipFile(src, "r") as zin:
        infolist = zin.infolist()
        # 修复文件名编码，同时保留 orig→fixed 的映射用于读取数据
        name_map = [(info.filename, _fix_zip_filename(info)) for info in infolist]
        entries = {}
        for orig, fixed in name_map:
            entries[fixed] = zin.read(orig) if not orig.endswith("/") else b""

    new_entries = _process_keynote_entries(entries, preset)

    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_STORED) as zout:
        for _orig, fixed in name_map:
            if fixed.endswith("/"):
                zout.writestr(fixed, b"")
            else:
                zout.writestr(fixed, new_entries.get(fixed, entries[fixed]))
    return True


def _compress_keynote_dir(src: Path, dst: Path, preset: dict) -> bool:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)

    file_map: Dict[str, bytes] = {}
    for f in dst.rglob("*"):
        if not f.is_file():
            continue
        rel = f.relative_to(dst).as_posix()
        file_map[rel] = f.read_bytes()

    new_entries = _process_keynote_entries(file_map, preset)

    for rel, data in new_entries.items():
        target = dst / Path(rel)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(data)

    return True


def compress_keynote_file(src: Path, dst: Path, preset: dict) -> bool:
    try:
        if src.is_dir():
            return _compress_keynote_dir(src, dst, preset)

        if not src.is_file():
            return False

        if zipfile.is_zipfile(src):
            return _compress_keynote_zip(src, dst, preset)

        return False
    except Exception as e:
        print(f"\n    ✗ Keynote 压缩失败 [{src.name}]: {e}")
        traceback.print_exc()
        return False
