import io
import re
import shutil
import subprocess
from pathlib import Path

from ..deps import GS_CMD, HAS_PIKEPDF, HAS_PILLOW, HAS_PYMUPDF, PdfImage, fitz, pikepdf
from ..utils.image_utils import _resize_if_needed, _to_rgb


def _pdf_super_dry(src: Path, dst: Path) -> bool:
    if not HAS_PIKEPDF:
        print(f"\n    ✗ super_dry 处理 PDF 需要 pikepdf: {src.name}")
        return False

    tmp_path = dst.with_suffix(".superdry.tmp.pdf")

    try:
        pdf = pikepdf.open(str(src))

        for page in pdf.pages:
            resources = page.get("/Resources")
            if not resources:
                continue

            xobj = resources.get("/XObject")
            if not xobj:
                continue

            image_names = []
            for key in list(xobj.keys()):
                try:
                    obj = xobj[key]
                    if obj.get("/Subtype") == pikepdf.Name("/Image"):
                        image_names.append(str(key))
                except Exception:
                    continue

            if not image_names:
                continue

            for key in list(xobj.keys()):
                if str(key) in image_names:
                    try:
                        del xobj[key]
                    except Exception:
                        pass

            try:
                contents = page.get("/Contents")
                if contents is not None:
                    streams = []
                    if isinstance(contents, pikepdf.Array):
                        streams = list(contents)
                    else:
                        streams = [contents]

                    for st in streams:
                        try:
                            raw = bytes(st.read_bytes())
                            txt = raw.decode("latin-1", errors="ignore")
                            for nm in image_names:
                                txt = re.sub(r"/" + re.escape(nm.lstrip("/")) + r"\s+Do", "", txt)
                            st.write(txt.encode("latin-1"))
                        except Exception:
                            continue
            except Exception:
                pass

        pdf.save(
            str(tmp_path),
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
        )
        pdf.close()

        if HAS_PYMUPDF:
            doc = fitz.open(str(tmp_path))
            doc.save(str(dst), garbage=4, deflate=True, clean=True)
            doc.close()
            tmp_path.unlink(missing_ok=True)
        else:
            shutil.move(str(tmp_path), str(dst))

        return True
    except Exception as e:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        print(f"\n    ✗ PDF super_dry 失败 [{src.name}]: {e}")
        return False


def _compress_pdf_gs(src: Path, dst: Path, preset: dict) -> bool:
    q = preset["doc_quality"]
    max_dpi = preset.get("doc_max_dpi")

    if q >= 80:
        setting = "/printer"
    elif q >= 65:
        setting = "/ebook"
    else:
        setting = "/screen"

    cmd = [
        GS_CMD,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={setting}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
    ]

    if max_dpi:
        cmd += [
            f"-dColorImageResolution={max_dpi}",
            f"-dGrayImageResolution={max_dpi}",
            f"-dMonoImageResolution={min(max_dpi * 2, 1200)}",
            "-dDownsampleColorImages=true",
            "-dDownsampleGrayImages=true",
            "-dDownsampleMonoImages=true",
        ]

    cmd += [f"-sOutputFile={dst}", str(src)]
    result = subprocess.run(cmd, capture_output=True, timeout=600)
    if result.returncode != 0:
        err = result.stderr.decode(errors="ignore")[-300:]
        print(f"\n    ✗ Ghostscript PDF 失败 [{src.name}]: {err}")
        return False
    return True


def _pdf_page_dpi_max_dim(page, max_dpi: int) -> int:
    try:
        mb = page.get("/MediaBox")
        if mb and len(mb) >= 4:
            pt_w = float(mb[2]) - float(mb[0])
            pt_h = float(mb[3]) - float(mb[1])
            return int(max(pt_w, pt_h) / 72 * max_dpi)
    except Exception:
        pass
    return int(11.69 * max_dpi)


def _pdf_page_recompress_images(page, quality: int, max_dim: int):
    try:
        resources = page.get("/Resources")
        if not resources:
            return
        xobjects = resources.get("/XObject")
        if not xobjects:
            return
        for key in list(xobjects.keys()):
            xobj = xobjects[key]
            try:
                if xobj.get("/Subtype") != pikepdf.Name("/Image"):
                    continue

                raw_orig_size = len(bytes(xobj.read_raw_bytes()))
                filt = str(xobj.get("/Filter"))

                pil_img = PdfImage(xobj).as_pil_image()
                orig_w, orig_h = pil_img.size
                pil_img = _resize_if_needed(pil_img, max_dim)
                was_resized = pil_img.size != (orig_w, orig_h)

                if not was_resized and (
                    "/JPXDecode" in filt or ("/FlateDecode" in filt and "/DCTDecode" in filt)
                ):
                    continue

                bpc = int(xobj.get("/BitsPerComponent", 8))
                if not was_resized and (bpc <= 1 or max(orig_w, orig_h) <= 256):
                    continue

                pil_img = _to_rgb(pil_img)
                buf = io.BytesIO()
                pil_img.save(buf, "JPEG", quality=quality, optimize=True)
                jpeg_bytes = buf.getvalue()

                if len(jpeg_bytes) >= raw_orig_size:
                    continue

                if was_resized and len(jpeg_bytes) > raw_orig_size * 0.95:
                    continue

                if not was_resized and len(jpeg_bytes) > raw_orig_size * 0.90:
                    continue

                xobj.write(jpeg_bytes, filter=pikepdf.Name("/DCTDecode"))
                xobj["/Width"] = pil_img.width
                xobj["/Height"] = pil_img.height
                xobj["/ColorSpace"] = (
                    pikepdf.Name("/DeviceRGB") if pil_img.mode == "RGB" else pikepdf.Name("/DeviceGray")
                )
                xobj["/BitsPerComponent"] = 8
                if "/DecodeParms" in xobj:
                    del xobj["/DecodeParms"]
            except Exception:
                pass
    except Exception:
        pass


def _compress_pdf_pikepdf(src: Path, dst: Path, preset: dict) -> bool:
    quality = preset["doc_quality"]
    max_dim = preset["doc_max_dim"]
    max_dpi = preset.get("doc_max_dpi")
    try:
        pdf = pikepdf.open(str(src))
        for page in pdf.pages:
            effective_max = max_dim
            if max_dpi:
                dpi_dim = _pdf_page_dpi_max_dim(page, max_dpi)
                effective_max = min(max_dim, dpi_dim)
            _pdf_page_recompress_images(page, quality, effective_max)
        pdf.save(
            str(dst),
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
        )
        return True
    except Exception as e:
        print(f"\n    ✗ pikepdf PDF 失败 [{src.name}]: {e}")
        return False


def _compress_pdf_pymupdf(src: Path, dst: Path) -> bool:
    try:
        doc = fitz.open(str(src))
        doc.save(str(dst), garbage=4, deflate=True, clean=True)
        doc.close()
        if dst.exists() and dst.stat().st_size >= src.stat().st_size:
            return False
        return True
    except Exception as e:
        print(f"\n    ✗ PyMuPDF PDF 优化失败 [{src.name}]: {e}")
        return False


def compress_pdf_file(src: Path, dst: Path, preset: dict) -> bool:
    dst.parent.mkdir(parents=True, exist_ok=True)
    src_size = src.stat().st_size

    if preset.get("super_dry", False):
        ok = _pdf_super_dry(src, dst)
        return ok

    if GS_CMD:
        ok = _compress_pdf_gs(src, dst, preset)
    elif HAS_PIKEPDF and HAS_PILLOW:
        ok = _compress_pdf_pikepdf(src, dst, preset)
    elif HAS_PYMUPDF:
        ok = _compress_pdf_pymupdf(src, dst)
    else:
        print(f"\n    ✗ 未找到 PDF 压缩工具（需要 ghostscript 或 pikepdf），跳过: {src.name}")
        return False

    if ok and dst.exists() and dst.stat().st_size >= src_size:
        shutil.copy2(src, dst)

    return ok
