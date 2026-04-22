#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
媒体文件有损压缩工具 v1.0
================================
支持处理的文件类型：
  独立文件：JPG/PNG/GIF/BMP/TIFF/WebP（图片）
           MP3/AAC/WAV/FLAC/OGG/M4A/WMA/OPUS（音频）
           MP4/MOV/AVI/MKV/WMV/FLV/WebM（视频）
  文档文件：PPTX（演示文稿）、PDF、PSD（Photoshop）、AI（Illustrator）

压缩策略（balanced 预设）：
  图片  → JPEG 75% 质量，超过 2560px 长边时等比缩放
  音频  → AAC 128 kbps
  视频  → H.264 CRF 23，最大 1920×1080，音频 128 kbps
  文档内图片 → JPEG 72%，最大 1920px（保持文档内尺寸不变，仅降低像素密度）

依赖安装：
  pip install Pillow python-pptx PyMuPDF pikepdf
  ffmpeg  (音视频压缩): https://ffmpeg.org/download.html
  ghostscript (可选, PDF): https://ghostscript.com/releases/gsdnld.html

用法示例：
  python compress_media.py 素材文件夹/
  python compress_media.py 素材文件夹/ -o 输出文件夹/
  python compress_media.py 素材文件夹/ --preset aggressive
  python compress_media.py 素材文件夹/ --inplace
  python compress_media.py 素材文件夹/ -q 80 --max-dim 1920
"""

import os
import io
import re
import sys
import shutil
import zipfile
import tempfile
import argparse
import subprocess
from pathlib import Path
from typing import Optional, Dict, List, Tuple

# ──────────────────────────────────────────────────────────────────
# 可选依赖检测
# ──────────────────────────────────────────────────────────────────
try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import pikepdf
    from pikepdf import PdfImage
    HAS_PIKEPDF = True
except ImportError:
    HAS_PIKEPDF = False

try:
    from pptx import Presentation  # noqa: F401 (仅用于测试导入)
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False


def _check_ffmpeg() -> bool:
    try:
        subprocess.run(["ffmpeg", "-version"], capture_output=True, timeout=10)
        return True
    except Exception:
        return False


def _check_ghostscript() -> Optional[str]:
    for cmd in ["gs", "gswin64c", "gswin32c"]:
        try:
            subprocess.run([cmd, "--version"], capture_output=True, timeout=10)
            return cmd
        except Exception:
            pass
    return None


HAS_FFMPEG = _check_ffmpeg()
GS_CMD = _check_ghostscript()

# ──────────────────────────────────────────────────────────────────
# 压缩预设
# ──────────────────────────────────────────────────────────────────
PRESETS: Dict[str, dict] = {
    # 均衡 —— 肉眼几乎感知不到质量损失，压缩率可观
    "balanced": {
        "image_quality": 75,
        "image_max_dim": 2560,
        "audio_codec": "aac",
        "audio_bitrate": "128k",
        "video_codec": "libx264",
        "video_crf": 23,
        "video_preset": "medium",
        "video_max_w": 1920,
        "video_max_h": 1080,
        "video_audio_bitrate": "128k",
        "doc_quality": 72,
        "doc_max_dim": 1920,
    },
    # 激进 —— 最大压缩，较小分辨率，少量可感知损失
    "aggressive": {
        "image_quality": 60,
        "image_max_dim": 1920,
        "audio_codec": "aac",
        "audio_bitrate": "96k",
        "video_codec": "libx264",
        "video_crf": 28,
        "video_preset": "medium",
        "video_max_w": 1280,
        "video_max_h": 720,
        "video_audio_bitrate": "96k",
        "doc_quality": 60,
        "doc_max_dim": 1280,
    },
    # 高质量 —— 质量优先，适合归档
    "high": {
        "image_quality": 85,
        "image_max_dim": 4096,
        "audio_codec": "aac",
        "audio_bitrate": "192k",
        "video_codec": "libx264",
        "video_crf": 20,
        "video_preset": "slow",
        "video_max_w": 3840,
        "video_max_h": 2160,
        "video_audio_bitrate": "192k",
        "doc_quality": 82,
        "doc_max_dim": 2560,
    },
}

# ──────────────────────────────────────────────────────────────────
# 文件扩展名分类
# ──────────────────────────────────────────────────────────────────
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
AUDIO_EXTS = {".mp3", ".aac", ".m4a", ".wav", ".flac", ".ogg", ".opus", ".wma"}
VIDEO_EXTS = {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".webm", ".m4v", ".3gp"}
DOC_EXTS   = {".pptx", ".ppt", ".pdf", ".psd", ".ai"}

# Content-Type 映射（用于 PPTX/DOCX 内部 XML 更新）
CONTENT_TYPES = {
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".gif":  "image/gif",
    ".bmp":  "image/bmp",
    ".tiff": "image/tiff",
    ".tif":  "image/tiff",
    ".webp": "image/webp",
}


# ──────────────────────────────────────────────────────────────────
# 统计
# ──────────────────────────────────────────────────────────────────
class Stats:
    def __init__(self):
        self.processed = 0
        self.skipped = 0
        self.errors = 0
        self.original_bytes = 0
        self.compressed_bytes = 0

    def add(self, orig: int, comp: int):
        self.processed += 1
        self.original_bytes += orig
        self.compressed_bytes += comp

    def report(self):
        saved = self.original_bytes - self.compressed_bytes
        ratio = (saved / self.original_bytes * 100) if self.original_bytes > 0 else 0.0
        print(f"\n{'=' * 52}")
        print(f"  处理完成")
        print(f"{'=' * 52}")
        print(f"  已压缩：{self.processed} 个文件")
        print(f"  已跳过：{self.skipped} 个文件（非媒体，直接复制）")
        print(f"  出错：  {self.errors} 个文件（已复制原件）")
        print(f"  原始大小：  {fmt_size(self.original_bytes)}")
        print(f"  压缩后大小：{fmt_size(self.compressed_bytes)}")
        print(f"  共节省：    {fmt_size(saved)}  ({ratio:.1f}%)")
        print()


def fmt_size(b: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if b < 1024:
            return f"{b:.1f} {unit}"
        b /= 1024
    return f"{b:.1f} TB"


# ──────────────────────────────────────────────────────────────────
# 图片压缩核心
# ──────────────────────────────────────────────────────────────────
def _resize_if_needed(img: "Image.Image", max_dim: int) -> "Image.Image":
    """等比缩放，保证长边不超过 max_dim；若已满足则原样返回。"""
    w, h = img.size
    if max(w, h) <= max_dim:
        return img
    scale = max_dim / max(w, h)
    return img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)


def _to_rgb(img: "Image.Image", bg: Tuple[int, int, int] = (255, 255, 255)) -> "Image.Image":
    """将任意模式转换为 RGB，透明区域填充白色背景。"""
    if img.mode == "P":
        img = img.convert("RGBA")
    if img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", img.size, bg)
        background.paste(img, mask=img.split()[-1])
        return background
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def _has_alpha(img: "Image.Image") -> bool:
    if img.mode in ("RGBA", "LA"):
        return True
    if img.mode == "P" and "transparency" in img.info:
        return True
    return False


def compress_image_to_jpeg_bytes(data: bytes, quality: int, max_dim: int) -> bytes:
    """
    将任意图片字节压缩为 JPEG 字节。
    透明区域用白色填充。
    """
    img = Image.open(io.BytesIO(data))
    img.load()
    img = _resize_if_needed(img, max_dim)
    img = _to_rgb(img)
    out = io.BytesIO()
    img.save(out, "JPEG", quality=quality, optimize=True, progressive=True)
    return out.getvalue()


def compress_image_to_png_bytes(data: bytes, max_dim: int) -> bytes:
    """压缩 PNG（保留透明通道），仅缩放，不改格式。"""
    img = Image.open(io.BytesIO(data))
    img.load()
    img = _resize_if_needed(img, max_dim)
    out = io.BytesIO()
    img.save(out, "PNG", optimize=True)
    return out.getvalue()


def compress_image_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    """
    压缩独立图片文件，输出为 JPEG。
    返回实际输出路径（.jpg），失败返回 None。
    """
    if not HAS_PILLOW:
        return None
    try:
        data = src.read_bytes()
        compressed = compress_image_to_jpeg_bytes(
            data, preset["image_quality"], preset["image_max_dim"]
        )
        # 只有压缩后更小才保存压缩版
        out_data = compressed if len(compressed) < len(data) else data
        out_path = dst.with_suffix(".jpg")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(out_data)
        return out_path
    except Exception as e:
        print(f"\n    ✗ 图片压缩失败 [{src.name}]: {e}")
        return None


# ──────────────────────────────────────────────────────────────────
# 音频压缩
# ──────────────────────────────────────────────────────────────────
def compress_audio_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    """使用 ffmpeg 将音频压缩为 AAC (.m4a)。"""
    if not HAS_FFMPEG:
        return None
    out_path = dst.with_suffix(".m4a")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    cmd = [
        "ffmpeg", "-y", "-i", str(src),
        "-c:a", preset["audio_codec"],
        "-b:a", preset["audio_bitrate"],
        "-ar", "44100",
        str(out_path),
    ]
    result = subprocess.run(cmd, capture_output=True, timeout=600)
    if result.returncode != 0:
        err = result.stderr.decode(errors="ignore")[-300:]
        print(f"\n    ✗ 音频压缩失败 [{src.name}]: {err}")
        return None
    # 若压缩后更大，保留原文件副本
    if out_path.stat().st_size >= src.stat().st_size:
        out_path.unlink()
        shutil.copy2(src, dst)
        return dst
    return out_path


# ──────────────────────────────────────────────────────────────────
# 视频压缩
# ──────────────────────────────────────────────────────────────────
def compress_video_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    """使用 ffmpeg 将视频压缩为 H.264 MP4。"""
    if not HAS_FFMPEG:
        return None
    out_path = dst.with_suffix(".mp4")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    max_w, max_h = preset["video_max_w"], preset["video_max_h"]
    # 缩放滤镜：保持比例不超过最大尺寸，宽高对齐为偶数
    scale_filter = (
        f"scale='if(gt(iw,{max_w}),{max_w},iw)':"
        f"'if(gt(ih,{max_h}),{max_h},ih)':"
        f"force_original_aspect_ratio=decrease,"
        f"scale=trunc(iw/2)*2:trunc(ih/2)*2"
    )
    cmd = [
        "ffmpeg", "-y", "-i", str(src),
        "-c:v", preset["video_codec"],
        "-crf", str(preset["video_crf"]),
        "-preset", preset["video_preset"],
        "-vf", scale_filter,
        "-c:a", "aac",
        "-b:a", preset["video_audio_bitrate"],
        "-movflags", "+faststart",
        str(out_path),
    ]
    result = subprocess.run(cmd, capture_output=True, timeout=7200)
    if result.returncode != 0:
        err = result.stderr.decode(errors="ignore")[-300:]
        print(f"\n    ✗ 视频压缩失败 [{src.name}]: {err}")
        return None
    if out_path.stat().st_size >= src.stat().st_size:
        out_path.unlink()
        shutil.copy2(src, dst)
        return dst
    return out_path


# ──────────────────────────────────────────────────────────────────
# PPTX 压缩（操作 ZIP 内的媒体文件）
# ──────────────────────────────────────────────────────────────────
def _pptx_compress_image_entry(
    name: str, data: bytes, quality: int, max_dim: int
) -> Tuple[str, bytes]:
    """
    对 PPTX/ZIP 内单张图片进行压缩。
    - 有透明通道 → 压缩 PNG，保留原文件名
    - 无透明通道 → 转换为 JPEG，文件名后缀改为 .jpg
    返回 (新文件名, 新字节数据)
    """
    img = Image.open(io.BytesIO(data))
    img.load()
    alpha = _has_alpha(img)

    if alpha:
        new_data = compress_image_to_png_bytes(data, max_dim)
        new_name = name
    else:
        new_data = compress_image_to_jpeg_bytes(data, quality, max_dim)
        stem = str(Path(name).with_suffix(""))
        new_name = stem + ".jpg"

    # 只在压缩后更小时替换
    if len(new_data) >= len(data):
        return name, data
    return new_name, new_data


def compress_pptx_file(src: Path, dst: Path, preset: dict) -> bool:
    """
    压缩 PPTX 文件内嵌的图片。
    通过操作 ZIP 内部实现，不依赖 python-pptx。
    """
    if not HAS_PILLOW:
        print(f"\n    ✗ 需要 Pillow 才能压缩 PPTX: {src.name}")
        return False

    quality  = preset["doc_quality"]
    max_dim  = preset["doc_max_dim"]

    try:
        # 一次性读入全部内容避免上下文问题
        with zipfile.ZipFile(src, "r") as zin:
            all_entries = {name: zin.read(name) for name in zin.namelist()}

        # 记录名称映射 old_name -> new_name
        name_map: Dict[str, str] = {}
        new_contents: Dict[str, bytes] = {}

        for name, data in all_entries.items():
            parts = name.split("/")
            in_media = len(parts) >= 2 and parts[-2] == "media"
            if not in_media:
                continue
            ext = Path(name).suffix.lower()
            if ext not in IMAGE_EXTS or ext in (".emf", ".wmf"):
                continue
            try:
                new_name, new_data = _pptx_compress_image_entry(
                    name, data, quality, max_dim
                )
                new_contents[new_name] = new_data
                name_map[name] = new_name
            except Exception as e:
                print(f"\n      跳过图片 [{name}]: {e}")

        # 更新所有 XML / rels 文件中的文件名引用
        updated_texts: Dict[str, bytes] = {}
        for name, data in all_entries.items():
            if not (name.endswith(".xml") or name.endswith(".rels")):
                continue
            try:
                text = data.decode("utf-8")
                changed = False
                for old, new in name_map.items():
                    if old == new:
                        continue
                    old_base = Path(old).name
                    new_base = Path(new).name
                    if old_base in text:
                        text = text.replace(old_base, new_base)
                        changed = True
                    # Content_Types.xml 中同时需要更新 ContentType 属性
                    if name == "[Content_Types].xml" and changed:
                        old_ct = CONTENT_TYPES.get(Path(old).suffix.lower(), "")
                        new_ct = CONTENT_TYPES.get(Path(new).suffix.lower(), "")
                        if old_ct and new_ct and old_ct != new_ct:
                            text = text.replace(
                                f'ContentType="{old_ct}"',
                                f'ContentType="{new_ct}"',
                            )
                if changed:
                    updated_texts[name] = text.encode("utf-8")
            except Exception:
                pass

        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for name, data in all_entries.items():
                if name in name_map and name_map[name] != name:
                    continue  # 旧文件名被新文件名替代，不写入
                if name in updated_texts:
                    zout.writestr(name, updated_texts[name])
                elif name in new_contents:
                    zout.writestr(name, new_contents[name])
                else:
                    zout.writestr(name, data)
            # 写入经过重命名的图片
            for new_name, new_data in new_contents.items():
                if new_name not in all_entries:  # 确实是改名后的新条目
                    zout.writestr(new_name, new_data)

        return True

    except Exception as e:
        import traceback
        print(f"\n    ✗ PPTX 压缩失败 [{src.name}]: {e}")
        traceback.print_exc()
        return False


# ──────────────────────────────────────────────────────────────────
# PDF 压缩
# ──────────────────────────────────────────────────────────────────
def compress_pdf_file(src: Path, dst: Path, preset: dict) -> bool:
    """按优先级尝试各种 PDF 压缩方案。"""
    dst.parent.mkdir(parents=True, exist_ok=True)

    if GS_CMD:
        return _compress_pdf_gs(src, dst, preset)

    if HAS_PIKEPDF and HAS_PILLOW:
        return _compress_pdf_pikepdf(src, dst, preset)

    if HAS_PYMUPDF:
        return _compress_pdf_pymupdf(src, dst)

    print(f"\n    ✗ 未找到 PDF 压缩工具（需要 ghostscript 或 pikepdf），跳过: {src.name}")
    return False


def _compress_pdf_gs(src: Path, dst: Path, preset: dict) -> bool:
    """Ghostscript 方案：稳定可靠，支持完整图片重压缩。"""
    q = preset["doc_quality"]
    if q >= 80:
        setting = "/printer"    # ~300 dpi
    elif q >= 65:
        setting = "/ebook"      # ~150 dpi
    else:
        setting = "/screen"     # ~72 dpi

    cmd = [
        GS_CMD,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={setting}",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        f"-sOutputFile={dst}",
        str(src),
    ]
    result = subprocess.run(cmd, capture_output=True, timeout=600)
    if result.returncode != 0:
        err = result.stderr.decode(errors="ignore")[-300:]
        print(f"\n    ✗ Ghostscript PDF 失败 [{src.name}]: {err}")
        return False
    return True


def _compress_pdf_pikepdf(src: Path, dst: Path, preset: dict) -> bool:
    """pikepdf 方案：逐页替换图片为 JPEG。"""
    quality = preset["doc_quality"]
    max_dim = preset["doc_max_dim"]
    try:
        pdf = pikepdf.open(str(src))
        for page in pdf.pages:
            _pdf_page_recompress_images(page, quality, max_dim)
        pdf.save(
            str(dst),
            compress_streams=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
        )
        return True
    except Exception as e:
        print(f"\n    ✗ pikepdf PDF 失败 [{src.name}]: {e}")
        return False


def _pdf_page_recompress_images(page, quality: int, max_dim: int):
    """对 PDF 页面内所有 /Image XObject 进行 JPEG 重压缩。"""
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
                pil_img = PdfImage(xobj).as_pil_image()
                pil_img = _resize_if_needed(pil_img, max_dim)
                pil_img = _to_rgb(pil_img)
                buf = io.BytesIO()
                pil_img.save(buf, "JPEG", quality=quality, optimize=True)
                jpeg_bytes = buf.getvalue()
                xobj.write(jpeg_bytes, filter=pikepdf.Name("/DCTDecode"))
                xobj["/Width"]            = pil_img.width
                xobj["/Height"]           = pil_img.height
                xobj["/ColorSpace"]       = (
                    pikepdf.Name("/DeviceRGB")
                    if pil_img.mode == "RGB"
                    else pikepdf.Name("/DeviceGray")
                )
                xobj["/BitsPerComponent"] = 8
                if "/DecodeParms" in xobj:
                    del xobj["/DecodeParms"]
            except Exception:
                pass  # 跳过无法处理的图片对象（矢量/蒙版等）
    except Exception:
        pass


def _compress_pdf_pymupdf(src: Path, dst: Path) -> bool:
    """PyMuPDF 方案：无法重压缩图片，但可清理空间/重建 xref。"""
    try:
        doc = fitz.open(str(src))
        doc.save(str(dst), garbage=4, deflate=True, clean=True)
        doc.close()
        return True
    except Exception as e:
        print(f"\n    ✗ PyMuPDF PDF 优化失败 [{src.name}]: {e}")
        return False


# ──────────────────────────────────────────────────────────────────
# PSD 压缩（合并图层后另存为 JPEG）
# ──────────────────────────────────────────────────────────────────
def compress_psd_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    """
    Pillow 能读取 PSD 并取得合并后的图像。
    由于 PSD 多图层信息无法无损保留在更小体积内，
    此处将合并图像另存为高质量 JPEG。
    """
    if not HAS_PILLOW:
        return None
    try:
        img = Image.open(src)
        img.load()
        img = _resize_if_needed(img, preset["image_max_dim"])
        img = _to_rgb(img)
        out_path = dst.with_suffix(".jpg")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(out_path, "JPEG", quality=preset["image_quality"], optimize=True)
        print(f"\n      注意: PSD 已合并图层 → JPEG ({src.name} → {out_path.name})")
        return out_path
    except Exception as e:
        print(f"\n    ✗ PSD 压缩失败 [{src.name}]: {e}")
        return None


# ──────────────────────────────────────────────────────────────────
# AI 压缩（现代 AI 文件基于 PDF，尝试作为 PDF 处理）
# ──────────────────────────────────────────────────────────────────
def compress_ai_file(src: Path, dst: Path, preset: dict) -> bool:
    """
    CS 版本以后的 AI 文件实质是带 PDF 兼容层的文件，
    尝试作为 PDF 压缩；若失败则保留原文件。
    """
    dst_pdf = dst.with_suffix(".pdf")
    if compress_pdf_file(src, dst_pdf, preset):
        print(f"\n      注意: AI 已作为 PDF 压缩 ({src.name} → {dst_pdf.name})")
        return True
    # 回退：作为普通文件复制
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    except Exception:
        pass
    print(f"\n      注意: AI 文件无法深度压缩，已复制原件 ({src.name})")
    return False


# ──────────────────────────────────────────────────────────────────
# 主处理逻辑
# ──────────────────────────────────────────────────────────────────
def process_file(
    src: Path, dst: Path, preset: dict, stats: Stats, verbose: bool
):
    """处理单个文件，根据类型分派到对应压缩函数。"""
    ext = src.suffix.lower()
    orig_size = src.stat().st_size

    if verbose:
        print(f"  {src.name}  ({fmt_size(orig_size)})", end="", flush=True)

    # 根据扩展名分派
    result_path: Optional[Path] = None
    success = False

    if ext in IMAGE_EXTS:
        result_path = compress_image_file(src, dst, preset)
        success = result_path is not None

    elif ext in AUDIO_EXTS:
        result_path = compress_audio_file(src, dst, preset)
        success = result_path is not None

    elif ext in VIDEO_EXTS:
        result_path = compress_video_file(src, dst, preset)
        success = result_path is not None

    elif ext in (".pptx",):
        success = compress_pptx_file(src, dst, preset)
        result_path = dst if success else None

    elif ext == ".ppt":
        # 旧版二进制 PPT —— 需要 LibreOffice 转换，暂不支持
        _fallback_copy(src, dst)
        if verbose:
            print(f"  → 跳过（旧版 .ppt 需 LibreOffice 转换，已复制原件）")
        stats.skipped += 1
        return

    elif ext == ".pdf":
        success = compress_pdf_file(src, dst, preset)
        result_path = dst if success else None

    elif ext == ".psd":
        result_path = compress_psd_file(src, dst, preset)
        success = result_path is not None

    elif ext == ".ai":
        success = compress_ai_file(src, dst, preset)
        result_path = dst.with_suffix(".pdf") if success else None

    else:
        # 非媒体文件直接复制
        _fallback_copy(src, dst)
        if verbose:
            print(f"  → 已复制（非媒体文件）")
        stats.skipped += 1
        return

    if success and result_path and result_path.exists():
        new_size = result_path.stat().st_size
        stats.add(orig_size, new_size)
        pct = (1.0 - new_size / orig_size) * 100 if orig_size > 0 else 0.0
        if verbose:
            print(f"  →  {fmt_size(new_size)}  (节省 {pct:.0f}%)")
    else:
        stats.errors += 1
        _fallback_copy(src, dst)
        if verbose:
            print(f"  → 压缩失败，已复制原件")


def _fallback_copy(src: Path, dst: Path):
    """安全地将文件复制到目标路径。"""
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def process_folder(
    src: Path, dst: Path, preset: dict, stats: Stats, verbose: bool
):
    """递归遍历文件夹处理所有文件。"""
    for item in sorted(src.iterdir()):
        if item.name.startswith("."):
            continue
        if item.is_symlink():
            continue
        if item.is_dir():
            process_folder(item, dst / item.name, preset, stats, verbose)
        elif item.is_file():
            process_file(item, dst / item.name, preset, stats, verbose)


def process_inplace(src: Path, preset: dict, stats: Stats, verbose: bool):
    """
    原地压缩：将每个文件压缩到临时文件后替换原文件（若变小）。
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_root = Path(tmpdir) / "work"
        tmp_root.mkdir()
        process_folder(src, tmp_root, preset, stats, verbose=False)
        # 将压缩后的文件替换回原处
        for tmp_file in tmp_root.rglob("*"):
            if not tmp_file.is_file():
                continue
            rel = tmp_file.relative_to(tmp_root)
            orig = src / rel.parent / rel.stem  # 可能扩展名不同
            # 尝试按原名找原始文件
            target = src / rel
            if not target.exists():
                # 可能改名了（如 png→jpg），找同 stem 的原文件
                candidates = list((src / rel.parent).glob(rel.stem + ".*"))
                target = candidates[0] if candidates else src / rel

            if target.exists() and tmp_file.stat().st_size < target.stat().st_size:
                shutil.copy2(tmp_file, target)
                if verbose:
                    print(f"  替换: {target.relative_to(src)}")
            elif not target.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(tmp_file, target)


# ──────────────────────────────────────────────────────────────────
# 命令行入口
# ──────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        prog="compress_media",
        description="媒体文件有损压缩工具 v1.0\n\n"
                    "对文件夹内所有图片/音频/视频及 PPTX/PDF/PSD/AI 中的嵌入媒体\n"
                    "进行有损压缩，在尽量不影响观感的前提下减小文件体积。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
预设说明：
  balanced   均衡（默认）  图片 75%/2560px  视频 CRF23/1080p  音频 128k
  aggressive 激进压缩      图片 60%/1920px  视频 CRF28/720p   音频 96k
  high       高质量归档    图片 85%/4096px  视频 CRF20/2160p  音频 192k

示例：
  python compress_media.py 素材/
  python compress_media.py 素材/ -o 素材_压缩/
  python compress_media.py 素材/ --preset aggressive
  python compress_media.py 素材/ --inplace
  python compress_media.py 素材/ -q 80 --max-dim 2048 --crf 25
        """,
    )
    parser.add_argument("input", help="输入文件夹路径")
    parser.add_argument("-o", "--output", help="输出文件夹（默认：<input>_compressed）")
    parser.add_argument(
        "-p", "--preset",
        choices=["balanced", "aggressive", "high"],
        default="balanced",
        help="压缩预设（默认: balanced）",
    )
    parser.add_argument("--inplace", action="store_true",
                        help="原地压缩，直接替换原文件（谨慎使用）")
    parser.add_argument("-q", "--quality", type=int, metavar="1-95",
                        help="覆盖图片 JPEG 质量（1-95）")
    parser.add_argument("--max-dim", type=int, metavar="PX",
                        help="覆盖图片/文档图片最大边长（像素）")
    parser.add_argument("--crf", type=int, metavar="0-51",
                        help="覆盖视频 CRF 值（越小越好，推荐 18-28）")
    parser.add_argument("--audio-bitrate", metavar="RATE",
                        help="覆盖音频码率（如 128k、192k）")
    parser.add_argument("--quiet", action="store_true",
                        help="静默模式，不逐文件打印进度")

    args = parser.parse_args()
    verbose = not args.quiet

    # ── 验证输入
    src = Path(args.input).resolve()
    if not src.exists() or not src.is_dir():
        print(f"错误：输入路径不存在或不是文件夹：{src}")
        sys.exit(1)

    # ── 建立预设
    preset = dict(PRESETS[args.preset])
    if args.quality:
        preset["image_quality"] = max(1, min(95, args.quality))
        preset["doc_quality"]   = max(1, min(95, args.quality))
    if args.max_dim:
        preset["image_max_dim"] = args.max_dim
        preset["doc_max_dim"]   = args.max_dim
    if args.crf:
        preset["video_crf"] = max(0, min(51, args.crf))
    if args.audio_bitrate:
        preset["audio_bitrate"]       = args.audio_bitrate
        preset["video_audio_bitrate"] = args.audio_bitrate

    # ── 打印配置
    if verbose:
        print()
        print("  媒体文件有损压缩工具 v1.0")
        print("  " + "─" * 48)
        print(f"  输入目录：  {src}")
        if not args.inplace:
            dst_show = args.output if args.output else str(src.parent / (src.name + "_compressed"))
            print(f"  输出目录：  {dst_show}")
        else:
            print(f"  模式：      原地压缩（inplace）")
        print(f"  预设：      {args.preset}")
        print(f"  图片质量：  {preset['image_quality']}%  最大边长：{preset['image_max_dim']}px")
        print(f"  视频 CRF：  {preset['video_crf']}  音频码率：{preset['audio_bitrate']}")
        print()
        print("  依赖状态：")
        print(f"    Pillow      {'✓' if HAS_PILLOW else '✗  pip install Pillow'}")
        print(f"    PyMuPDF     {'✓' if HAS_PYMUPDF else '✗  pip install PyMuPDF'}")
        print(f"    pikepdf     {'✓' if HAS_PIKEPDF else '✗  pip install pikepdf'}")
        print(f"    python-pptx {'✓' if HAS_PPTX else '✗  pip install python-pptx'}")
        print(f"    ffmpeg      {'✓' if HAS_FFMPEG else '✗  https://ffmpeg.org/download.html'}")
        print(f"    ghostscript {'✓ (' + GS_CMD + ')' if GS_CMD else '✗  可选，用于 PDF（https://ghostscript.com）'}")
        print()
        print("  " + "─" * 48)
        print("  开始处理...")
        print()

    stats = Stats()

    if args.inplace:
        process_inplace(src, preset, stats, verbose)
    else:
        if args.output:
            dst = Path(args.output).resolve()
        else:
            dst = src.parent / (src.name + "_compressed")
        dst.mkdir(parents=True, exist_ok=True)
        process_folder(src, dst, preset, stats, verbose)

    if verbose:
        stats.report()


if __name__ == "__main__":
    main()
