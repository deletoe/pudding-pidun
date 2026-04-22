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
import xml.etree.ElementTree as ET
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
    # 均衡 —— 极限体积优先（视频 x265 + 降帧 + 低码率音频）
    "balanced": {
        "super_dry": False,
        "image_quality": 75,
        "image_max_dim": 2560,
        "image_max_dpi": 150,       # 独立图片：超过此 DPI 则按比例缩小像素
        "audio_codec": "aac",
        "audio_bitrate": "8k",
        "video_codec": "libx265",
        "video_crf": 34,
        "video_preset": "medium",
        "video_audio_bitrate": "8k",
        "doc_quality": 72,
        "doc_max_dim": 1920,
        "doc_max_dpi": 150,         # 文档内图片：依据显示尺寸限制像素密度
    },
    # 激进 —— 最大压缩，画质次于 balanced
    "aggressive": {
        "super_dry": False,
        "image_quality": 60,
        "image_max_dim": 1920,
        "image_max_dpi": 96,
        "audio_codec": "aac",
        "audio_bitrate": "8k",
        "video_codec": "libx265",
        "video_crf": 38,
        "video_preset": "medium",
        "video_audio_bitrate": "8k",
        "doc_quality": 60,
        "doc_max_dim": 1280,
        "doc_max_dpi": 96,
    },
    # 高质量 —— 仍保持 x265，但给更低 CRF
    "high": {
        "super_dry": False,
        "image_quality": 85,
        "image_max_dim": 4096,
        "image_max_dpi": 200,
        "audio_codec": "aac",
        "audio_bitrate": "8k",
        "video_codec": "libx265",
        "video_crf": 30,
        "video_preset": "slow",
        "video_audio_bitrate": "8k",
        "doc_quality": 82,
        "doc_max_dim": 2560,
        "doc_max_dpi": 200,
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


def _etag(el: ET.Element) -> str:
    """返回 XML 标签的本地名（去命名空间前缀）。"""
    tag = el.tag
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _et_serialize(root: ET.Element) -> bytes:
    """统一 XML 输出格式。"""
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


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
    若图片包含 DPI 元数据且超过 image_max_dpi，则先按 DPI 比例缩小像素数，
    再结合 image_max_dim 取更严格的约束。
    返回实际输出路径（.jpg），失败返回 None。
    """
    if not HAS_PILLOW:
        return None
    try:
        data     = src.read_bytes()
        max_dim  = preset["image_max_dim"]
        max_dpi  = preset.get("image_max_dpi")

        if max_dpi:
            # 读取图片嵌入的 DPI 元数据
            img = Image.open(io.BytesIO(data))
            img.load()
            dpi_info = img.info.get("dpi")
            img_dpi: Optional[float] = None
            if isinstance(dpi_info, (tuple, list)) and len(dpi_info) >= 2:
                d0, d1 = float(dpi_info[0]), float(dpi_info[1])
                if d0 > 0 and d1 > 0:
                    img_dpi = max(d0, d1)
            elif isinstance(dpi_info, (int, float)) and float(dpi_info) > 0:
                img_dpi = float(dpi_info)

            if img_dpi and img_dpi > max_dpi:
                # 按 DPI 缩放：新像素长边 = 原像素长边 × (max_dpi / 当前dpi)
                scale     = max_dpi / img_dpi
                dpi_limit = int(max(img.width, img.height) * scale)
                max_dim   = min(max_dim, dpi_limit)

        compressed = compress_image_to_jpeg_bytes(
            data, preset["image_quality"], max_dim
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
        "-ar", "16000",
        "-ac", "1",
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
def _probe_video_fps(src: Path) -> Optional[float]:
    """用 ffprobe 获取原视频帧率（浮点）。"""
    try:
        cmd = [
            "ffprobe", "-v", "error",
            "-select_streams", "v:0",
            "-show_entries", "stream=r_frame_rate",
            "-of", "default=noprint_wrappers=1:nokey=1",
            str(src),
        ]
        r = subprocess.run(cmd, capture_output=True, timeout=30)
        if r.returncode != 0:
            return None
        txt = r.stdout.decode(errors="ignore").strip()
        if not txt:
            return None
        if "/" in txt:
            a, b = txt.split("/", 1)
            b = float(b)
            if b == 0:
                return None
            return float(a) / b
        return float(txt)
    except Exception:
        return None


def _build_video_scale_filter(max_dim: int) -> str:
    """
    统一视频缩放规则（与图片一致）：长边不超过 max_dim，保持比例并对齐偶数。
    使用 if(gt()) 避免 ffmpeg 表达式里 min() 在部分构建上的兼容问题。
    """
    return (
        f"scale='if(gt(iw,ih),if(gt(iw,{max_dim}),{max_dim},iw),-2)':"
        f"'if(gt(ih,iw),if(gt(ih,{max_dim}),{max_dim},ih),-2)':"
        f"force_original_aspect_ratio=decrease,"
        f"scale=trunc(iw/2)*2:trunc(ih/2)*2"
    )


def _compress_video_in_memory(src: Path, dst: Path, preset: dict) -> bool:
    """
    给 compress_pptx_file 用的内存内视频压缩。
    与 compress_video_file 类似，但返回 bool 而不是 Optional[Path]。
    若压缩后更大则返回 False（不替换）。
    """
    max_dim = int(preset.get("image_max_dim", 1920))
    scale_filter = _build_video_scale_filter(max_dim)

    src_fps = _probe_video_fps(src)
    target_fps = 24.0 if not src_fps else min(24.0, src_fps)

    cmd = [
        "ffmpeg", "-y", "-i", str(src),
        "-c:v", preset["video_codec"],
        "-crf", str(preset["video_crf"]),
        "-preset", preset["video_preset"],
        "-vf", f"{scale_filter},fps={target_fps:.3f}",
        "-c:a", "aac",
        "-b:a", preset["video_audio_bitrate"],
        "-ar", "16000",
        "-movflags", "+faststart",
    ]
    if preset.get("video_codec") == "libx265":
        cmd += ["-tag:v", "hvc1", "-x265-params", "log-level=error"]
    cmd.append(str(dst))

    result = subprocess.run(cmd, capture_output=True, timeout=7200)
    if result.returncode != 0:
        return False
    # 若压缩后更大则视为失败
    if dst.stat().st_size >= src.stat().st_size:
        dst.unlink(missing_ok=True)
        return False
    return True


def compress_video_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    """使用 ffmpeg 将视频压缩为 x265 MP4，降帧到 min(24, 原始帧率)。"""
    if not HAS_FFMPEG:
        return None
    out_path = dst.with_suffix(".mp4")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # 视频分辨率规则与图片一致：使用 image_max_dim 作为长边上限
    max_dim = int(preset.get("image_max_dim", 1920))
    scale_filter = _build_video_scale_filter(max_dim)

    src_fps = _probe_video_fps(src)
    target_fps = 24.0 if not src_fps else min(24.0, src_fps)

    cmd = [
        "ffmpeg", "-y", "-i", str(src),
        "-c:v", preset["video_codec"],
        "-crf", str(preset["video_crf"]),
        "-preset", preset["video_preset"],
        "-vf", f"{scale_filter},fps={target_fps:.3f}",
        "-c:a", "aac",
        "-b:a", preset["video_audio_bitrate"],
        "-ar", "16000",
        "-movflags", "+faststart",
    ]
    if preset.get("video_codec") == "libx265":
        cmd += ["-tag:v", "hvc1", "-x265-params", "log-level=error"]
    cmd.append(str(out_path))
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
def _pptx_build_dpi_limit_map(all_entries: dict, max_dpi: int) -> Dict[str, int]:
    """
    解析 PPTX/ZIP 内的幻灯片 XML 和关系文件，建立映射：
        { 媒体文件全路径 : 该图片在幻灯片中显示时对应 max_dpi 的像素长边上限 }

    原理：
      幻灯片 XML 里每个图片形状都有 <a:ext cx="..." cy="...">（单位：EMU）。
      EMU → 英寸 → × max_dpi = 像素上限。
      这样压缩后的像素数恰好满足"显示尺寸不变、像素密度不超过 max_dpi"。

    914400 EMU = 1 英寸。
    """
    from posixpath import normpath, join, dirname

    EMU_PER_INCH = 914400
    # r:embed 的完整命名空间形式
    R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

    limit_map: Dict[str, int] = {}

    for name in all_entries:
        # 只处理幻灯片的 rels 文件
        # 格式：ppt/slides/_rels/slide1.xml.rels
        if "/_rels/" not in name or not name.endswith(".rels"):
            continue
        if "/slides/_rels/" not in name:
            continue

        try:
            rels_root = ET.fromstring(all_entries[name])
        except ET.ParseError:
            continue

        # 幻灯片目录，如 ppt/slides
        slide_dir = dirname(name.replace("/_rels", ""))

        # 建立 rId → 媒体完整路径
        rid_to_media: Dict[str, str] = {}
        for rel in rels_root:
            if "image" not in rel.get("Type", ""):
                continue
            rid    = rel.get("Id", "")
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

        # 对应的幻灯片 XML：ppt/slides/_rels/slide1.xml.rels → ppt/slides/slide1.xml
        slide_xml = name.replace("/_rels/", "/").removesuffix(".rels")
        if slide_xml not in all_entries:
            continue

        try:
            slide_root = ET.fromstring(all_entries[slide_xml])
        except ET.ParseError:
            continue

        # ElementTree 不支持 .parent，手动建立父节点映射
        parent_map = {
            child: parent
            for parent in slide_root.iter()
            for child in parent
        }

        for elem in slide_root.iter():
            local = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if local != "blip":
                continue

            rid = elem.get(R_EMBED)
            if not rid or rid not in rid_to_media:
                continue

            media_path = rid_to_media[rid]

            # 向上最多 10 层寻找 <a:ext cx="..." cy="...">
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

            # EMU → 英寸 → 像素上限
            limit_px = int(max(cx, cy) / EMU_PER_INCH * max_dpi)
            # 保留该图片跨所有幻灯片出现的最大值
            limit_map[media_path] = max(limit_map.get(media_path, 0), limit_px)

    return limit_map


def _pptx_compress_image_entry(
    name: str, data: bytes, quality: int, max_dim: int,
    dpi_limit: Optional[int] = None,
) -> Tuple[str, bytes]:
    """
    对 PPTX/ZIP 内单张图片进行压缩（格式自适应）。

    目标：在 PPT 可稳定支持的格式内（JPEG/PNG）选体积最小方案，
    并确保不比原图更大。

    策略：
      1) 先按 max_dim 与 dpi_limit 做等比缩放
      2) 生成 JPEG 与 PNG 两个候选（带透明通道时仅 PNG）
      3) 与原始数据比较，选最小者

    返回 (新文件名, 新字节数据)
    """
    img = Image.open(io.BytesIO(data))
    img.load()

    # 综合 max_dim 与 DPI 限制，取更严格的约束
    effective_max = max_dim
    if dpi_limit and dpi_limit > 0:
        effective_max = min(effective_max, dpi_limit)

    img = _resize_if_needed(img, effective_max)
    alpha = _has_alpha(img)

    candidates: List[Tuple[str, bytes]] = [(name, data)]

    # PNG 候选（兼容透明）
    png_buf = io.BytesIO()
    if alpha:
        img.save(png_buf, "PNG", optimize=True)
    else:
        img.convert("RGB").save(png_buf, "PNG", optimize=True)
    candidates.append((str(Path(name).with_suffix(".png")), png_buf.getvalue()))

    # JPEG 候选（无透明时）
    if not alpha:
        jpg_buf = io.BytesIO()
        img_rgb = _to_rgb(img)
        img_rgb.save(jpg_buf, "JPEG", quality=quality, optimize=True, progressive=True)
        candidates.append((str(Path(name).with_suffix(".jpg")), jpg_buf.getvalue()))

    # 选择最小体积候选；若并列，优先保留原文件名
    best_name, best_data = min(
        candidates,
        key=lambda it: (len(it[1]), 0 if it[0] == name else 1),
    )

    # 不变小则不替换
    if len(best_data) >= len(data):
        return name, data
    return best_name, best_data


def _pptx_super_dry(all_entries: Dict[str, bytes]) -> None:
    """
    super_dry 模式：
    - 删除 ppt/media 下所有多媒体文件（图片/音频/视频）
    - 删除 slide rels 中所有 image/video/audio 关系
    - 删除 slides xml 中引用这些 rId 的形状节点（pic/movie 等）
    - 清理 [Content_Types].xml 中对应 media 的 Override
    """
    media_prefix = "ppt/media/"

    # 1) 删除实际媒体二进制
    media_keys = [
        k for k in all_entries
        if k.startswith(media_prefix) and Path(k).suffix.lower() in (IMAGE_EXTS | AUDIO_EXTS | VIDEO_EXTS)
    ]
    for k in media_keys:
        all_entries.pop(k, None)

    # 2) 清理每页 rels，记录被删除的 rId
    slide_removed_rids: Dict[str, set] = {}
    rels_names = sorted(
        n for n in all_entries
        if n.startswith("ppt/slides/_rels/") and n.endswith(".xml.rels")
    )
    for rels_name in rels_names:
        xml = all_entries.get(rels_name)
        if not xml:
            continue
        try:
            root = ET.fromstring(xml)
        except Exception:
            continue

        removed = set()
        for rel in list(root):
            rtype = rel.get("Type", "")
            target = rel.get("Target", "")
            rid = rel.get("Id", "")
            is_media_rel = (
                ("/image" in rtype) or ("/video" in rtype) or ("/audio" in rtype)
                or ("../media/" in target)
            )
            if not is_media_rel:
                continue
            root.remove(rel)
            if rid:
                removed.add(rid)

        all_entries[rels_name] = _et_serialize(root)

        slide_xml = rels_name.replace("/slides/_rels/", "/slides/").replace(".rels", "")
        if removed:
            slide_removed_rids[slide_xml] = removed

    # 3) 清理 slide xml 中对这些 rId 的引用节点
    for slide_xml, rid_set in slide_removed_rids.items():
        xml = all_entries.get(slide_xml)
        if not xml:
            continue
        try:
            root = ET.fromstring(xml)
        except Exception:
            continue

        parent_map = {c: p for p in root.iter() for c in p}
        nodes_to_remove = []

        for el in root.iter():
            # 检查任意属性是否引用了待删 rId（尤其 r:embed / r:link）
            hit = any(val in rid_set for val in el.attrib.values())
            if not hit:
                continue

            # 往上找到可删除的形状容器（pic / graphicFrame / movie / obj）
            node = el
            for _ in range(12):
                tag = _etag(node)
                if tag in ("pic", "graphicFrame", "movie", "video", "audio", "obj"):
                    nodes_to_remove.append(node)
                    break
                node = parent_map.get(node)
                if node is None:
                    break

        # 去重并删除
        seen = set()
        for n in nodes_to_remove:
            nid = id(n)
            if nid in seen:
                continue
            seen.add(nid)
            p = parent_map.get(n)
            if p is not None:
                try:
                    p.remove(n)
                except Exception:
                    pass

        all_entries[slide_xml] = _et_serialize(root)

    # 4) 清理 [Content_Types].xml 里已删除 media 的 Override
    ct_name = "[Content_Types].xml"
    if ct_name in all_entries:
        try:
            root = ET.fromstring(all_entries[ct_name])
            for ov in list(root):
                if _etag(ov) != "Override":
                    continue
                part_name = ov.get("PartName", "")
                if not part_name:
                    continue
                normalized = part_name.lstrip("/")
                if normalized.startswith(media_prefix):
                    root.remove(ov)
            all_entries[ct_name] = _et_serialize(root)
        except Exception:
            pass


def compress_pptx_file(src: Path, dst: Path, preset: dict) -> bool:
    """
    压缩 PPTX 文件内嵌的媒体（图片 + 视频）。
    通过操作 ZIP 内部实现，不依赖 python-pptx。
    """
    if not HAS_PILLOW:
        print(f"\n    ✗ 需要 Pillow 才能压缩 PPTX: {src.name}")
        return False

    quality  = preset["doc_quality"]
    max_dim  = preset["doc_max_dim"]
    max_dpi  = preset.get("doc_max_dpi")

    try:
        # 一次性读入全部内容避免上下文问题
        with zipfile.ZipFile(src, "r") as zin:
            all_entries = {name: zin.read(name) for name in zin.namelist()}

        # super_dry：直接删除全部多媒体，只保留文本结构
        if preset.get("super_dry", False):
            _pptx_super_dry(all_entries)
            dst.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
                for name, data in all_entries.items():
                    zout.writestr(name, data)
            return True

        # 预先建立 DPI 像素上限映射：媒体路径 → 最大像素长边
        dpi_limit_map: Dict[str, int] = {}
        if max_dpi:
            dpi_limit_map = _pptx_build_dpi_limit_map(all_entries, max_dpi)

        # 记录名称映射 old_name -> new_name
        name_map: Dict[str, str] = {}
        new_contents: Dict[str, bytes] = {}

        for name, data in all_entries.items():
            parts = name.split("/")
            in_media = len(parts) >= 2 and parts[-2] == "media"
            if not in_media:
                continue
            ext = Path(name).suffix.lower()

            # ── 处理视频 ────────────────────────────────────────────
            if ext in VIDEO_EXTS and HAS_FFMPEG:
                # 需要临时解压视频 → 压缩 → 读回内存
                try:
                    # 写临时文件
                    import tempfile
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tf:
                        tmp_in = Path(tf.name)
                    # 注意：若源文件本来就是 .mp4，with_suffix('.mp4') 会与 tmp_in 同路径
                    # 必须使用不同文件名，否则 ffmpeg 会原地覆盖导致比较失效。
                    tmp_out = tmp_in.with_name(tmp_in.stem + "_compressed.mp4")
                    tmp_in.write_bytes(data)

                    # 压缩视频（复用 compress_video_file 逻辑，但用内存路径）
                    ok = _compress_video_in_memory(tmp_in, tmp_out, preset)
                    if ok and tmp_out.exists():
                        new_data = tmp_out.read_bytes()
                        # 只有压缩后更小才替换
                        if len(new_data) < len(data):
                            new_contents[name] = new_data
                            name_map[name] = name
                        # 清理临时文件
                        tmp_out.unlink(missing_ok=True)
                    tmp_in.unlink(missing_ok=True)
                except Exception as e:
                    print(f"\n      跳过视频 [{name}]: {e}")
                continue

            # ── 处理图片 ─────────────────────────────────────────
            if ext not in IMAGE_EXTS or ext in (".emf", ".wmf"):
                continue
            try:
                new_name, new_data = _pptx_compress_image_entry(
                    name, data, quality, max_dim,
                    dpi_limit=dpi_limit_map.get(name),
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
def _pdf_super_dry(src: Path, dst: Path) -> bool:
    """
    super_dry 模式：仅保留文本内容，删除图片/图形等多媒体。
    实现方式：逐页提取纯文本，重建一个仅文本的新 PDF。
    """
    if not HAS_PYMUPDF:
        print(f"\n    ✗ super_dry 处理 PDF 需要 PyMuPDF: {src.name}")
        return False
    try:
        src_doc = fitz.open(str(src))
        out_doc = fitz.open()

        for page in src_doc:
            rect = page.rect
            new_page = out_doc.new_page(width=rect.width, height=rect.height)
            text = page.get_text("text")
            if not text.strip():
                continue
            margin = 36
            box = fitz.Rect(margin, margin, rect.width - margin, rect.height - margin)
            # 用内置字体写回文本
            new_page.insert_textbox(
                box,
                text,
                fontsize=10,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT,
            )

        out_doc.save(str(dst), garbage=4, deflate=True, clean=True)
        out_doc.close()
        src_doc.close()
        return True
    except Exception as e:
        print(f"\n    ✗ PDF super_dry 失败 [{src.name}]: {e}")
        return False


def compress_pdf_file(src: Path, dst: Path, preset: dict) -> bool:
    """
    按优先级尝试各种 PDF 压缩方案。
    任何方案执行完成后，若输出文件不比原文件小，则用原文件覆盖输出，
    避免"负优化"（文件变大）。
    """
    dst.parent.mkdir(parents=True, exist_ok=True)
    src_size = src.stat().st_size

    # super_dry：只保留文本，直接重建 PDF
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

    # 兜底：若压缩后文件不小于原文件，还原为原始文件
    if ok and dst.exists() and dst.stat().st_size >= src_size:
        shutil.copy2(src, dst)

    return ok


def _compress_pdf_gs(src: Path, dst: Path, preset: dict) -> bool:
    """Ghostscript 方案：稳定可靠，支持完整图片重压缩。"""
    q       = preset["doc_quality"]
    max_dpi = preset.get("doc_max_dpi")

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
    ]

    # 当指定了 max_dpi 时，显式覆盖彩色/灰度/单色图片的分辨率上限
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
    """
    根据页面 MediaBox 尺寸估算在给定 max_dpi 下的像素长边上限。
    PDF 点单位：72 pt = 1 英寸。
    用于 pikepdf 路径：无法精确得知每张图片的缩放矩阵，
    以整页尺寸作为图片可能占据的最大显示面积做保守估算。
    """
    try:
        mb = page.get("/MediaBox")
        if mb and len(mb) >= 4:
            pt_w = float(mb[2]) - float(mb[0])
            pt_h = float(mb[3]) - float(mb[1])
            return int(max(pt_w, pt_h) / 72 * max_dpi)
    except Exception:
        pass
    # 回退：A4 长边 842pt ≈ 11.69 英寸
    return int(11.69 * max_dpi)


def _compress_pdf_pikepdf(src: Path, dst: Path, preset: dict) -> bool:
    """pikepdf 方案：逐页替换图片为 JPEG。"""
    quality = preset["doc_quality"]
    max_dim = preset["doc_max_dim"]
    max_dpi = preset.get("doc_max_dpi")
    try:
        pdf = pikepdf.open(str(src))
        for page in pdf.pages:
            # 若指定了 max_dpi，按页面物理尺寸推算像素上限并与 max_dim 取较小值
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


def _pdf_page_recompress_images(page, quality: int, max_dim: int):
    """
    对 PDF 页面内所有 /Image XObject 进行 JPEG 重压缩。
    只有满足以下两个条件才替换：
      1. 需要缩放（尺寸超出 max_dim），或 质量压缩后确实变小
      2. 新编码后的字节数 < 原始流字节数（read_raw_bytes 大小）
    这样可以避免对已用 FlateDecode+DCTDecode 或 JPXDecode 高效压缩的图片
    进行无效替换导致文件膨胀。
    """
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

                # 记录原始流大小（用于后续比较）
                raw_orig_size = len(bytes(xobj.read_raw_bytes()))
                filt = str(xobj.get("/Filter"))

                pil_img = PdfImage(xobj).as_pil_image()
                orig_w, orig_h = pil_img.size
                pil_img = _resize_if_needed(pil_img, max_dim)
                was_resized = (pil_img.size != (orig_w, orig_h))

                # 对已高效编码（JPX / Flate+DCT）且未缩放的图片，直接跳过
                # 避免把更高效编码降级为 JPEG 导致膨胀。
                if not was_resized and (
                    "/JPXDecode" in filt or ("/FlateDecode" in filt and "/DCTDecode" in filt)
                ):
                    continue

                # 对位图图标/线稿（1bit 或极小图）通常 JPEG 更差，未缩放时跳过
                bpc = int(xobj.get("/BitsPerComponent", 8))
                if not was_resized and (bpc <= 1 or max(orig_w, orig_h) <= 256):
                    continue

                pil_img = _to_rgb(pil_img)
                buf = io.BytesIO()
                pil_img.save(buf, "JPEG", quality=quality, optimize=True)
                jpeg_bytes = buf.getvalue()

                # 核心判断：只在真正节省空间时才替换
                if len(jpeg_bytes) >= raw_orig_size:
                    continue  # 跳过：替换后不会变小

                # 若进行了缩放，要求至少有可见收益（默认至少减少 5%）
                if was_resized and len(jpeg_bytes) > raw_orig_size * 0.95:
                    continue

                # 若没缩放，要求收益更明显（至少减少 10%）
                if not was_resized and len(jpeg_bytes) > raw_orig_size * 0.90:
                    continue

                # ────────────────────────────────────────────────────────  

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
    """PyMuPDF 方案：无法重压缩图片，但可清理冗余结构/重建 xref。
    若重建后反而更大，放弃并报告失败（上层会回退到原文件）。"""
    try:
        doc = fitz.open(str(src))
        doc.save(str(dst), garbage=4, deflate=True, clean=True)
        doc.close()
        # 若无法缩减，返回 False 让上层跳过
        if dst.exists() and dst.stat().st_size >= src.stat().st_size:
            return False
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
  balanced   极限体积（默认） 图片 75%/2560px  视频 x265 CRF34 + FPS<=24  音频 8k
  aggressive 更小体积         图片 60%/1920px  视频 x265 CRF38 + FPS<=24  音频 8k
  high       质量优先         图片 85%/4096px  视频 x265 CRF30 + FPS<=24  音频 8k

示例：
  python compress_media.py 素材/
  python compress_media.py 素材/ -o 素材_压缩/
  python compress_media.py 素材/ --preset aggressive
  python compress_media.py 素材/ --inplace
  python compress_media.py 素材/ -q 80 --max-dim 2048 --max-dpi 150 --crf 25
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
    parser.add_argument("--max-dpi", type=int, metavar="DPI",
                        help="覆盖最大 DPI：独立图片依据嵌入 DPI 元数据缩小像素；"
                             "文档内图片依据显示尺寸×DPI 计算像素上限，与 --max-dim 共同约束")
    parser.add_argument("--crf", type=int, metavar="0-51",
                        help="覆盖视频 CRF 值（越小越好，推荐 18-28）")
    parser.add_argument("--audio-bitrate", metavar="RATE",
                        help="覆盖音频码率（如 128k、192k）")
    parser.add_argument("--quiet", action="store_true",
                        help="静默模式，不逐文件打印进度")
    parser.add_argument("--super-dry", action="store_true",
                        help="超级瘦身模式：PPTX/PDF 删除所有多媒体，仅保留文本")

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
    if args.max_dpi:
        preset["image_max_dpi"] = max(1, args.max_dpi)
        preset["doc_max_dpi"]   = max(1, args.max_dpi)
    if args.crf:
        preset["video_crf"] = max(0, min(51, args.crf))
    if args.audio_bitrate:
        preset["audio_bitrate"]       = args.audio_bitrate
        preset["video_audio_bitrate"] = args.audio_bitrate
    if args.super_dry:
        preset["super_dry"] = True

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
        dpi_show = str(preset.get("doc_max_dpi", "—"))
        print(f"  图片质量：  {preset['image_quality']}%  最大边长：{preset['image_max_dim']}px  最大DPI：{dpi_show}")
        print(f"  视频编码：  {preset['video_codec']}  CRF：{preset['video_crf']}  FPS：<=24")
        print(f"  音频码率：  {preset['audio_bitrate']}  采样率：16kHz 单声道")
        print(f"  super_dry： {'开启（PPTX/PDF 仅保留文本）' if preset.get('super_dry') else '关闭'}")
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
