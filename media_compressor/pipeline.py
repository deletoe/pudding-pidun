import shutil
import tempfile
from pathlib import Path
from typing import Optional

from .constants import AUDIO_EXTS, IMAGE_EXTS, VIDEO_EXTS
from .processors.ai import compress_ai_file
from .processors.audio import compress_audio_file
from .processors.image import compress_image_file
from .processors.keynote import compress_keynote_file
from .processors.pdf import compress_pdf_file
from .processors.pptx import compress_pptx_file
from .processors.psd import compress_psd_file
from .processors.video import compress_video_file
from .stats import Stats, fmt_size


def _fallback_copy(src: Path, dst: Path):
    dst.parent.mkdir(parents=True, exist_ok=True)
    if src.is_dir():
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(src, dst)
    else:
        shutil.copy2(src, dst)


def _path_total_size(path: Path) -> int:
    if path.is_file():
        return path.stat().st_size
    total = 0
    for child in path.rglob("*"):
        if child.is_file():
            total += child.stat().st_size
    return total


def process_file(src: Path, dst: Path, preset: dict, stats: Stats, verbose: bool):
    ext = src.suffix.lower()
    orig_size = _path_total_size(src)

    if verbose:
        print(f"  {src.name}  ({fmt_size(orig_size)})", end="", flush=True)

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
        _fallback_copy(src, dst)
        if verbose:
            print("  → 跳过（旧版 .ppt 需 LibreOffice 转换，已复制原件）")
        stats.skipped += 1
        return

    elif ext == ".pdf":
        success = compress_pdf_file(src, dst, preset)
        result_path = dst if success else None

    elif ext == ".key":
        success = compress_keynote_file(src, dst, preset)
        result_path = dst if success else None

    elif ext == ".psd":
        result_path = compress_psd_file(src, dst, preset)
        success = result_path is not None

    elif ext == ".ai":
        success = compress_ai_file(src, dst, preset)
        result_path = dst.with_suffix(".pdf") if success else None

    else:
        _fallback_copy(src, dst)
        if verbose:
            print("  → 已复制（非媒体文件）")
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
            print("  → 压缩失败，已复制原件")


def process_folder(src: Path, dst: Path, preset: dict, stats: Stats, verbose: bool):
    for item in sorted(src.iterdir()):
        if item.name.startswith("."):
            continue
        if item.is_symlink():
            continue
        if item.is_dir() and item.suffix.lower() == ".key":
            process_file(item, dst / item.name, preset, stats, verbose)
        elif item.is_dir():
            process_folder(item, dst / item.name, preset, stats, verbose)
        elif item.is_file():
            process_file(item, dst / item.name, preset, stats, verbose)


def process_inplace(src: Path, preset: dict, stats: Stats, verbose: bool):
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_root = Path(tmpdir) / "work"
        tmp_root.mkdir()
        process_folder(src, tmp_root, preset, stats, verbose=False)
        for tmp_file in tmp_root.rglob("*"):
            if not tmp_file.is_file():
                continue
            rel = tmp_file.relative_to(tmp_root)
            orig = src / rel.parent / rel.stem
            target = src / rel
            if not target.exists():
                candidates = list((src / rel.parent).glob(rel.stem + ".*"))
                target = candidates[0] if candidates else src / rel

            if target.exists() and tmp_file.stat().st_size < target.stat().st_size:
                shutil.copy2(tmp_file, target)
                if verbose:
                    print(f"  替换: {target.relative_to(src)}")
            elif not target.exists():
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(tmp_file, target)
