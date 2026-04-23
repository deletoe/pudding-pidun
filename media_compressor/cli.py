import argparse
import sys
from pathlib import Path

from .constants import PRESETS
from .deps import GS_CMD, HAS_FFMPEG, HAS_PIKEPDF, HAS_PILLOW, HAS_PPTX, HAS_PYMUPDF
from .pipeline import process_folder, process_inplace
from .stats import Stats


def main():
    parser = argparse.ArgumentParser(
        prog="compress_media",
        description="媒体文件有损压缩工具 v1.0\n\n"
        "对文件夹内所有图片/音频/视频及 PPTX/PDF/PSD/AI 中的嵌入媒体\n"
        "进行有损压缩，在尽量不影响观感的前提下减小文件体积。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
预设说明：
  balanced   极限体积（默认） 图片 75%/2560px  视频 x265 CRF34 + FPS<=24  音频 24k
  aggressive 更小体积         图片 60%/1920px  视频 x265 CRF38 + FPS<=24  音频 16k
  high       质量优先         图片 85%/4096px  视频 x265 CRF30 + FPS<=24  音频 32k

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
        "-p",
        "--preset",
        choices=["balanced", "aggressive", "high"],
        default="balanced",
        help="压缩预设（默认: balanced）",
    )
    parser.add_argument("--inplace", action="store_true", help="原地压缩，直接替换原文件（谨慎使用）")
    parser.add_argument("-q", "--quality", type=int, metavar="1-95", help="覆盖图片 JPEG 质量（1-95）")
    parser.add_argument("--max-dim", type=int, metavar="PX", help="覆盖图片/文档图片最大边长（像素）")
    parser.add_argument(
        "--max-dpi",
        type=int,
        metavar="DPI",
        help="覆盖最大 DPI：独立图片依据嵌入 DPI 元数据缩小像素；"
        "文档内图片依据显示尺寸×DPI 计算像素上限，与 --max-dim 共同约束",
    )
    parser.add_argument("--crf", type=int, metavar="0-51", help="覆盖视频 CRF 值（越小越好，推荐 18-28）")
    parser.add_argument("--audio-bitrate", metavar="RATE", help="覆盖音频码率（如 128k、192k）")
    parser.add_argument("--quiet", action="store_true", help="静默模式，不逐文件打印进度")
    parser.add_argument("--super-dry", action="store_true", help="超级瘦身模式：PPTX/PDF 删除所有多媒体，仅保留文本")

    args = parser.parse_args()
    verbose = not args.quiet

    src = Path(args.input).resolve()
    if not src.exists() or not src.is_dir():
        print(f"错误：输入路径不存在或不是文件夹：{src}")
        sys.exit(1)

    preset = dict(PRESETS[args.preset])
    if args.quality:
        preset["image_quality"] = max(1, min(95, args.quality))
        preset["doc_quality"] = max(1, min(95, args.quality))
    if args.max_dim:
        preset["image_max_dim"] = args.max_dim
        preset["doc_max_dim"] = args.max_dim
    if args.max_dpi:
        preset["image_max_dpi"] = max(1, args.max_dpi)
        preset["doc_max_dpi"] = max(1, args.max_dpi)
    if args.crf:
        preset["video_crf"] = max(0, min(51, args.crf))
    if args.audio_bitrate:
        preset["audio_bitrate"] = args.audio_bitrate
        preset["video_audio_bitrate"] = args.audio_bitrate
    if args.super_dry:
        preset["super_dry"] = True

    if verbose:
        print()
        print("  媒体文件有损压缩工具 v1.0")
        print("  " + "─" * 48)
        print(f"  输入目录：  {src}")
        if not args.inplace:
            dst_show = args.output if args.output else str(src.parent / (src.name + "_compressed"))
            print(f"  输出目录：  {dst_show}")
        else:
            print("  模式：      原地压缩（inplace）")
        print(f"  预设：      {args.preset}")
        dpi_show = str(preset.get("doc_max_dpi", "—"))
        print(f"  图片质量：  {preset['image_quality']}%  最大边长：{preset['image_max_dim']}px  最大DPI：{dpi_show}")
        print(f"  视频编码：  {preset['video_codec']}  CRF：{preset['video_crf']}  FPS：<=24")
        print(f"  音频码率：  {preset['audio_bitrate']}  采样率：12kHz 单声道")
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
