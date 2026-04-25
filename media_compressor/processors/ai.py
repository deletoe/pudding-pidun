import shutil
from pathlib import Path

from .pdf import compress_pdf_file


def compress_ai_file(src: Path, dst: Path, preset: dict) -> bool:
    dst_pdf = dst.with_suffix(".pdf")
    if compress_pdf_file(src, dst_pdf, preset):
        print(f"\n      注意: AI 已作为 PDF 压缩 ({src.name} → {dst_pdf.name})")
        return True
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    except Exception:
        pass
    print(f"\n      注意: AI 文件无法深度压缩，已复制原件 ({src.name})")
    return False
