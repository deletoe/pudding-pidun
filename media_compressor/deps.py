import subprocess
from typing import Optional

try:
    from PIL import Image
    HAS_PILLOW = True
except ImportError:
    Image = None
    HAS_PILLOW = False

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    fitz = None
    HAS_PYMUPDF = False

try:
    import pikepdf
    from pikepdf import PdfImage
    HAS_PIKEPDF = True
except ImportError:
    pikepdf = None
    PdfImage = None
    HAS_PIKEPDF = False

try:
    from pptx import Presentation  # noqa: F401
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
