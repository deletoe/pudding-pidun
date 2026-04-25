from .ai import compress_ai_file
from .audio import compress_audio_file
from .image import compress_image_file, compress_image_to_jpeg_bytes, compress_image_to_png_bytes
from .keynote import compress_keynote_file
from .pdf import compress_pdf_file
from .pptx import compress_pptx_file
from .psd import compress_psd_file
from .video import compress_video_file

__all__ = [
    "compress_ai_file",
    "compress_audio_file",
    "compress_image_file",
    "compress_image_to_jpeg_bytes",
    "compress_image_to_png_bytes",
    "compress_keynote_file",
    "compress_pdf_file",
    "compress_pptx_file",
    "compress_psd_file",
    "compress_video_file",
]
