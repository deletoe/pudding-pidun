import io
from pathlib import Path
from typing import Optional

from ..deps import HAS_PILLOW, Image
from ..utils.image_utils import _resize_if_needed, _to_rgb


def compress_image_to_jpeg_bytes(data: bytes, quality: int, max_dim: int) -> bytes:
    img = Image.open(io.BytesIO(data))
    img.load()
    img = _resize_if_needed(img, max_dim)
    img = _to_rgb(img)
    out = io.BytesIO()
    img.save(out, "JPEG", quality=quality, optimize=True, progressive=True)
    return out.getvalue()


def compress_image_to_png_bytes(data: bytes, max_dim: int) -> bytes:
    img = Image.open(io.BytesIO(data))
    img.load()
    img = _resize_if_needed(img, max_dim)
    out = io.BytesIO()
    img.save(out, "PNG", optimize=True)
    return out.getvalue()


def compress_image_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    if not HAS_PILLOW:
        return None
    try:
        data = src.read_bytes()
        max_dim = preset["image_max_dim"]
        max_dpi = preset.get("image_max_dpi")

        if max_dpi:
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
                scale = max_dpi / img_dpi
                dpi_limit = int(max(img.width, img.height) * scale)
                max_dim = min(max_dim, dpi_limit)

        compressed = compress_image_to_jpeg_bytes(data, preset["image_quality"], max_dim)
        out_data = compressed if len(compressed) < len(data) else data
        out_path = dst.with_suffix(".jpg")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(out_data)
        return out_path
    except Exception as e:
        print(f"\n    ✗ 图片压缩失败 [{src.name}]: {e}")
        return None
