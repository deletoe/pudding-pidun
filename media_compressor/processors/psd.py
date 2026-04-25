from pathlib import Path
from typing import Optional

from ..deps import HAS_PILLOW, Image
from ..utils.image_utils import _resize_if_needed, _to_rgb


def compress_psd_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
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
