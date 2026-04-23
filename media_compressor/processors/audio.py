import shutil
import subprocess
from pathlib import Path
from typing import Optional

from ..deps import HAS_FFMPEG


def compress_audio_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
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
    if out_path.stat().st_size >= src.stat().st_size:
        out_path.unlink()
        shutil.copy2(src, dst)
        return dst
    return out_path
