import shutil
import subprocess
from pathlib import Path
from typing import Optional

from ..deps import HAS_FFMPEG


def _probe_video_fps(src: Path) -> Optional[float]:
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
    return (
        f"scale='if(gt(iw,ih),if(gt(iw,{max_dim}),{max_dim},iw),-2)':"
        f"'if(gt(ih,iw),if(gt(ih,{max_dim}),{max_dim},ih),-2)':"
        "force_original_aspect_ratio=decrease,"
        "scale=trunc(iw/2)*2:trunc(ih/2)*2"
    )


def _compress_video_in_memory(src: Path, dst: Path, preset: dict) -> bool:
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
        "-ar", "12000",
        "-ac", "1",
        "-movflags", "+faststart",
    ]
    if preset.get("video_codec") == "libx265":
        cmd += ["-tag:v", "hvc1", "-x265-params", "log-level=error"]
    cmd.append(str(dst))

    result = subprocess.run(cmd, capture_output=True, timeout=7200)
    if result.returncode != 0:
        return False
    if dst.stat().st_size >= src.stat().st_size:
        dst.unlink(missing_ok=True)
        return False
    return True


def compress_video_file(src: Path, dst: Path, preset: dict) -> Optional[Path]:
    if not HAS_FFMPEG:
        return None
    out_path = dst.with_suffix(".mp4")
    out_path.parent.mkdir(parents=True, exist_ok=True)

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
        "-ar", "12000",
        "-ac", "1",
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
