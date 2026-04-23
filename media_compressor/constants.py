from typing import Dict

PRESETS: Dict[str, dict] = {
    "balanced": {
        "super_dry": False,
        "image_quality": 75,
        "image_max_dim": 2560,
        "image_max_dpi": 150,
        "audio_codec": "aac",
        "audio_bitrate": "8k",
        "video_codec": "libx265",
        "video_crf": 34,
        "video_preset": "medium",
        "video_audio_bitrate": "8k",
        "doc_quality": 72,
        "doc_max_dim": 1920,
        "doc_max_dpi": 150,
    },
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

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
AUDIO_EXTS = {".mp3", ".aac", ".m4a", ".wav", ".flac", ".ogg", ".opus", ".wma"}
VIDEO_EXTS = {".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".webm", ".m4v", ".3gp"}
DOC_EXTS = {".pptx", ".ppt", ".pdf", ".psd", ".ai"}

CONTENT_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".webp": "image/webp",
}
