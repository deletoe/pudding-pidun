def fmt_size(b: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if b < 1024:
            return f"{b:.1f} {unit}"
        b /= 1024
    return f"{b:.1f} TB"


class Stats:
    def __init__(self):
        self.processed = 0
        self.skipped = 0
        self.errors = 0
        self.original_bytes = 0
        self.compressed_bytes = 0

    def add(self, orig: int, comp: int):
        self.processed += 1
        self.original_bytes += orig
        self.compressed_bytes += comp

    def report(self):
        saved = self.original_bytes - self.compressed_bytes
        ratio = (saved / self.original_bytes * 100) if self.original_bytes > 0 else 0.0
        print(f"\n{'=' * 52}")
        print("  处理完成")
        print(f"{'=' * 52}")
        print(f"  已压缩：{self.processed} 个文件")
        print(f"  已跳过：{self.skipped} 个文件（非媒体，直接复制）")
        print(f"  出错：  {self.errors} 个文件（已复制原件）")
        print(f"  原始大小：  {fmt_size(self.original_bytes)}")
        print(f"  压缩后大小：{fmt_size(self.compressed_bytes)}")
        print(f"  共节省：    {fmt_size(saved)}  ({ratio:.1f}%)")
        print()
