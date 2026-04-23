import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "compress_media.py"


class TestCliSmoke(unittest.TestCase):
    def test_help(self):
        result = subprocess.run(
            ["python", str(SCRIPT), "-h"],
            capture_output=True,
            text=True,
            cwd=str(ROOT),
            timeout=30,
        )
        self.assertEqual(result.returncode, 0)
        self.assertIn("媒体文件有损压缩工具", result.stdout)

    def test_basic_folder_mode(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            src = tmp / "input"
            dst = tmp / "output"
            src.mkdir()
            (src / "note.txt").write_text("smoke", encoding="utf-8")

            result = subprocess.run(
                ["python", str(SCRIPT), str(src), "-o", str(dst), "--quiet"],
                capture_output=True,
                text=True,
                cwd=str(ROOT),
                timeout=60,
            )
            self.assertEqual(result.returncode, 0, msg=result.stderr)
            self.assertTrue((dst / "note.txt").exists())
            self.assertEqual((dst / "note.txt").read_text(encoding="utf-8"), "smoke")

    def test_super_dry_and_inplace_mode(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            src = tmp / "input"
            src.mkdir()
            f = src / "note.txt"
            f.write_text("smoke", encoding="utf-8")

            result = subprocess.run(
                ["python", str(SCRIPT), str(src), "--super-dry", "--inplace", "--quiet"],
                capture_output=True,
                text=True,
                cwd=str(ROOT),
                timeout=60,
            )
            self.assertEqual(result.returncode, 0, msg=result.stderr)
            self.assertTrue(f.exists())
            self.assertEqual(f.read_text(encoding="utf-8"), "smoke")


if __name__ == "__main__":
    unittest.main()
