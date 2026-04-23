import shutil
import subprocess
import tempfile
import unittest
import zipfile
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
        self.assertIn("Keynote(.key)", result.stdout)

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

    def test_keynote_zip_path(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            src = tmp / "input"
            dst = tmp / "output"
            src.mkdir()

            key_file = src / "demo.key"
            with zipfile.ZipFile(key_file, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("Data/", b"")
                zf.writestr("Data/hello.txt", b"hello")

            result = subprocess.run(
                ["python", str(SCRIPT), str(src), "-o", str(dst), "--quiet"],
                capture_output=True,
                text=True,
                cwd=str(ROOT),
                timeout=60,
            )
            self.assertEqual(result.returncode, 0, msg=result.stderr)
            out_key = dst / "demo.key"
            self.assertTrue(out_key.exists())
            with zipfile.ZipFile(out_key, "r") as zf:
                self.assertIn("Data/hello.txt", zf.namelist())
                self.assertEqual(zf.read("Data/hello.txt"), b"hello")

    def test_super_dry_and_inplace_mode(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            src = tmp / "input"
            src.mkdir()
            f = src / "note.txt"
            f.write_text("smoke", encoding="utf-8")

            key_file = src / "demo.key"
            with zipfile.ZipFile(key_file, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("Data/", b"")
                zf.writestr("Data/hello.txt", b"hello")

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
            self.assertTrue(key_file.exists())


if __name__ == "__main__":
    unittest.main()
