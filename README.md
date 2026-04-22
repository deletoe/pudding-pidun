# 媒体文件有损压缩工具

对文件夹内所有多媒体文件进行批量有损压缩，在不影响观感的前提下尽可能缩小体积。

支持处理独立的图片、音频、视频文件，也能深入 PPTX、PDF、PSD、AI 文档内部，对其中嵌入的媒体元素逐一压缩。

---

## 支持的文件类型

| 类型 | 扩展名 | 处理方式 |
|------|--------|----------|
| 图片 | `.jpg` `.png` `.gif` `.bmp` `.tiff` `.webp` | 转换为 JPEG，超长边等比缩放 |
| 音频 | `.mp3` `.wav` `.flac` `.ogg` `.m4a` `.aac` `.wma` `.opus` | 重编码为 AAC |
| 视频 | `.mp4` `.mov` `.avi` `.mkv` `.wmv` `.flv` `.webm` | 重编码为 H.264 MP4 |
| 演示文稿 | `.pptx` | 压缩内部嵌入图片，保持幻灯片尺寸不变 |
| PDF | `.pdf` | 重压缩嵌入图片（需 Ghostscript 或 pikepdf） |
| Photoshop | `.psd` | 合并图层后转为 JPEG |
| Illustrator | `.ai` | 作为 PDF 解析压缩（适用于 CS 及以后版本） |

> 文档内图片（PPTX/PDF）：降低像素密度和质量，但**不改变文档中的显示尺寸**。

---

## 安装

### 1. 克隆 / 下载项目

```bash
git clone <repo-url>
cd budingyamianbao
```

### 2. 创建虚拟环境并安装 Python 依赖

```bash
python -m venv venv

# Windows
.\venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

### 3. 安装外部工具

#### ffmpeg（处理音频/视频必需）

- **Windows**：从 https://ffmpeg.org/download.html 下载，将 `bin/` 目录加入系统 `PATH`
- **macOS**：`brew install ffmpeg`
- **Linux**：`sudo apt install ffmpeg`

#### Ghostscript（可选，PDF 图片压缩效果最佳）

- **Windows**：从 https://ghostscript.com/releases/gsdnld.html 下载安装
- **macOS**：`brew install ghostscript`
- **Linux**：`sudo apt install ghostscript`

---

## 使用方法

```bash
# 激活虚拟环境后运行
python compress_media.py <输入文件夹> [选项]
```

### 常用示例

```bash
# 压缩到新文件夹（默认，输出至 素材_compressed/）
python compress_media.py 素材/

# 指定输出目录
python compress_media.py 素材/ -o 素材_压缩版/

# 原地压缩，直接替换原文件（建议先备份！）
python compress_media.py 素材/ --inplace

# 使用激进压缩预设
python compress_media.py 素材/ --preset aggressive

# 自定义参数（含 DPI 限制）
python compress_media.py 素材/ -q 80 --max-dim 2048 --max-dpi 150 --crf 25 --audio-bitrate 96k

# 静默模式（不逐文件打印进度）
python compress_media.py 素材/ --quiet
```

### 全部参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入文件夹路径 | 必填 |
| `-o, --output` | 输出文件夹路径 | `<input>_compressed` |
| `-p, --preset` | 压缩预设：`balanced` / `aggressive` / `high` | `balanced` |
| `--inplace` | 原地压缩，替换原文件 | 关闭 |
| `-q, --quality` | 图片 JPEG 质量（1–95） | 由预设决定 |
| `--max-dim` | 图片/文档内图片的像素长边硬性上限 | 由预设决定 |
| `--max-dpi` | 最大 DPI（与 `--max-dim` 共同约束像素数，取较小值） | 由预设决定 |
| `--crf` | 视频 CRF 值（0–51，越小越好） | 由预设决定 |
| `--audio-bitrate` | 音频码率（如 `128k`、`192k`） | 由预设决定 |
| `--quiet` | 静默模式 | 关闭 |

---

## 压缩预设

| 预设 | 图片质量 | 最大边长 | 最大 DPI | 视频 CRF | 视频最大分辨率 | 音频码率 |
|------|----------|----------|----------|----------|----------------|----------|
| `balanced`（默认） | 75% | 2560 px | 150 | 23 | 1920×1080 | 128 kbps |
| `aggressive` | 60% | 1920 px | 96 | 28 | 1280×720 | 96 kbps |
| `high` | 85% | 4096 px | 200 | 20 | 3840×2160 | 192 kbps |

`balanced` 预设下，大多数图片/视频**肉眼感知不到质量损失**，体积可减少 40%–70%。

### `--max-dpi` 与 `--max-dim` 的关系

两个参数同时生效，脚本取**更严格**（更小）的一侧作为实际像素约束：

```
实际像素长边 = min(max-dim, 显示尺寸(英寸) × max-dpi)
```

| 场景 | 生效逻辑 |
|------|----------|
| 独立图片（有 DPI 元数据） | 按 `当前DPI > max-dpi` 时等比缩小像素，再受 `max-dim` 限制 |
| PPTX 内嵌图片 | 解析幻灯片 XML 获取真实显示尺寸（EMU），精确计算每张图片的像素上限 |
| PDF（Ghostscript） | 直接传入 `-dColorImageResolution` 参数，由 GS 精确控制 |
| PDF（pikepdf） | 以页面物理尺寸（MediaBox）作为图片最大显示面积的保守估算 |

---

## 依赖说明

| 依赖 | 用途 | 安装方式 |
|------|------|----------|
| [Pillow](https://pillow.readthedocs.io/) | 图片读写与压缩（必需） | `pip install Pillow` |
| [pikepdf](https://pikepdf.readthedocs.io/) | PDF 内嵌图片替换 | `pip install pikepdf` |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF 结构优化（回退方案） | `pip install PyMuPDF` |
| [python-pptx](https://python-pptx.readthedocs.io/) | PPTX 导入验证 | `pip install python-pptx` |
| ffmpeg | 音频/视频重编码 | 见上方安装说明 |
| Ghostscript | PDF 图片深度压缩（最优方案） | 见上方安装说明，可选 |

未安装的依赖不会导致脚本崩溃，对应类型的文件会跳过压缩并保留原件。

---

## 注意事项

- **建议先备份原始文件**，尤其是使用 `--inplace` 模式时。
- PSD 文件压缩后会**合并所有图层**并转为 JPEG，不可逆，原 `.psd` 文件不会被删除（保留在原目录）。
- 旧版二进制 `.ppt` 格式（PowerPoint 97–2003）暂不支持内部压缩，会直接复制原件。
- AI 文件仅支持 Adobe Illustrator CS（2003）及以后基于 PDF 的格式；更早的 PostScript 格式无法处理。
- 若压缩后文件反而变大，脚本会自动保留原始文件。
