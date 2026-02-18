# Office Image Extractor

这是一个简单的 Python GUI 工具，用于从 Microsoft Office 文档（.pptx 和 .docx）中提取原始图片。

## 功能特点

- **原始质量**：直接从文档结构中提取图片，不进行重编码，保留原始清晰度。
- **拖拽支持**：直接将文件拖入应用窗口即可。
- **批量处理**：自动创建文件夹保存提取的图片。
- **格式支持**：提取所有嵌入的媒体文件（包括 png, jpg, emf, wmf 等）。

## 如何运行

### 方法 1：直接运行源码（推荐开发人员）

1. 确保已安装 Python 3.x。
2. 克隆仓库：
   ```bash
   git clone https://github.com/SmallZhao-1/OfficeImageExtractor.git
   cd OfficeImageExtractor
   ```
3. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
4. 运行程序：
   ```bash
   python app_v2.py
   ```

### 方法 2：使用打包好的 Exe（推荐普通用户）
(如果您已经生成了 exe 文件，可以在此处提供下载链接或说明在 dist 文件夹中找到)
