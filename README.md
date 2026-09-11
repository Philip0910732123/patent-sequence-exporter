# Patent Sequence Exporter

> 专利序列自动提取工具 v5.0

从专利文档（PDF / Word）中自动识别并提取蛋白质序列、核酸序列，导出为 FASTA / CSV 格式，支持批量处理。

## 功能特点

- 🔍 **自动识别**：从专利全文中识别 SEQ ID NO 编号的序列段落
- 📄 **多格式输入**：支持 PDF、Word 文档
- 📤 **多格式导出**：FASTA、CSV
- 🔄 **批量处理**：支持多文件批量提取
- 🏷️ **序列编号**：自动保留专利原始 SEQ ID NO 编号
- 🧬 **序列类型识别**：区分蛋白质序列与核酸序列

## 下载使用

1. 前往 [Releases](https://github.com/Philip0910732123/patent-sequence-exporter/releases) 页面
2. 下载最新版 `patent-sequence-exporter-v5.exe`
3. 双击运行即可（无需安装 Python 环境）

> ⚠️ Windows 安全提示：首次运行时可能弹出 SmartScreen 警告，点击"更多信息"→"仍要运行"即可。

## 使用方法

```bash
# GUI 模式（双击运行）
patent-sequence-exporter-v5.exe

# 命令行模式
patent-sequence-exporter-v5.exe --input patent.pdf --output sequences.fasta
```

## 系统要求

- Windows 10/11 (64-bit)
- 无需额外依赖，exe 自带运行环境

## 开发者

- **作者**：Philip0910732123
- **打包方式**：Python + PyInstaller

## 许可证

MIT License © Philip0910732123
