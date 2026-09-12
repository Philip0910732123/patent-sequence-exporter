#!/usr/bin/env python3
"""构建 patent-sequence-exporter-v6.exe（含自动打开浏览器）"""
import subprocess
import sys
import os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

cmd = [
    sys.executable, "-m", "PyInstaller",
    "patent-sequence-exporter-v6.spec",
    "--noconfirm",
    "--clean",
]
print(f"运行: {' '.join(cmd)}")
ret = subprocess.call(cmd)
if ret == 0:
    print("✅ 构建成功")
    dist = os.path.join(os.getcwd(), "dist", "patent-sequence-exporter-v6.exe")
    size = os.path.getsize(dist) / 1024 / 1024
    print(f"   文件: {dist}")
    print(f"   大小: {size:.1f} MB")
else:
    print(f"❌ 构建失败 (exit={ret})")
    sys.exit(1)
