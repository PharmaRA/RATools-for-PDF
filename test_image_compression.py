"""测试图像压缩功能"""
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

from ratools_pdf.pdf.processor import process_document
from PySide6.QtCore import QSettings

# 设置 DPI 到配置文件
settings = QSettings("settings.ini", QSettings.Format.IniFormat)
settings.setValue("Compression/ImageDPI", 150)
print("已设置 DPI = 150")

# 测试图像压缩
print("\n测试图像压缩（150 DPI）...")
success, msg = process_document('test_sample.pdf', 'test_image_150dpi.pdf', ['compress_images'])
print(f"成功: {success}")
print(f"消息: {msg}")

if os.path.exists('test_image_150dpi.pdf'):
    original_size = os.path.getsize('test_sample.pdf')
    new_size = os.path.getsize('test_image_150dpi.pdf')
    print(f"原始: {original_size / 1024:.2f} KB")
    print(f"压缩后: {new_size / 1024:.2f} KB")
    print(f"压缩率: {(1 - new_size/original_size) * 100:.1f}%")

# 测试 300 DPI
settings.setValue("Compression/ImageDPI", 300)
print("\n\n测试图像压缩（300 DPI）...")
success, msg = process_document('test_sample.pdf', 'test_image_300dpi.pdf', ['compress_images'])
print(f"成功: {success}")
print(f"消息: {msg}")

if os.path.exists('test_image_300dpi.pdf'):
    new_size = os.path.getsize('test_image_300dpi.pdf')
    print(f"压缩后: {new_size / 1024:.2f} KB")
    print(f"压缩率: {(1 - new_size/original_size) * 100:.1f}%")
