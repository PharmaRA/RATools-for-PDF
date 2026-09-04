"""测试组合压缩效果"""
import sys
import os
sys.stdout.reconfigure(encoding='utf-8')

from ratools_pdf.pdf.processor import process_document
from PySide6.QtCore import QSettings

original_size = os.path.getsize('test_sample.pdf')
print(f"原始文件: {original_size / 1024:.2f} KB\n")

# 测试1: 仅标准压缩
print("=" * 60)
print("测试1: 仅标准压缩")
print("=" * 60)
success, msg = process_document('test_sample.pdf', 'out1_standard.pdf', ['compress_standard'])
if os.path.exists('out1_standard.pdf'):
    size = os.path.getsize('out1_standard.pdf')
    print(f"文件大小: {size / 1024:.2f} KB")
    print(f"压缩率: {(1 - size/original_size) * 100:.1f}%\n")

# 测试2: 仅深度压缩
print("=" * 60)
print("测试2: 仅深度压缩")
print("=" * 60)
success, msg = process_document('test_sample.pdf', 'out2_aggressive.pdf', ['compress_aggressive'])
if os.path.exists('out2_aggressive.pdf'):
    size = os.path.getsize('out2_aggressive.pdf')
    print(f"文件大小: {size / 1024:.2f} KB")
    print(f"压缩率: {(1 - size/original_size) * 100:.1f}%\n")

# 测试3: 仅图像压缩 150 DPI
settings = QSettings("settings.ini", QSettings.Format.IniFormat)
settings.setValue("Compression/ImageDPI", 150)
print("=" * 60)
print("测试3: 仅图像压缩 (150 DPI)")
print("=" * 60)
success, msg = process_document('test_sample.pdf', 'out3_image150.pdf', ['compress_images'])
if os.path.exists('out3_image150.pdf'):
    size = os.path.getsize('out3_image150.pdf')
    print(f"文件大小: {size / 1024:.2f} KB")
    print(f"压缩率: {(1 - size/original_size) * 100:.1f}%\n")

# 测试4: 深度压缩 + 图像压缩 150 DPI
print("=" * 60)
print("测试4: 深度压缩 + 图像压缩 (150 DPI)")
print("=" * 60)
success, msg = process_document('test_sample.pdf', 'out4_both.pdf', ['compress_aggressive', 'compress_images'])
if os.path.exists('out4_both.pdf'):
    size = os.path.getsize('out4_both.pdf')
    print(f"文件大小: {size / 1024:.2f} KB")
    print(f"压缩率: {(1 - size/original_size) * 100:.1f}%\n")

print("=" * 60)
print("所有测试完成！")
print("=" * 60)
