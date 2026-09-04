"""直接测试压缩功能"""
import sys
import os

# 设置 UTF-8 输出
sys.stdout.reconfigure(encoding='utf-8')

from ratools_pdf.pdf.processor import process_document

# 测试标准压缩
print("测试标准压缩...")
success, msg = process_document('test_sample.pdf', 'test_standard.pdf', ['compress_standard'])
print(f"成功: {success}")
print(f"消息: {msg}")

if os.path.exists('test_standard.pdf'):
    original_size = os.path.getsize('test_sample.pdf')
    new_size = os.path.getsize('test_standard.pdf')
    print(f"原始: {original_size / 1024:.2f} KB")
    print(f"压缩后: {new_size / 1024:.2f} KB")
    print(f"压缩率: {(1 - new_size/original_size) * 100:.1f}%")
