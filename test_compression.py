"""测试压缩功能的简单脚本"""
import os
from ratools_pdf.pdf.processor import process_pdf

def test_compression():
    """测试各压缩级别"""
    test_pdf = "test_sample.pdf"

    if not os.path.exists(test_pdf):
        print(f"测试文件 {test_pdf} 不存在，请提供一个测试 PDF")
        return

    # 获取原始文件大小
    original_size = os.path.getsize(test_pdf)
    print(f"原始文件大小: {original_size / 1024:.2f} KB")

    # 测试标准压缩
    print("\n测试标准压缩 (garbage=3)...")
    try:
        result = process_pdf(
            test_pdf,
            "test_standard.pdf",
            ["compress_standard"]
        )
        if os.path.exists("test_standard.pdf"):
            new_size = os.path.getsize("test_standard.pdf")
            print(f"  压缩后: {new_size / 1024:.2f} KB")
            print(f"  压缩率: {(1 - new_size/original_size) * 100:.1f}%")
            print(f"  处理日志: {result['marks']}")
    except Exception as e:
        print(f"  错误: {e}")

    # 测试深度压缩
    print("\n测试深度压缩 (garbage=4 + clean)...")
    try:
        result = process_pdf(
            test_pdf,
            "test_aggressive.pdf",
            ["compress_aggressive"]
        )
        if os.path.exists("test_aggressive.pdf"):
            new_size = os.path.getsize("test_aggressive.pdf")
            print(f"  压缩后: {new_size / 1024:.2f} KB")
            print(f"  压缩率: {(1 - new_size/original_size) * 100:.1f}%")
            print(f"  处理日志: {result['marks']}")
    except Exception as e:
        print(f"  错误: {e}")

    print("\n✓ 压缩功能测试完成")
    print("\n注意：图像压缩需要勾选 'compress_images' 选项并在 UI 中设置 DPI")

if __name__ == "__main__":
    test_compression()
