"""创建测试用 PDF 文件"""
import fitz
from PIL import Image
import io

# 创建一个包含大图像的 PDF
doc = fitz.open()
page = doc.new_page(width=595, height=842)  # A4

# 添加一些文字
page.insert_text((50, 50), "测试 PDF 压缩功能", fontsize=20)
page.insert_text((50, 100), "此文件包含高分辨率图像用于测试图像压缩", fontsize=12)

# 创建一个高分辨率图像 (2400x3000)
img = Image.new('RGB', (2400, 3000), color=(100, 150, 200))
# 添加一些渐变效果
pixels = img.load()
for y in range(3000):
    for x in range(2400):
        pixels[x, y] = (100 + x % 156, 150 + y % 106, 200 - (x + y) % 56)

# 转换为字节
img_bytes = io.BytesIO()
img.save(img_bytes, format='PNG')
img_bytes.seek(0)

# 插入图像到 PDF
rect = fitz.Rect(50, 150, 545, 792)
page.insert_image(rect, stream=img_bytes.getvalue())

# 保存
doc.save("test_sample.pdf", garbage=0, deflate=False)  # 不压缩，方便测试
doc.close()

print("创建测试文件: test_sample.pdf")
import os
size = os.path.getsize("test_sample.pdf")
print(f"文件大小: {size / 1024:.2f} KB")
