"""
將目錄下的圖片（1.png ~ N.png）每4張合併成一頁，輸出為單一 PDF 檔案。

使用方式：
    python combine_images_to_pdf.py [圖片目錄] [輸出PDF路徑]

預設值：
    圖片目錄 = 當前目錄 (.)
    輸出PDF  = output.pdf

安裝依賴：
    pip install Pillow reportlab
"""

import sys
from pathlib import Path
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def get_sorted_images(folder: str) -> list[Path]:
    """取得資料夾內所有 .png/.jpg 圖片，依數字順序排列"""
    folder = Path(folder)
    images = []
    for ext in ("*.png", "*.jpg", "*.jpeg"):
        images.extend(folder.glob(ext))

    def sort_key(p):
        try:
            return int(p.stem)
        except ValueError:
            return float("inf")

    return sorted(images, key=sort_key)


def fit_rect(img_w, img_h, cell_w, cell_h):
    """計算圖片在格子內等比例縮放後的尺寸與置中偏移"""
    scale = min(cell_w / img_w, cell_h / img_h)
    new_w = img_w * scale
    new_h = img_h * scale
    x_off = (cell_w - new_w) / 2
    y_off = (cell_h - new_h) / 2
    return new_w, new_h, x_off, y_off


def images_to_pdf(folder: str, output_pdf: str):
    images = get_sorted_images(folder)
    if not images:
        print(f"❌ 在 '{folder}' 找不到任何圖片！")
        return

    print(f"✅ 找到 {len(images)} 張圖片，開始處理...")

    page_w, page_h = A4   # 595 x 842 points
    margin = 20           # 頁面外邊距
    gutter = 10           # 圖片間距

    cell_w = (page_w - margin * 2 - gutter) / 2
    cell_h = (page_h - margin * 2 - gutter) / 2

    # 2x2 格子左下角座標（reportlab 原點在左下）
    # 排列：左上=0, 右上=1, 左下=2, 右下=3
    cells = [
        (margin,                   margin + gutter + cell_h),
        (margin + gutter + cell_w, margin + gutter + cell_h),
        (margin,                   margin),
        (margin + gutter + cell_w, margin),
    ]

    c = canvas.Canvas(output_pdf, pagesize=A4)

    for page_num, i in enumerate(range(0, len(images), 4)):
        batch = images[i:i+4]
        print(f"  處理第 {page_num + 1} 頁（圖片 {i+1}~{min(i+4, len(images))}）...")

        for j, img_path in enumerate(batch):
            with Image.open(img_path) as im:
                img_w, img_h = im.size

            draw_w, draw_h, x_off, y_off = fit_rect(img_w, img_h, cell_w, cell_h)
            cell_x, cell_y = cells[j]
            draw_x = cell_x + x_off
            draw_y = cell_y + y_off

            # 直接傳路徑給 reportlab，避免 PIL 二次縮放導致模糊
            c.drawImage(
                str(img_path),
                draw_x, draw_y,
                width=draw_w,
                height=draw_h,
            )

        c.showPage()

    c.save()
    print(f"\n🎉 完成！PDF 已儲存至：{output_pdf}")
    print(f"   共 {-(-len(images) // 4)} 頁（{len(images)} 張圖片）")


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    output  = sys.argv[2] if len(sys.argv) > 2 else "output.pdf"
    images_to_pdf(folder, output)
