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
from reportlab.lib.pagesizes import A4, landscape
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


def images_to_pdf(folder: str, output_pdf: str, layout: str = "4up_portrait"):
    images = get_sorted_images(folder)
    if not images:
        print(f"❌ 在 '{folder}' 找不到任何圖片！")
        return

    print(f"✅ 找到 {len(images)} 張圖片，開始處理...")

    if layout == "8up_landscape":
        page_w, page_h = landscape(A4)
        cols = 4
        rows = 2
        per_page = 8
        cm_to_pt = 72 / 2.54
        cell_w = 7 * cm_to_pt
        cell_h = 10 * cm_to_pt
        gutter = 5
        margin_x = (page_w - cols * cell_w - (cols - 1) * gutter) / 2
        margin_y = (page_h - rows * cell_h - (rows - 1) * gutter) / 2
    elif layout == "8up_portrait_10x7":
        page_w, page_h = A4
        cols = 2
        rows = 4
        per_page = 8
        cm_to_pt = 72 / 2.54
        cell_w = 10 * cm_to_pt
        cell_h = 7 * cm_to_pt
        gutter = 5
        margin_x = (page_w - cols * cell_w - (cols - 1) * gutter) / 2
        margin_y = (page_h - rows * cell_h - (rows - 1) * gutter) / 2
    else:
        page_w, page_h = A4
        cols = 2
        rows = 2
        per_page = 4
        gutter = 10
        margin_x = 20
        margin_y = 20
        cell_w = (page_w - margin_x * 2 - gutter * (cols - 1)) / cols
        cell_h = (page_h - margin_y * 2 - gutter * (rows - 1)) / rows

    # 建立所有格子的座標 (左下角為原點)
    # 排列順序：從上到下，從左到右
    cells = []
    for row in range(rows):
        for col in range(cols):
            x = margin_x + col * (cell_w + gutter)
            y = page_h - margin_y - (row + 1) * cell_h - row * gutter
            cells.append((x, y))

    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    for page_num, i in enumerate(range(0, len(images), per_page)):
        batch = images[i:i+per_page]
        print(f"  處理第 {page_num + 1} 頁（圖片 {i+1}~{min(i+per_page, len(images))}）...")

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
    print(f"   共 {-(-len(images) // per_page)} 頁（{len(images)} 張圖片）")


if __name__ == "__main__":
    folder = sys.argv[1] if len(sys.argv) > 1 else "."
    output = sys.argv[2] if len(sys.argv) > 2 else "output.pdf"
    layout = sys.argv[3] if len(sys.argv) > 3 else "4up_portrait"
    images_to_pdf(folder, output, layout)
