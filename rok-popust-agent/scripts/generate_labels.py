"""
Rok-Popust Agent - генератор на ценовни налепници (термален принтер формат).

Чита ја веќе обработената база (`data/products.xlsx`, со колоните Discount_Percent,
New_Price, Status кои ги пишува run_discount_check.py) и генерира PDF со по една
налепница на страница (40x30мм, термален принтер формат) за производите со активен
попуст (име, стара цена прецртана, нова цена, % попуст, баркод). Производите означени
"ЗА ПОВЛЕКУВАЊЕ" не добиваат налепница - тие само се тргнуваат од рафт, не им треба
нова цена.

Cowork го повикува main() директно (без argparse/input()).
"""

import os
import io

import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

try:
    import barcode
    from barcode.writer import ImageWriter
    HAS_BARCODE = True
except ImportError:
    HAS_BARCODE = False


CONFIG = {
    "input_file": os.path.join(os.path.dirname(__file__), "..", "data", "products.xlsx"),
    "output_pdf": os.path.join(os.path.dirname(__file__), "..", "data", "nalepnici.pdf"),
    # димензии на налепница - стандардна термална ролна за market/retail
    "label_width_mm": 40,
    "label_height_mm": 30,
    "font_regular": "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "font_bold": "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
}

FONT_NAME = "DejaVuSans"
FONT_BOLD = "DejaVuSans-Bold"


def register_fonts(config):
    pdfmetrics.registerFont(TTFont(FONT_NAME, config["font_regular"]))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, config["font_bold"]))


def wrap_text(text, font, size, line_widths_pt):
    """Го дели текстот во редови според листа на дозволени ширини (по еден елемент
    за секој ред - овозможува првиот ред да биде потесен, на пр. поради значка во
    аголот). Ако не стигнат сите зборови, последниот ред завршува со „…“."""
    words = str(text).split()
    max_lines = len(line_widths_pt)
    lines, current = [], ""
    idx = 0
    for word in words:
        max_width_pt = line_widths_pt[idx]
        trial = (current + " " + word).strip()
        if pdfmetrics.stringWidth(trial, font, size) <= max_width_pt:
            current = trial
        else:
            if current:
                lines.append(current)
                idx += 1
            current = word
        if idx == max_lines:
            break
    if current and idx < max_lines:
        lines.append(current)
    if len(lines) == max_lines:
        last_width = line_widths_pt[max_lines - 1]
        last = lines[-1]
        while pdfmetrics.stringWidth(last + "…", font, size) > last_width and len(last) > 1:
            last = last[:-1]
        used_words = len(" ".join(lines).split())
        if used_words < len(words):
            lines[-1] = last + "…"
    return lines[:max_lines]


def barcode_image(sku):
    """Генерира Code128 баркод како PIL слика во меморија. Враќа None ако либраријата
    не е достапна или SKU-то не може да се енкодира."""
    if not HAS_BARCODE:
        return None
    try:
        code = barcode.Code128(str(sku), writer=ImageWriter())
        buf = io.BytesIO()
        code.write(buf, options={"write_text": False, "module_height": 8.0, "quiet_zone": 1})
        buf.seek(0)
        from PIL import Image
        return Image.open(buf)
    except Exception:
        return None


def draw_discount_label(c, row, w, h):
    pad = 2.5 * mm
    name = str(row["Product"])
    old_price = float(row["Price"])
    new_price = float(row["New_Price"])
    pct = float(row["Discount_Percent"])
    sku = str(row["SKU"])

    # значка со попуст - горен десен агол (се пресметува прво за да се знае колку
    # простор одзема од првиот ред на името)
    badge_w, badge_h = 11 * mm, 5 * mm
    bx, by = w - pad - badge_w, h - pad - badge_h

    # производ - до 2 реда, мал фонт (првиот ред е потесен за да не се преклопи со значката)
    name_size = 7.5
    first_line_w = w - 2 * pad - badge_w - 1.5 * mm
    second_line_w = w - 2 * pad
    lines = wrap_text(name, FONT_NAME, name_size, [first_line_w, second_line_w])
    y = h - pad - name_size
    c.setFont(FONT_NAME, name_size)
    c.setFillColorRGB(0.15, 0.15, 0.15)
    for line in lines:
        c.drawString(pad, y, line)
        y -= name_size + 1.5

    c.setFillColorRGB(0.816, 0.231, 0.231)
    c.roundRect(bx, by, badge_w, badge_h, 1.2 * mm, stroke=0, fill=1)
    c.setFillColorRGB(1, 1, 1)
    c.setFont(FONT_BOLD, 8.5)
    c.drawCentredString(bx + badge_w / 2, by + 1.5, f"-{pct:.0f}%")

    # стара цена - прецртана
    old_y = y - 1
    c.setFont(FONT_NAME, 7)
    c.setFillColorRGB(0.45, 0.45, 0.45)
    old_txt = f"{old_price:.2f} ден."
    c.drawString(pad, old_y, old_txt)
    old_w = pdfmetrics.stringWidth(old_txt, FONT_NAME, 7)
    c.setLineWidth(0.6)
    c.line(pad, old_y + 2.3, pad + old_w, old_y + 2.3)

    # нова цена - голема, задебелена
    new_y = old_y - 10.5
    c.setFont(FONT_BOLD, 15)
    c.setFillColorRGB(0.816, 0.231, 0.231)
    c.drawString(pad, new_y, f"{new_price:.2f} ден.")

    # баркод
    img = barcode_image(sku)
    bc_h = 7 * mm
    bc_y = pad + 3
    if img is not None:
        bc_w = w - 2 * pad
        c.drawImage(ImageReader(img), pad, bc_y, width=bc_w, height=bc_h,
                    preserveAspectRatio=False, mask="auto")
        c.setFont(FONT_NAME, 5.5)
        c.setFillColorRGB(0.3, 0.3, 0.3)
        c.drawCentredString(w / 2, pad, sku)
    else:
        c.setFont(FONT_NAME, 6)
        c.setFillColorRGB(0.3, 0.3, 0.3)
        c.drawCentredString(w / 2, bc_y + bc_h / 2, sku)


def generate_labels(df, config):
    register_fonts(config)
    w = config["label_width_mm"] * mm
    h = config["label_height_mm"] * mm

    discounted = df[df["Discount_Percent"] > 0].copy()

    c = canvas.Canvas(config["output_pdf"], pagesize=(w, h))
    count = 0
    for _, row in discounted.iterrows():
        draw_discount_label(c, row, w, h)
        c.showPage()
        c.setPageSize((w, h))
        count += 1
    c.save()
    return {
        "count": count,
        "discount_labels": len(discounted),
        "path": config["output_pdf"],
    }


def main():
    print("Генерирам ценовни налепници...")
    try:
        df = pd.read_excel(CONFIG["input_file"])
    except Exception as e:
        print(f"Грешка при читање база: {e}")
        return

    required = {"Product", "SKU", "Price", "New_Price", "Discount_Percent", "Status"}
    missing = required - set(df.columns)
    if missing:
        print(f"Недостасуваат колони {missing}. Прво пушти го run_discount_check.py.")
        return

    if not HAS_BARCODE:
        print("Предупредување: 'python-barcode' не е инсталирано - налепниците ќе бидат "
              "без баркод слика (само SKU текст). Инсталирај со: "
              "pip install python-barcode --break-system-packages")

    result = generate_labels(df, CONFIG)
    print(
        f"Готово. Генерирани {result['count']} налепници со попуст во {result['path']}"
    )


if __name__ == "__main__":
    main()
