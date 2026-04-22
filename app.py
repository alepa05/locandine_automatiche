import os
import re
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

# =========================================
# CONFIG
# =========================================
TEMPLATE_PATH = "template/template_locandina.jpg"

FONT_DESC = "fonts/Montserrat-ExtraBold.ttf"
FONT_PRICE = "fonts/BebasNeue-Regular.ttf"
FONT_CODE = "fonts/Oswald-light.ttf"
FONT_FOOTER = "fonts/Montserrat-Bold.ttf"

OUTPUT_DIR = "output"

IMG_W = 2483
IMG_H = 3509

st.set_page_config(page_title="Generatore Locandine", layout="wide")
st.title("Generatore Locandine")


# =========================================
# FUNZIONI UTILI
# =========================================
def ensure_output_dir():
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def text_size(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_centered(draw, text, font, y, image_width, fill="black"):
    w, h = text_size(draw, text, font)
    x = (image_width - w) // 2
    draw.text((x, y), text, font=font, fill=fill)
    return x, y, w, h


def wrap_text_two_lines(draw, text, font_path, max_width, start_size, min_size=20):
    text = str(text).strip().upper()
    words = text.split()

    size = start_size
    while size >= min_size:
        font = ImageFont.truetype(font_path, size)

        w, _ = text_size(draw, text, font)
        if w <= max_width:
            return [text], font

        for i in range(1, len(words)):
            line1 = " ".join(words[:i])
            line2 = " ".join(words[i:])
            w1, _ = text_size(draw, line1, font)
            w2, _ = text_size(draw, line2, font)
            if w1 <= max_width and w2 <= max_width:
                return [line1, line2], font

        size -= 2

    font = ImageFont.truetype(font_path, min_size)
    return [text], font


def format_price(value):
    try:
        return f"{float(value):.2f}".replace(".", ",")
    except Exception:
        return str(value).replace(".", ",")


def format_date_it(value):
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    mesi = {
        1: "GENNAIO", 2: "FEBBRAIO", 3: "MARZO", 4: "APRILE",
        5: "MAGGIO", 6: "GIUGNO", 7: "LUGLIO", 8: "AGOSTO",
        9: "SETTEMBRE", 10: "OTTOBRE", 11: "NOVEMBRE", 12: "DICEMBRE"
    }

    if pd.notna(parsed):
        return f"{parsed.day} {mesi[parsed.month]}"

    return str(value).strip().upper()


# =========================================
# GENERAZIONE LOCANDINA
# =========================================
def generate_locandina(row):
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    descrizione = str(row["descrizione"]).strip().upper()

    # separa automaticamente la grammatura finale tipo (330 G), (1 KG), (50 PZ) ecc.
    gram = None
    match = re.search(r"\(([^()]+)\)\s*$", descrizione)
    if match:
       gram = match.group(1).strip()   # prende solo il contenuto interno, senza parentesi
       descrizione = descrizione[:match.start()].strip()
  
    prezzo = format_price(row["prezzo"])
    codice = str(row["codice_articolo"]).strip()
    data = format_date_it(row["scadenza_offerta"])

    RED = (236, 0, 19)
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)

    # =========================
    # DESCRIZIONE SU MAX 3 RIGHE A FONT FISSO
    # =========================
    MAX_WIDTH = 1800
    FONT_SIZE_DESC = 125
    MAX_LINES = 3

    font_desc = ImageFont.truetype(FONT_DESC, FONT_SIZE_DESC)

    words = descrizione.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = word if current_line == "" else current_line + " " + word
        w, _ = text_size(draw, test_line, font_desc)

        if w <= MAX_WIDTH:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    # se c'è la grammatura, una riga è riservata a quella
    max_desc_lines = MAX_LINES - 1 if gram else MAX_LINES

    # taglia eventuali righe in eccesso
    lines = lines[:max_desc_lines]

    _, line_height = text_size(draw, "TEST", font_desc)
    line_spacing = 40

    total_lines = len(lines) + (1 if gram else 0)
    total_height = total_lines * line_height + (total_lines - 1) * line_spacing
    start_y = 1650 - total_height // 2

    for line in lines:
        draw_centered(draw, line, font_desc, start_y, IMG_W, BLACK)
        start_y += line_height + line_spacing

    if gram:
        draw_centered(draw, gram, font_desc, start_y, IMG_W, BLACK)

    # =========================
    # PREZZO CON CENTRO FISSO
    # =========================
    numero, decimali = prezzo.split(",")

    FONT_SIZE = 1200
    PRICE_CENTER_X = 1235
    PRICE_Y = 1900

    font_price = ImageFont.truetype(FONT_PRICE, FONT_SIZE)

    num_w, num_h = text_size(draw, numero, font_price)
    dec_w, dec_h = text_size(draw, decimali, font_price)

    comma_overlap = 30
    comma_gap = 250

    total_w = (num_w - comma_overlap) + comma_gap + dec_w
    start_x = PRICE_CENTER_X - (total_w // 2)

    # numero principale
    draw.text((start_x, PRICE_Y), numero, font=font_price, fill=RED)

    # virgola
    comma_x = start_x + num_w - comma_overlap
    comma_y = PRICE_Y + 50
    draw.text((comma_x, comma_y), ",", font=font_price, fill=RED)

    # decimali
    dec_x = comma_x + comma_gap
    dec_y = PRICE_Y
    draw.text((dec_x, dec_y), decimali, font=font_price, fill=RED)

    # =========================
    # CODICE
    # =========================
    code_font = ImageFont.truetype(FONT_CODE, 165)
    code_x = comma_x + 400
    code_y = PRICE_Y + 1080
    draw.text((code_x, code_y), f"COD. {codice}", font=code_font, fill=RED)

    # =========================
    # FOOTER
    # =========================
    footer_font = ImageFont.truetype(FONT_FOOTER, 100)
    footer_text = f"OFFERTA VALIDA FINO AL {data}"
    draw_centered(draw, footer_text, footer_font, 3330, IMG_W, WHITE)

    # =========================
    # SAVE
    # =========================
    ensure_output_dir()
    path = os.path.join(OUTPUT_DIR, f"{codice}.jpg")
    img.save(path, quality=95)

    return path
    
# =========================================
# APP STREAMLIT
# =========================================
file = st.file_uploader("Carica Excel", type=["xlsx"])

if file:
    df = pd.read_excel(file, dtype={"codice_articolo": str})
    df["codice_articolo"] = df["codice_articolo"].str.strip()
    df.columns = [c.lower().strip() for c in df.columns]

    required = ["codice_articolo", "descrizione", "prezzo", "scadenza_offerta"]
    missing = [c for c in required if c not in df.columns]

    if missing:
        st.error("Mancano queste colonne nel file Excel: " + ", ".join(missing))
    else:
        st.dataframe(df[required])

        selected = []

        for i, row in df.iterrows():
            label = f"{row['codice_articolo']} - {row['descrizione']}"
            if st.checkbox(label, key=i):
                selected.append(i)

        if st.button("Genera Locandine"):
            for i in selected:
                generate_locandina(df.loc[i])

            st.success("Fatto!")