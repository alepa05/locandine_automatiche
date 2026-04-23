import os
import re
import io
import zipfile
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# =========================================
# CONFIG
# =========================================
TEMPLATE_PATH = "template/template_locandina.jpg"

FONT_DESC = "fonts/Montserrat-ExtraBold.ttf"
FONT_PRICE = "fonts/BebasNeue-Regular.ttf"
FONT_CODE = "fonts/Oswald-Light.ttf"
FONT_FOOTER = "fonts/Montserrat-Bold.ttf"

IMG_W = 2483
IMG_H = 3509

st.set_page_config(page_title="Generatore Locandine", layout="wide")
st.title("Generatore Locandine")


# =========================================
# FUNZIONI UTILI
# =========================================
def text_size(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_centered(draw, text, font, y, image_width, fill="black"):
    w, h = text_size(draw, text, font)
    x = (image_width - w) // 2
    draw.text((x, y), text, font=font, fill=fill)
    return x, y, w, h


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


def build_description_lines(draw, descrizione, font_path):
    MAX_WIDTH = 1800
    MAX_LINES = 3
    FONT_SIZE_DESC = 125

    words = descrizione.split()
    font_desc = ImageFont.truetype(font_path, FONT_SIZE_DESC)

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

    return lines[:MAX_LINES], font_desc


def generate_locandina_bytes(row):
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    draw = ImageDraw.Draw(img)

    descrizione = str(row["descrizione"]).strip().upper()

    # separa grammatura finale tra parentesi e la stampa sotto senza parentesi
    gram = None
    match = re.search(r"\(([^()]+)\)\s*$", descrizione)
    if match:
        gram = match.group(1).strip()
        descrizione = descrizione[:match.start()].strip()

    prezzo = format_price(row["prezzo"])
    codice = str(row["codice_articolo"]).strip().zfill(7)
    data = format_date_it(row["scadenza_offerta"])

    RED = (236, 0, 19)
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)

    # =========================
    # DESCRIZIONE
    # =========================
    lines, font_desc = build_description_lines(draw, descrizione, FONT_DESC)

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

    num_w, _ = text_size(draw, numero, font_price)
    dec_w, _ = text_size(draw, decimali, font_price)

    comma_overlap = 30
    comma_gap = 250

    total_w = (num_w - comma_overlap) + comma_gap + dec_w
    start_x = PRICE_CENTER_X - (total_w // 2)

    draw.text((start_x, PRICE_Y), numero, font=font_price, fill=RED)

    comma_x = start_x + num_w - comma_overlap
    comma_y = PRICE_Y + 50
    draw.text((comma_x, comma_y), ",", font=font_price, fill=RED)

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

    # salva in memoria
    img_bytes = io.BytesIO()
    img.save(img_bytes, format="JPEG", quality=95)
    img_bytes.seek(0)

    return codice, img_bytes


def build_zip_from_rows(df, selected_indices):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx in selected_indices:
            row = df.iloc[idx]
            codice, img_bytes = generate_locandina_bytes(row)

            safe_name = re.sub(r'[\\/*?:"<>|]', "", str(row["descrizione"]))
            safe_name = safe_name.replace(" ", "_")
            safe_name = safe_name[:80]

            zf.writestr(f"{safe_name}.jpg", img_bytes.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


def reset_selezione(df):
    for i in df.index:
        check_key = f"check_{i}"
        desc_key = f"desc_{i}"

        if check_key in st.session_state:
            st.session_state[check_key] = False

        if desc_key in st.session_state:
            del st.session_state[desc_key]


# =========================================
# APP STREAMLIT
# =========================================
file = st.file_uploader("Carica file Excel", type=["xlsx"])
st.caption("Carica un file Excel con questa struttura: codice_articolo - descrizione - prezzo - scadenza_offerta")

if file:
    df = pd.read_excel(file, dtype={"codice_articolo": str})
    df.columns = [c.lower().strip() for c in df.columns]
    df["codice_articolo"] = df["codice_articolo"].astype(str).str.strip().str.zfill(7)

    required = ["codice_articolo", "descrizione", "prezzo", "scadenza_offerta"]
    missing = [c for c in required if c not in df.columns]

    if missing:
        st.error("Mancano queste colonne nel file Excel: " + ", ".join(missing))
    else:
        left, center, right = st.columns([1, 2, 1])

        with center:
            st.subheader("Seleziona prodotti")

            search_code = st.text_input(
                "Cerca prodotto per codice",
                placeholder="Es. 0326542"
            ).strip()

            if search_code:
                df_filtered = df[df["codice_articolo"].str.contains(search_code, na=False)]
            else:
                df_filtered = df

            if df_filtered.empty:
                st.warning("Nessun prodotto trovato con questo codice.")
            else:
                selected_rows = []

                st.subheader("Seleziona prodotti e modifica descrizione")

                for i, row in df_filtered.iterrows():
                    label = f"{row['codice_articolo']} - {row['descrizione']}"
                    checked = st.checkbox(label, key=f"check_{i}")

                    if checked:
                        nuova_descrizione = st.text_input(
                            "Modifica descrizione",
                            value=str(row["descrizione"]),
                            key=f"desc_{i}"
                        )

                        selected_rows.append({
                            "index": i,
                            "descrizione_modificata": nuova_descrizione
                        })

                if st.button("Genera ZIP locandine"):
                    st.markdown("---")

                if st.button("Deseleziona articoli"):
                    reset_selezione(df)
                    st.rerun()
                    if not selected_rows:
                        st.warning("Seleziona almeno un prodotto.")
                    else:
                        righe_finali = []

                        for item in selected_rows:
                            row = df.loc[item["index"]].copy()
                            row["descrizione"] = item["descrizione_modificata"]
                            righe_finali.append(row)

                        zip_file = build_zip_from_rows(
                            pd.DataFrame(righe_finali).reset_index(drop=True),
                            range(len(righe_finali))
                        )

                        st.success(f"ZIP creato con {len(righe_finali)} locandine.")
                        today = datetime.now().strftime("%d-%m-%Y")

                        downloaded = st.download_button(
                            label="Scarica ZIP",
                            data=zip_file,
                            file_name=f"locandine_{today}.zip",
                            mime="application/zip"
                        )

                        if downloaded:
                            st.session_state.clear()
                            st.rerun()
