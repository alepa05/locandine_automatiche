import os
import re
import io
import zipfile
import base64
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

TEMPLATE_PATH = "template/template_locandina.jpg"

FONT_DESC = "fonts/Montserrat-ExtraBold.ttf"
FONT_PRICE = "fonts/BebasNeue-Regular.ttf"
FONT_CODE = "fonts/Oswald-Light.ttf"
FONT_FOOTER = "fonts/Montserrat-Bold.ttf"

IMG_W = 2483
IMG_H = 3509

st.set_page_config(page_title="Generatore Locandine", layout="wide")
st.title("Generatore Locandine")

st.markdown("""
<style>
.block-container {
    max-width: 1150px;
    padding-top: 2rem;
}

div[data-testid="stVerticalBlock"] {
    gap: 1.1rem;
}

.stButton button {
    border-radius: 8px;
    height: 42px;
}

.stDownloadButton button {
    border-radius: 8px;
    height: 48px;
    font-weight: 700;
}

[data-testid="stVerticalBlock"] div::-webkit-scrollbar {
    width: 16px;
}

[data-testid="stVerticalBlock"] div::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 10px;
}

[data-testid="stVerticalBlock"] div::-webkit-scrollbar-thumb {
    background: #888;
    border-radius: 2px;
}

[data-testid="stVerticalBlock"] div::-webkit-scrollbar-thumb:hover {
    background: #555;
}
</style>
""", unsafe_allow_html=True)


def text_size(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def draw_centered(draw, text, font, y, image_width, fill="black"):
    w, h = text_size(draw, text, font)
    x = (image_width - w) // 2
    draw.text((x, y), text, font=font, fill=fill)
    return x, y, w, h


def normalizza_colonna(nome):
    nome = str(nome).lower().strip()
    nome = nome.replace(" ", "_").replace("-", "_").replace(".", "_")
    nome = re.sub(r"_+", "_", nome)
    return nome


def trova_colonna(df, possibili_nomi):
    colonne = list(df.columns)

    for nome in possibili_nomi:
        if nome in colonne:
            return nome

    for col in colonne:
        for nome in possibili_nomi:
            if nome in col:
                return col

    return None


def sistema_colonne_excel(df):
    df.columns = [normalizza_colonna(c) for c in df.columns]

    mapping = {
        "codice_articolo": [
            "codice_articolo", "codice", "cod_articolo", "cod_art",
            "codice_prodotto", "codiceprodotto", "articolo", "sku", "ean", "cod"
        ],
        "descrizione": [
            "descrizione", "desc", "descrizione_articolo", "nome",
            "nome_prodotto", "prodotto", "articolo_descrizione"
        ],
        "prezzo": [
            "prezzo", "price", "prezzo_offerta", "prezzo_vendita",
            "offerta", "prezzo_promo", "promo"
        ],
        "scadenza_offerta": [
            "scadenza_offerta", "scadenza", "data_scadenza",
            "valido_fino", "validita", "fine_offerta", "data_fine", "fino_al"
        ]
    }

    nuove_colonne = {}

    for colonna_standard, sinonimi in mapping.items():
        trovata = trova_colonna(df, sinonimi)
        if trovata:
            nuove_colonne[trovata] = colonna_standard

    return df.rename(columns=nuove_colonne)


def leggi_excel_auto(file):
    excel_preview = pd.read_excel(file, header=None, dtype=str)

    header_row = 0

    for i in range(min(20, len(excel_preview))):
        row_values = (
            excel_preview.iloc[i]
            .fillna("")
            .astype(str)
            .str.lower()
            .tolist()
        )

        row_text = " ".join(row_values)

        if (
            "cod" in row_text
            or "descr" in row_text
            or "prezzo" in row_text
            or "scadenza" in row_text
            or "valido" in row_text
        ):
            header_row = i
            break

    file.seek(0)

    df = pd.read_excel(file, dtype=str, header=header_row)
    df = sistema_colonne_excel(df)

    return df


def format_price(value):
    try:
        return f"{float(str(value).replace(',', '.')):.2f}".replace(".", ",")
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


def separa_descrizione_grammatura(descrizione):
    descrizione = str(descrizione).strip().upper()

    match = re.search(r"\(([^()]+)\)\s*$", descrizione)
    if match:
        gram = match.group(1).strip()
        desc = descrizione[:match.start()].strip()
        return desc, gram

    pattern = r"\b((?:AL\s+)?(?:KG|G|GR|LT|L|ML|CL|PZ|PEZZI|CONF|CF|X\s*\d+|[0-9]+(?:[,.][0-9]+)?\s*(?:KG|G|GR|LT|L|ML|CL|PZ|PEZZI|CONF|CF)(?:\s*X\s*\d+)?|[0-9]+/[0-9]+\s*(?:KG|G|GR|LT|L|ML|CL)))$"

    match = re.search(pattern, descrizione)
    if match:
        gram = match.group(1).strip()
        desc = descrizione[:match.start()].strip()
        return desc, gram

    return descrizione, None


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

    descrizione_originale = str(row["descrizione"]).strip().upper()
    descrizione, gram = separa_descrizione_grammatura(descrizione_originale)

    prezzo = format_price(row["prezzo"])
    codice = str(row["codice_articolo"]).strip().zfill(7)
    data = format_date_it(row["scadenza_offerta"])

    RED = (236, 0, 19)
    BLACK = (0, 0, 0)
    WHITE = (255, 255, 255)

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
    draw.text((dec_x, PRICE_Y), decimali, font=font_price, fill=RED)

    code_font = ImageFont.truetype(FONT_CODE, 165)
    code_x = comma_x + 400
    code_y = PRICE_Y + 1080
    draw.text((code_x, code_y), f"COD. {codice}", font=code_font, fill=RED)

    footer_font = ImageFont.truetype(FONT_FOOTER, 100)
    footer_text = f"OFFERTA VALIDA FINO AL {data}"
    draw_centered(draw, footer_text, footer_font, 3330, IMG_W, WHITE)

    img_bytes = io.BytesIO()
    img.save(img_bytes, format="JPEG", quality=95)
    img_bytes.seek(0)

    return codice, img_bytes


def build_zip_from_rows(df, selected_indices, status_text=None):
    zip_buffer = io.BytesIO()
    total = len(selected_indices)
    used_names = {}

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for count, idx in enumerate(selected_indices, start=1):
            row = df.iloc[idx]
            codice, img_bytes = generate_locandina_bytes(row)

            safe_name = re.sub(r'[\\/*?:"<>|]', "", str(row["descrizione"]))
            safe_name = safe_name.replace(" ", "_")
            safe_name = safe_name[:80]

            base_filename = safe_name
            filename = f"{base_filename}.jpg"

            if filename in used_names:
                used_names[filename] += 1
                filename = f"{base_filename}_{used_names[filename]}.jpg"
            else:
                used_names[filename] = 1

            zf.writestr(filename, img_bytes.getvalue())

            if status_text is not None:
                percent = int((count / total) * 100)
                status_text.info(f"Generazione locandine in corso... {percent}%")

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


def seleziona_tutto(df):
    for i in df.index:
        st.session_state[f"check_{i}"] = True


def auto_download_zip(zip_bytes, filename):
    b64 = base64.b64encode(zip_bytes).decode()

    components.html(
        f"""
        <html>
            <body>
                <script>
                    const link = document.createElement('a');
                    link.href = "data:application/zip;base64,{b64}";
                    link.download = "{filename}";
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                </script>
            </body>
        </html>
        """,
        height=0
    )


with st.container(border=True):
    st.subheader("1. Carica file Excel")

    file = st.file_uploader("Carica file Excel", type=["xlsx"])

    st.caption(
        "Carica un file Excel contenente almeno queste informazioni: "
        "codice articolo, descrizione, e prezzo. "
    )


if file:
    df = leggi_excel_auto(file)

    required_base = [
        "codice_articolo",
        "descrizione",
        "prezzo"
    ]

    missing_base = [c for c in required_base if c not in df.columns]

    if missing_base:
        with st.container(border=True):
            st.error(
                "Non riesco a riconoscere queste colonne obbligatorie: "
                + ", ".join(missing_base)
            )

            st.info(
                "Rinomina le colonne del file Excel oppure usa nomi simili a: "
                "codice_articolo, descrizione, prezzo."
            )

            st.write("Colonne trovate nel file:")
            st.write(list(df.columns))

    else:
        with st.container(border=True):
            st.subheader("2. Scadenza volantino")

            if "scadenza_offerta" not in df.columns:
                left_date, center_date, right_date = st.columns([1, 2, 1])

                with center_date:
                    st.info("Inserisci la data di scadenza del volantino.")

                    mesi_select = [
                        "GENNAIO", "FEBBRAIO", "MARZO", "APRILE",
                        "MAGGIO", "GIUGNO", "LUGLIO", "AGOSTO",
                        "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"
                    ]

                    col_giorno, col_mese = st.columns(2)

                    with col_giorno:
                        giorno_scadenza = st.selectbox(
                            "Giorno",
                            list(range(1, 32)),
                            index=0
                        )

                    with col_mese:
                        mese_scadenza = st.selectbox(
                            "Mese",
                            mesi_select,
                            index=0
                        )

                df["scadenza_offerta"] = f"{giorno_scadenza} {mese_scadenza}"

            else:
                st.success("Scadenza offerta rilevata automaticamente dal file.")

        required = [
            "codice_articolo",
            "descrizione",
            "prezzo",
            "scadenza_offerta"
        ]

        df = df[required].copy()

        df = df.dropna(
            subset=["codice_articolo", "descrizione", "prezzo"],
            how="all"
        )

        df["codice_articolo"] = (
            df["codice_articolo"]
            .astype(str)
            .str.strip()
            .str.replace(".0", "", regex=False)
            .str.zfill(7)
        )

        with st.container(border=True):
            st.subheader("3. Ricerca e selezione prodotti")

            search_code = st.text_input(
                "Cerca prodotto per codice",
                placeholder="Es. 0326542"
            ).strip()

            if search_code:
                df_filtered = df[
                    df["codice_articolo"].str.contains(search_code, na=False)
                ]
            else:
                df_filtered = df

            if df_filtered.empty:
                st.warning("Nessun prodotto trovato con questo codice.")

            else:
                col1, col2 = st.columns(2)

                with col1:
                    if st.button("Seleziona tutto", use_container_width=True):
                        seleziona_tutto(df_filtered)
                        st.rerun()

                with col2:
                    if st.button("Deseleziona articoli", use_container_width=True):
                        reset_selezione(df)
                        st.rerun()

                selected_rows = []

                st.markdown("### Seleziona prodotti e modifica descrizione")

                with st.container(height=600):
                    for i, row in df_filtered.iterrows():
                        label = (
                            f"{row['codice_articolo']} - "
                            f"{row['descrizione']}"
                        )

                        checked = st.checkbox(
                            label,
                            key=f"check_{i}"
                        )

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

        with st.container(border=True):
            st.subheader("4. Genera locandine")
            st.markdown(
            f"<p style='font-size:14px;color:gray;'>File selezionati: <b>{len(selected_rows)}</b></p>",
            unsafe_allow_html=True
        )

            if st.button(
                "Genera e scarica ZIP locandine",
                use_container_width=True
            ):
                if not selected_rows:
                    st.warning("Seleziona almeno un prodotto.")

                else:
                    righe_finali = []

                    for item in selected_rows:
                        row = df.loc[item["index"]].copy()
                        row["descrizione"] = item["descrizione_modificata"]
                        righe_finali.append(row)

                    status_text = st.empty()

                    zip_file = build_zip_from_rows(
                        pd.DataFrame(righe_finali).reset_index(drop=True),
                        range(len(righe_finali)),
                        status_text=status_text
                    )

                    today = datetime.now().strftime("%d-%m-%Y")
                    filename = f"locandine_{today}.zip"

                    status_text.success("File pronto. Download in avvio...")

                    auto_download_zip(zip_file.getvalue(), filename)
