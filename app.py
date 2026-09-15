import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import csv

st.set_page_config(page_title="Poizon → CLIARTFATT CSV", page_icon="📄", layout="centered")

# Intestazioni IDENTICHE al CSV di importazione, compresi gli spazi iniziali.
HEADERS = [
    "TIPO_CF", " COD_CLI", " COD_CLI_XMAG", " RAG_SOCIALE", " PARTITA_IVA",
    " COD_FISCALE", " NAZIONE", " INDIRIZZO", " CAP", " CITTA", " PROVINCIA",
    " MAIL", " CELLULARE", " TELEFONO1", " FAX", " PEC", " INSTRAD_FATTELETT",
    " COD_SDI", " ALI_IVA", " COD_ART", " HSCODE", " DESCR_ART",
    " DESCR_ART_ESTESA", " UNITA_DI_MISURA", " COD_CLIENTE_DOCUMENTO",
    " COD_DOC", " SEZIONALE", " VALUTA", " DATA_DOC", " NUM_DOC",
    " PROGRESSIVO_RIGA", " INDICATORE_TIPORIGA", " COD_ART_DOC",
    " DESCRIZIONE_RIGA", " UM", " QUANTITA", " IVA", " PREZZO_1",
    " SCONTO1", " SCONTO2", " MAGGIORAZIONE1", " MAGGIORAZIONE2",
    " NOTE_MOVIMENTO", " COSTI_SPEDIZIONE", " EXSTRASCONTO"
]

# Valori fissi presi dal CSV di esempio.
DEFAULTS = {
    "TIPO_CF": "0",
    " COD_CLI": "89778",
    " COD_CLI_XMAG": "11146",
    " RAG_SOCIALE": "Poizon",
    " PARTITA_IVA": "12246280965",
    " COD_FISCALE": "",
    " NAZIONE": "86",
    " INDIRIZZO": "CORSO VERCELLI 40",
    " CAP": "20145",
    " CITTA": "Milano",
    " PROVINCIA": "",
    " MAIL": "",
    " CELLULARE": "",
    " TELEFONO1": "",
    " FAX": "",
    " PEC": "",
    " INSTRAD_FATTELETT": "FEPR",
    " COD_SDI": "BA6ET11",
    " ALI_IVA": "0",
    " HSCODE": "",
    " UNITA_DI_MISURA": "PZ",
    " COD_CLIENTE_DOCUMENTO": "89778",
    " COD_DOC": "FA",
    " SEZIONALE": "PZ",
    " VALUTA": "EURO",
    " INDICATORE_TIPORIGA": "0",
    " UM": "PZ",
    " IVA": "22",
    " SCONTO1": "0",
    " SCONTO2": "0",
    " MAGGIORAZIONE1": "0",
    " MAGGIORAZIONE2": "0",
    " COSTI_SPEDIZIONE": "0",
    " EXSTRASCONTO": "0",
}

REQUIRED_EXCEL_COLUMNS = [
    "Order No.",
    "Item Name",
    "Article Number/Style ID",
    "Quantity",
    "Item Price",
]


def clean_excel_value(value):
    """Converte un valore Excel in testo senza aggiungere .0 ai codici numerici."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_quantity(value):
    if pd.isna(value):
        return ""
    try:
        d = Decimal(str(value))
        if d == d.to_integral():
            return str(int(d))
        return format(d.normalize(), "f").replace(".", ",")
    except (InvalidOperation, ValueError):
        return str(value).strip().replace(".", ",")


def format_price_without_vat(value):
    """
    Il CSV di esempio conferma il calcolo imponibile = Item Price / 1,22.
    Il risultato viene scritto con 3 decimali e virgola decimale.
    Esempio: 872,00 / 1,22 = 714,754.
    """
    if pd.isna(value) or str(value).strip() == "":
        return ""

    raw = str(value).strip().replace("€", "").replace(" ", "")

    # Gestisce sia 872.00 sia 872,00.
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    else:
        raw = raw.replace(",", ".")

    try:
        gross = Decimal(raw)
    except InvalidOperation:
        raise ValueError(f"Prezzo non valido: {value}")

    net = (gross / Decimal("1.22")).quantize(
        Decimal("0.001"), rounding=ROUND_HALF_UP
    )
    return format(net, ".3f").replace(".", ",")


def read_statement_number(excel_bytes):
    overview = pd.read_excel(
        BytesIO(excel_bytes),
        sheet_name="Statement Overview",
        header=None,
        dtype=object,
    )

    for _, row in overview.iterrows():
        for col_idx, value in enumerate(row):
            if str(value).strip() == "Statement No.":
                # Normalmente il valore è nella cella immediatamente a destra.
                for next_idx in range(col_idx + 1, len(row)):
                    candidate = row.iloc[next_idx]
                    if not pd.isna(candidate) and str(candidate).strip():
                        return clean_excel_value(candidate)

    raise ValueError('Impossibile trovare "Statement No." nel foglio "Statement Overview".')


def read_settled_orders(excel_bytes):
    # Nel file Poizon ci sono tre righe di intestazione. La terza contiene
    # i nomi effettivi delle colonne.
    df = pd.read_excel(
        BytesIO(excel_bytes),
        sheet_name="Settled Orders",
        header=2,
        dtype=object,
    )

    missing = [c for c in REQUIRED_EXCEL_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            "Nel foglio 'Settled Orders' mancano queste colonne: "
            + ", ".join(missing)
        )

    # Elimina eventuali righe completamente vuote.
    df = df.dropna(how="all").copy()

    # Considera righe articolo solo quelle con Order No. valorizzato.
    df = df[df["Order No."].notna()].copy()
    return df


def build_csv(excel_bytes, num_doc):
    statement_no = read_statement_number(excel_bytes)
    orders = read_settled_orders(excel_bytes)

    if orders.empty:
        raise ValueError("Il foglio 'Settled Orders' non contiene ordini da esportare.")

    today = date.today().strftime("%d/%m/%Y")
    output_rows = []

    for progressivo, (_, src) in enumerate(orders.iterrows(), start=1):
        row = {header: DEFAULTS.get(header, "") for header in HEADERS}

        order_no = clean_excel_value(src["Order No."])
        item_name = clean_excel_value(src["Item Name"])
        article = clean_excel_value(src["Article Number/Style ID"])
        description = f"{order_no} - {item_name}"

        row[" COD_ART"] = article
        row[" DATA_DOC"] = today
        row[" NUM_DOC"] = str(num_doc).strip()
        row[" PROGRESSIVO_RIGA"] = str(progressivo)
        row[" COD_ART_DOC"] = article
        row[" DESCR_ART"] = description
        row[" DESCR_ART_ESTESA"] = description
        row[" DESCRIZIONE_RIGA"] = description
        row[" QUANTITA"] = format_quantity(src["Quantity"])
        row[" PREZZO_1"] = format_price_without_vat(src["Item Price"])
        row[" NOTE_MOVIMENTO"] = f"Statement No. {statement_no}"

        output_rows.append(row)

    # StringIO con newline="" + lineterminator CRLF mantiene il formato del CSV esempio.
    sio = StringIO(newline="")
    writer = csv.DictWriter(
        sio,
        fieldnames=HEADERS,
        delimiter=";",
        quotechar='"',
        quoting=csv.QUOTE_MINIMAL,
        lineterminator="\r\n",
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(output_rows)

    # Il CSV originale è UTF-8 senza BOM.
    return sio.getvalue().encode("utf-8"), output_rows, statement_no


st.title("Poizon Statement → CLIARTFATT.csv")
st.caption("Carica l'Excel dello statement Poizon e genera il CSV pronto per l'importazione.")

uploaded_file = st.file_uploader(
    "File Excel Poizon",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
)

num_doc = st.text_input(
    "PROGRESSIVO",
    placeholder="Es. 123",
    help="Numero documento da riportare su tutte le righe del CSV.",
)

if uploaded_file is not None:
    excel_bytes = uploaded_file.getvalue()

    try:
        # Piccolo controllo immediato sul file.
        orders_preview = read_settled_orders(excel_bytes)
        statement_preview = read_statement_number(excel_bytes)
        st.success(
            f"File letto correttamente: {len(orders_preview)} righe in Settled Orders. "
            f"Statement No. {statement_preview}"
        )
    except Exception as e:
        st.error(str(e))
        st.stop()

    if st.button("Genera CSV", type="primary", use_container_width=True):
        if not num_doc.strip():
            st.error("Inserisci NUM_DOC prima di generare il CSV.")
        else:
            try:
                csv_bytes, output_rows, statement_no = build_csv(excel_bytes, num_doc)

                st.success(
                    f"CSV generato: {len(output_rows)} righe articolo. "
                    f"Statement No. {statement_no}"
                )

                preview_df = pd.DataFrame(output_rows, columns=HEADERS)
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

                st.download_button(
                    "Scarica CLIARTFATT.csv",
                    data=csv_bytes,
                    file_name="CLIARTFATT.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"Errore durante la generazione: {e}")

