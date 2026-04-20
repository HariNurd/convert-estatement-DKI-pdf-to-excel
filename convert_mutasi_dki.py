import re
from pathlib import Path

import pandas as pd
import tabula


# =========================
# PDF EXTRACTION CONFIG
# =========================
PDF_AREA = [200, 15, 710, 830]
PDF_COLUMNS = [90, 330, 390, 500, 830]
# split into:
# TanggalJam | Keterangan | Debit/kredit | Nominal | Saldo_berjalan

PDF_AREA_SUMMARY = [290, 15, 710, 830]
PDF_COLUMNS_SUMMARY = [180, 300, 370, 500, 830]
# split into:
# Rekening | Saldo_Awal | Transaksi_Masuk | Transaksi_Keluar | Saldo_Akhir


# =========================
# HELPER FUNCTIONS
# =========================
def clean_text(value):
    """Normalize text: remove line breaks, trim spaces, collapse multiple spaces."""
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def parse_number(value):
    """
    Convert:
    - '4,060,274.00' -> 4060274
    - '119,954.18'   -> 119954.18
    """
    text = str(value).strip().upper()

    if text in ("", "NAN"):
        return pd.NA

    text = text.replace(",", "").strip()

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except ValueError:
        return pd.NA


def is_datetime_start(text):
    """
    Cek apakah text dimulai dengan format seperti:
    06 Mar 26 16:20
    """
    text = str(text).strip()
    return bool(re.match(r"^\d{2}\s+[A-Za-z]{3}\s+\d{2}\s+\d{2}:\d{2}$", text))


def append_text(base, extra):
    """Append text with single space separation."""
    base = (base or "").strip()
    extra = (extra or "").strip()

    if not extra:
        return base
    if not base:
        return extra
    return f"{base} {extra}"


def parse_mutation_type(jenis, nominal):
    """
    Pisahkan nominal ke Debit / Kredit
    """
    jenis_text = str(jenis).strip().upper()
    number = parse_number(nominal)

    if pd.isna(number):
        return pd.Series([pd.NA, pd.NA], index=["Debit", "Kredit"])

    if "DEBIT" in jenis_text:
        return pd.Series([number, pd.NA], index=["Debit", "Kredit"])

    if "KREDIT" in jenis_text:
        return pd.Series([pd.NA, number], index=["Debit", "Kredit"])

    return pd.Series([pd.NA, pd.NA], index=["Debit", "Kredit"])


# =========================
# PDF READING FUNCTIONS
# =========================
def read_pdf_table(pdf_file):
    """Read all pages from PDF using Tabula and combine them into one dataframe."""
    dfs = tabula.read_pdf(
        str(pdf_file),
        pages="all",
        stream=True,
        guess=False,
        area=PDF_AREA,
        columns=PDF_COLUMNS,
        pandas_options={"header": None},
        multiple_tables=False,
    )

    if not dfs:
        raise ValueError("No table found. Try adjusting area/columns slightly.")

    print(f"Proses membaca PDF selesai, ditemukan {len(dfs)} potongan tabel. Menggabungkan data...")

    df = pd.concat(dfs, ignore_index=True)
    df = df.iloc[:, :5].copy()
    df.columns = ["TanggalJam", "Keterangan", "Debit/kredit", "Nominal", "Saldo_berjalan"]

    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    return df


def read_summary_page(pdf_file):
    """
    Ambil summary dari halaman 1.
    Kolom hasil:
    Rekening | Saldo_Awal | Transaksi_Masuk | Transaksi_Keluar | Saldo_Akhir
    """
    dfs = tabula.read_pdf(
        str(pdf_file),
        pages=1,
        stream=True,
        guess=False,
        area=PDF_AREA_SUMMARY,
        columns=PDF_COLUMNS_SUMMARY,
        pandas_options={"header": None},
        multiple_tables=True,
    )

    if not dfs:
        return pd.DataFrame(
            columns=[
                "Rekening",
                "Saldo_Awal",
                "Transaksi_Masuk",
                "Transaksi_Keluar",
                "Saldo_Akhir",
            ]
        )

    df = pd.concat(dfs, ignore_index=True)
    df = df.iloc[:, :5].copy()
    df.columns = ["Rekening", "Saldo_Awal", "Transaksi_Masuk", "Transaksi_Keluar", "Saldo_Akhir"]

    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    return df


# =========================
# DATA CLEANING FUNCTIONS
# =========================
def merge_summary_rows(df_summary):
    """
    Gabungkan baris summary yang kepotong.
    Contoh:
    - TABUNGAN MONAS PEGAWAI
    - DKI - 10123893782
    menjadi satu kolom Rekening.
    """
    df_summary = df_summary.copy()

    if df_summary.empty:
        return pd.DataFrame(
            columns=[
                "Rekening",
                "Saldo_Awal",
                "Transaksi_Masuk",
                "Transaksi_Keluar",
                "Saldo_Akhir",
            ]
        )

    for col in df_summary.columns:
        df_summary[col] = df_summary[col].apply(clean_text)

    first_row = df_summary.iloc[0].to_dict()

    rekening_parts = [first_row["Rekening"]]

    for i in range(1, len(df_summary)):
        extra_rekening = df_summary.loc[i, "Rekening"]
        if extra_rekening:
            rekening_parts.append(extra_rekening)

    first_row["Rekening"] = " ".join(part for part in rekening_parts if part).strip()
    first_row["Saldo_Awal"] = parse_number(first_row["Saldo_Awal"])
    first_row["Transaksi_Masuk"] = parse_number(first_row["Transaksi_Masuk"])
    first_row["Transaksi_Keluar"] = parse_number(first_row["Transaksi_Keluar"])
    first_row["Saldo_Akhir"] = parse_number(first_row["Saldo_Akhir"])

    return pd.DataFrame([first_row])


def merge_continuation_rows(df):
    """
    Jika baris tidak punya TanggalJam, anggap lanjutan dari transaksi sebelumnya.
    Cocok untuk kasus seperti:
    26 Mar 26 15:38   Dr 30091007495 8548/00749/304866 SUSPECT
                      WDW 202603
                      KREDIT 500,000.00 875,050.18
    """
    merged_rows = []
    current = None

    for _, row in df.iterrows():
        tanggal_jam = str(row["TanggalJam"]).strip()
        rincian = str(row["Keterangan"]).strip()
        jenis = str(row["Debit/kredit"]).strip()
        nominal = str(row["Nominal"]).strip()
        saldo = str(row["Saldo_berjalan"]).strip()

        if is_datetime_start(tanggal_jam):
            if current is not None:
                merged_rows.append(current)

            current = {
                "TanggalJam": tanggal_jam,
                "Rincian": rincian,
                "Jenis": jenis,
                "Nominal": nominal,
                "Saldo": saldo,
            }
        else:
            if current is None:
                continue

            if tanggal_jam:
                current["Rincian"] = append_text(current["Rincian"], tanggal_jam)

            if rincian:
                current["Rincian"] = append_text(current["Rincian"], rincian)

            if jenis:
                if current["Jenis"] == "":
                    current["Jenis"] = jenis
                else:
                    current["Rincian"] = append_text(current["Rincian"], jenis)

            if nominal:
                if current["Nominal"] == "":
                    current["Nominal"] = nominal
                else:
                    current["Rincian"] = append_text(current["Rincian"], nominal)

            if saldo:
                if current["Saldo"] == "":
                    current["Saldo"] = saldo
                else:
                    current["Rincian"] = append_text(current["Rincian"], saldo)

    if current is not None:
        merged_rows.append(current)

    return pd.DataFrame(merged_rows, columns=["TanggalJam", "Rincian", "Jenis", "Nominal", "Saldo"])


def finalize_transactions(df):
    """Split debit/kredit, parse saldo, lalu pecah tanggal dan jam."""
    df[["Debit", "Kredit"]] = df.apply(
        lambda row: parse_mutation_type(row["Jenis"], row["Nominal"]),
        axis=1
    )

    df["Saldo"] = df["Saldo"].apply(parse_number)

    df["Tanggal"] = df["TanggalJam"].str.extract(r"^(\d{2}\s+[A-Za-z]{3}\s+\d{2})", expand=False)
    df["Jam"] = df["TanggalJam"].str.extract(r"(\d{2}:\d{2})$", expand=False)

    df = df[["Tanggal", "Jam", "Rincian", "Debit", "Kredit", "Saldo"]].copy()

    return df


# =========================
# EXCEL EXPORT
# =========================
def auto_fit_columns(sheet):
    """Auto-fit Excel column widths based on content length."""
    for col_cells in sheet.columns:
        col_letter = col_cells[0].column_letter
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in col_cells
        )
        sheet.column_dimensions[col_letter].width = max_length + 2


def export_to_excel(clean_df, df_summary_final, output_file):
    """Write transaction and summary dataframes to Excel."""
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        clean_df.to_excel(writer, sheet_name="Transaksi", index=False)
        df_summary_final.to_excel(writer, sheet_name="Summary", index=False)

        sheet1 = writer.sheets["Transaksi"]
        sheet2 = writer.sheets["Summary"]

        auto_fit_columns(sheet1)
        auto_fit_columns(sheet2)


# =========================
# MAIN
# =========================
def main():
    file_name = input("Enter the filename without extension: ").strip()

    pdf_dir = Path("pdf_file")
    excel_dir = Path("excel_file")
    excel_dir.mkdir(parents=True, exist_ok=True)

    pdf_file = pdf_dir / f"{file_name}.pdf"
    output_file = excel_dir / f"{file_name}.xlsx"

    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")

    print("Membaca summary halaman 1...")
    df_summary_raw = read_summary_page(pdf_file)

    print("Membaca detail transaksi halaman 2 sampai akhir...")
    df = read_pdf_table(pdf_file)

    mask = df["TanggalJam"].apply(lambda x: is_datetime_start(x) or str(x).strip() == "")
    df = df[mask].reset_index(drop=True)

    df_merged = merge_continuation_rows(df)
    df_summary = merge_summary_rows(df_summary_raw)

    print("Membersihkan dan memfinalisasi data transaksi...")
    df_final = finalize_transactions(df_merged)

    print("Menulis hasil ke Excel...")
    export_to_excel(df_final, df_summary, output_file)

    print(f"Data successfully written to {output_file}")


if __name__ == "__main__":
    main()