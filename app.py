# app.py
# Streamlit Rekonsiliasi Naik/Turun Golongan — Multi-file, CSV+XLS+XLSX, with Forced Engine Option

from __future__ import annotations

import io
import re
from typing import Iterable, Optional, Tuple, List

import numpy as np
import pandas as pd
import streamlit as st


# ---------- Capability detection ----------
def has_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False


def excel_reader_available() -> bool:
    return has_module("openpyxl") or has_module("pandas_calamine") or has_module("xlrd")


def excel_writer_available() -> bool:
    return has_module("xlsxwriter")


def available_reader_engines() -> List[str]:
    engines = []
    if has_module("openpyxl"):
        engines.append("openpyxl")
    if has_module("pandas_calamine"):
        engines.append("calamine")
    if has_module("xlrd"):
        engines.append("xlrd")
    return engines


XLSX_WRITER_OK = excel_writer_available()


# ---------- Loaders ----------
@st.cache_data(show_spinner=False)
def load_dataframe(file, forced_engine: str) -> pd.DataFrame:
    """
    Load single file; use forced_engine if compatible, else skip with warning.
    forced_engine in {"Auto", "openpyxl", "calamine", "xlrd"}.
    """
    if file is None:
        return pd.DataFrame()
    name = file.name
    low = name.lower()
    try:
        file.seek(0)
        if low.endswith(".csv"):
            return pd.read_csv(file, dtype=str, encoding_errors="ignore")

        # Excel branches
        if low.endswith(".xls"):
            # .xls only supported by xlrd
            if forced_engine != "Auto" and forced_engine != "xlrd":
                st.warning(f"Lewati `{name}`: format .xls membutuhkan engine `xlrd`, bukan `{forced_engine}`.")
                return pd.DataFrame()
            if not has_module("xlrd"):
                st.warning(f"Lewati `{name}`: engine `xlrd` tidak tersedia. Konversi ke CSV.")
                return pd.DataFrame()
            return pd.read_excel(file, dtype=str, engine="xlrd")

        if low.endswith(".xlsx") or low.endswith(".xlsm"):
            if forced_engine == "Auto":
                if has_module("openpyxl"):
                    return pd.read_excel(file, dtype=str, engine="openpyxl")
                if has_module("pandas_calamine"):
                    return pd.read_excel(file, dtype=str, engine="calamine")
                st.warning(f"Lewati `{name}`: tidak ada engine `openpyxl`/`calamine`. Konversi ke CSV.")
                return pd.DataFrame()
            # Forced engine
            if forced_engine == "openpyxl":
                if not has_module("openpyxl"):
                    st.warning(f"Lewati `{name}`: engine `openpyxl` tidak tersedia.")
                    return pd.DataFrame()
                return pd.read_excel(file, dtype=str, engine="openpyxl")
            if forced_engine == "calamine":
                if not has_module("pandas_calamine"):
                    st.warning(f"Lewati `{name}`: engine `calamine` tidak tersedia.")
                    return pd.DataFrame()
                return pd.read_excel(file, dtype=str, engine="calamine")
            if forced_engine == "xlrd":
                st.warning(f"Lewati `{name}`: `xlrd` tidak mendukung .xlsx/.xlsm. Pilih `openpyxl`/`calamine`.")
                return pd.DataFrame()

        # Fallback: coba CSV
        return pd.read_csv(file, dtype=str, encoding_errors="ignore")

    except Exception as e:
        st.error(f"Gagal membaca `{name}`: {e}")
        return pd.DataFrame()


def load_many(files: Optional[List], forced_engine: str) -> pd.DataFrame:
    if not files:
        return pd.DataFrame()
    frames: List[pd.DataFrame] = []
    for f in files:
        df = load_dataframe(f, forced_engine)
        if not df.empty:
            tmp = df.copy()
            tmp["Sumber File"] = f.name
            frames.append(tmp)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def parse_pasted_table(text: str) -> pd.DataFrame:
    text = (text or "").strip()
    if not text:
        return pd.DataFrame()
    if "\t" in text:
        sep = "\t"
    elif text.count(";") >= text.count(",") and ";" in text:
        sep = ";"
    elif "," in text:
        sep = ","
    else:
        sep = "|"
    try:
        return pd.read_csv(io.StringIO(text), dtype=str, sep=sep)
    except Exception:
        try:
            return pd.read_csv(io.StringIO(text), dtype=str, sep=None, engine="python")
        except Exception:
            st.error("Gagal mengurai data PASTE. Pastikan kolom dipisah TAB/;/,/|.")
            return pd.DataFrame()


# ---------- Helpers ----------
def guess_column(columns: Iterable[str], candidates: Iterable[str]) -> Optional[str]:
    cols = list(columns)
    norm = {c: normalize_colname(c) for c in cols}
    cand_norm = [normalize_colname(x) for x in candidates]
    for cn in cand_norm:
        for orig, nn in norm.items():
            if nn == cn:
                return orig
    for orig, nn in norm.items():
        if any(cn in nn for cn in cand_norm):
            return orig
    for orig, nn in norm.items():
        if any(nn.startswith(cn) or nn.endswith(cn) for cn in cand_norm):
            return orig
    return None


def normalize_colname(s: str) -> str:
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = (
        s.replace("no ", "nomor ")
        .replace("no", "nomor")
        .replace("inv", "invoice")
        .replace("harga total", "harga")
        .replace("nilai", "harga")
        .replace("tarip", "tarif")
    )
    return s


def coerce_invoice_key(series: pd.Series) -> pd.Series:
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r"\s+", "", regex=True)
    return s.str.upper()


def parse_money(value) -> float:
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    if s == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") and s.count("."):
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s and "." not in s:
            if re.search(r",[0-9]{1,2}$", s):
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "." in s and "," not in s:
            if re.search(r"\.[0-9]{1,2}$", s):
                pass
            else:
                s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        s2 = re.sub(r"[^\d\-]", "", s)
        try:
            return float(s2)
        except Exception:
            return 0.0


def coerce_money_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_money).astype(float)


def aggregate_by_invoice(df: pd.DataFrame, key_col: str, amount_col: str) -> pd.DataFrame:
    tmp = df.copy()
    tmp[key_col] = coerce_invoice_key(tmp[key_col])
    tmp[amount_col] = coerce_money_series(tmp[amount_col])
    g = tmp.groupby(key_col, dropna=False, as_index=False)[amount_col].sum()
    return g


def reconcile(
    inv_df: pd.DataFrame,
    inv_key: str,
    inv_amt: str,
    tik_df: pd.DataFrame,
    tik_key: str,
    tik_amt: str,
    only_diff: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    a = aggregate_by_invoice(inv_df, inv_key, inv_amt)
    b = aggregate_by_invoice(tik_df, tik_key, tik_amt)
    a = a.rename(columns={inv_key: "Nomor Invoice", inv_amt: "Nominal Invoice"})
    b = b.rename(columns={tik_key: "Nomor Invoice", tik_amt: "Nominal T-Summary"})
    merged = pd.merge(a, b, on="Nomor Invoice", how="outer")
    merged["Nominal Invoice"] = merged["Nominal Invoice"].fillna(0.0)
    merged["Nominal T-Summary"] = merged["Nominal T-Summary"].fillna(0.0)
    merged["Selisih"] = merged["Nominal Invoice"] - merged["Nominal T-Summary"]
    merged["Kategori"] = np.where(
        merged["Selisih"] > 0, "Naik", np.where(merged["Selisih"] < 0, "Turun", "Sama")
    )
    if only_diff:
        merged = merged.loc[merged["Selisih"] != 0]
    merged = merged.sort_values(["Kategori", "Nomor Invoice"], kind="stable")
    return a, b, merged


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Rekonsiliasi") -> Optional[bytes]:
    if not XLSX_WRITER_OK:
        return None
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def fmt_currency(x: float) -> str:
    if pd.isna(x):
        return ""
    n = float(x)
    s = f"{n:,.2f}"
    s = s.replace(",", "_").replace(".", ",").replace("_", ".")
    return s


def display_table(df: pd.DataFrame) -> None:
    if df.empty:
        st.warning("Tidak ada data untuk ditampilkan.")
        return
    display = df.copy()
    for col in ["Nominal Invoice", "Nominal T-Summary", "Selisih"]:
        if col in display.columns:
            display[col] = display[col].apply(fmt_currency)
    st.dataframe(display, use_container_width=True, hide_index=True)


# ---------- UI ----------
st.set_page_config(page_title="Rekonsiliasi Naik/Turun Golongan", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan")

with st.expander("ℹ️ Engine Excel"):
    avail = available_reader_engines()
    if avail:
        st.success("Engine tersedia: " + ", ".join(avail))
    else:
        st.error("Tidak ada engine Excel. File Excel akan di-skip. Gunakan CSV atau pasang engine.")
    forced_engine = st.selectbox(
        "Paksa engine Excel (opsional)",
        options=["Auto"] + avail,
        index=0,
        help="Auto: openpyxl → calamine untuk .xlsx/.xlsm; xlrd untuk .xls.",
    )

with st.sidebar:
    st.header("1) Upload File (Multiple)")
    f_inv_list = st.file_uploader(
        "📄 File Invoice — bisa banyak",
        type=["csv", "xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
    )
    f_tik_list = st.file_uploader(
        "🎫 File Tiket Summary — bisa banyak",
        type=["csv", "xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
    )
    st.caption("Nilai di-sum per Nomor Invoice, lalu dibandingkan.")

st.subheader("Opsional: Tempel Data dari Excel")
c1, c2 = st.columns(2)
with c1:
    paste_inv = st.text_area("PASTE — Invoice (TSV/CSV dari Excel)", height=160, placeholder="Tempel data Invoice di sini…")
with c2:
    paste_tik = st.text_area("PASTE — Tiket Summary (TSV/CSV dari Excel)", height=160, placeholder="Tempel data Tiket Summary di sini…")

# Compose sources
df_inv_files = load_many(f_inv_list, forced_engine)
df_inv_paste = parse_pasted_table(paste_inv)
if not df_inv_paste.empty:
    df_inv_paste["Sumber File"] = "PASTE:Invoice"
df_inv = pd.concat([df_inv_files, df_inv_paste], ignore_index=True, sort=False)

df_tik_files = load_many(f_tik_list, forced_engine)
df_tik_paste = parse_pasted_table(paste_tik)
if not df_tik_paste.empty:
    df_tik_paste["Sumber File"] = "PASTE:TiketSummary"
df_tik = pd.concat([df_tik_files, df_tik_paste], ignore_index=True, sort=False)

# Previews
if not df_inv.empty:
    st.subheader(f"Preview: Invoice (gabungan {len(df_inv)} baris)")
    st.dataframe(df_inv.head(10), use_container_width=True, hide_index=True)
if not df_tik.empty:
    st.subheader(f"Preview: Tiket Summary (gabungan {len(df_tik)} baris)")
    st.dataframe(df_tik.head(10), use_container_width=True, hide_index=True)

if df_inv.empty or df_tik.empty:
    st.info("Unggah minimal satu file (CSV/XLS/XLSX) atau tempel data untuk **Invoice** dan **Tiket Summary**.")
    st.stop()

st.divider()
st.subheader("2) Pemetaan Kolom")

invoice_key_guess = guess_column(
    df_inv.columns, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"]
)
invoice_amt_guess = guess_column(
    df_inv.columns, ["harga", "nilai", "amount", "nominal", "total", "grand total"]
)
tiket_key_guess = guess_column(
    df_tik.columns, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"]
)
tiket_amt_guess = guess_column(
    df_tik.columns, ["tarif", "harga", "nilai", "amount", "nominal", "total", "grand total"]
)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Invoice**")
    inv_key = st.selectbox(
        "Kolom Nomor Invoice (Invoice)",
        options=list(df_inv.columns),
        index=(list(df_inv.columns).index(invoice_key_guess) if invoice_key_guess in df_inv.columns else 0),
    )
    inv_amt = st.selectbox(
        "Kolom Nominal/Harga (Invoice)",
        options=list(df_inv.columns),
        index=(list(df_inv.columns).index(invoice_amt_guess) if invoice_amt_guess in df_inv.columns else 0),
    )
with col2:
    st.markdown("**Tiket Summary**")
    tik_key = st.selectbox(
        "Kolom Nomor Invoice (Tiket Summary)",
        options=list(df_tik.columns),
        index=(list(df_tik.columns).index(tiket_key_guess) if tiket_key_guess in df_tik.columns else 0),
    )
    tik_amt = st.selectbox(
        "Kolom Nominal/Tarif (Tiket Summary)",
        options=list(df_tik.columns),
        index=(list(df_tik.columns).index(tiket_amt_guess) if tiket_amt_guess in df_tik.columns else 0),
    )

st.divider()
st.subheader("3) Proses Rekonsiliasi")
only_diff = st.checkbox("Hanya tampilkan yang berbeda (Selisih ≠ 0)", value=False)
go = st.button("🚀 Proses")

if go:
    for df, need_cols, src in [
        (df_inv, [inv_key, inv_amt], "Invoice"),
        (df_tik, [tik_key, tik_amt], "Tiket Summary"),
    ]:
        for c in need_cols:
            if c not in df.columns:
                st.error(f"Kolom `{c}` tidak ditemukan di {src}")
                st.stop()

    agg_inv, agg_tik, merged = reconcile(df_inv, inv_key, inv_amt, df_tik, tik_key, tik_amt, only_diff)

    total_inv = float(agg_inv[agg_inv.columns[1]].sum()) if not agg_inv.empty else 0.0
    total_tik = float(agg_tik[agg_tik.columns[1]].sum()) if not agg_tik.empty else 0.0
    total_diff = float(merged["Selisih"].sum()) if not merged.empty else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nominal Invoice", fmt_currency(total_inv))
    m2.metric("Total Nominal T-Summary", fmt_currency(total_tik))
    m3.metric("Total Selisih (Invoice − T-Summary)", fmt_currency(total_diff))
    naik = int((merged["Kategori"] == "Naik").sum()) if not merged.empty else 0
    turun = int((merged["Kategori"] == "Turun").sum()) if not merged.empty else 0
    sama = int((merged["Kategori"] == "Sama").sum()) if not merged.empty else 0
    m4.metric("Naik / Turun / Sama", f"{naik} / {turun} / {sama}")

    st.subheader("Hasil Rekonsiliasi")
    display_table(merged)

    st.markdown("**Unduh Hasil**")
    csv_bytes = merged.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download CSV", data=csv_bytes, file_name="rekonsiliasi.csv", mime="text/csv")
    xlsx_bytes = df_to_excel_bytes(merged)
    if xlsx_bytes:
        st.download_button(
            "⬇️ Download Excel (XLSX)",
            data=xlsx_bytes,
            file_name="rekonsiliasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.caption("Excel writer tidak tersedia—gunakan CSV atau pasang paket `xlsxwriter`.")
    st.success("Selesai.")
