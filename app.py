# app.py
# Streamlit Rekonsiliasi Naik/Turun Golongan berdasarkan Nomor Invoice

from __future__ import annotations

import io
import re
from typing import Iterable, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st


# ---------- Util ----------
@st.cache_data(show_spinner=False)
def load_dataframe(file) -> pd.DataFrame:
    """Load CSV/XLSX into DataFrame. Fallback to CSV sniffing."""
    if file is None:
        return pd.DataFrame()
    name = file.name.lower()
    try:
        if name.endswith(".csv"):
            return pd.read_csv(file, dtype=str, encoding_errors="ignore")
        if name.endswith(".xlsx") or name.endswith(".xls"):
            # engine picked automatically; if missing, except will be caught
            return pd.read_excel(file, dtype=str)
        # default try CSV
        file.seek(0)
        return pd.read_csv(file, dtype=str, encoding_errors="ignore")
    except Exception as e:
        st.error(f"Gagal membaca file `{file.name}`: {e}")
        return pd.DataFrame()


def guess_column(
    columns: Iterable[str], candidates: Iterable[str]
) -> Optional[str]:
    """Pick best-matching column name by fuzzy/alias."""
    cols = list(columns)
    norm = {c: normalize_colname(c) for c in cols}
    cand_norm = [normalize_colname(x) for x in candidates]
    # exact normalized match
    for cn in cand_norm:
        for orig, nn in norm.items():
            if nn == cn:
                return orig
    # contains check
    for orig, nn in norm.items():
        if any(cn in nn for cn in cand_norm):
            return orig
    # prefix/suffix heuristic
    for orig, nn in norm.items():
        if any(nn.startswith(cn) or nn.endswith(cn) for cn in cand_norm):
            return orig
    return None


def normalize_colname(s: str) -> str:
    s = s.lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    # Indonesian aliases
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
    """Normalize invoice IDs to comparable string keys."""
    s = series.fillna("").astype(str).str.strip()
    s = s.str.replace(r"\s+", "", regex=True)  # remove internal spaces
    return s.str.upper()


def parse_money(value) -> float:
    """
    Robust parse for Indonesian/Intl formats.
    Keeps sign, removes currency & spaces. Decimal inference:
    - if both ',' and '.' exist → last separator is decimal.
    - else ',' with 1-2 digits at end → decimal; otherwise thousands.
    - else '.' with 1-2 digits at end → decimal; otherwise thousands.
    """
    if pd.isna(value):
        return 0.0
    s = str(value).strip()
    if s == "":
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") and s.count("."):
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
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
        # last resort: strip non-digits except minus
        s2 = re.sub(r"[^\d\-]", "", s)
        try:
            return float(s2)
        except Exception:
            return 0.0


def coerce_money_series(series: pd.Series) -> pd.Series:
    return series.apply(parse_money).astype(float)


def aggregate_by_invoice(
    df: pd.DataFrame, key_col: str, amount_col: str
) -> pd.DataFrame:
    """Group by invoice key and sum amount."""
    out = pd.DataFrame()
    if df.empty:
        return out
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
    """Return (agg_invoice, agg_tiket, merged_result)."""
    a = aggregate_by_invoice(inv_df, inv_key, inv_amt)
    b = aggregate_by_invoice(tik_df, tik_key, tik_amt)
    a = a.rename(columns={inv_key: "Nomor Invoice", inv_amt: "Nominal Invoice"})
    b = b.rename(columns={tik_key: "Nomor Invoice", tik_amt: "Nominal T-Summary"})
    merged = pd.merge(a, b, on="Nomor Invoice", how="outer")
    merged["Nominal Invoice"] = merged["Nominal Invoice"].fillna(0.0)
    merged["Nominal T-Summary"] = merged["Nominal T-Summary"].fillna(0.0)
    merged["Selisih"] = merged["Nominal Invoice"] - merged["Nominal T-Summary"]
    merged["Kategori"] = np.where(
        merged["Selisih"] > 0,
        "Naik",
        np.where(merged["Selisih"] < 0, "Turun", "Sama"),
    )
    if only_diff:
        merged = merged.loc[merged["Selisih"] != 0]
    merged = merged.sort_values(["Kategori", "Nomor Invoice"], kind="stable")
    return a, b, merged


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Rekonsiliasi") -> Optional[bytes]:
    """Try to produce an XLSX; fallback None if engine missing."""
    try:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception:
        return None


def fmt_currency(x: float) -> str:
    """Format as IDR without symbol; thousands sep dot, decimal comma."""
    if pd.isna(x):
        return ""
    n = float(x)
    s = f"{n:,.2f}"  # 1,234,567.89
    s = s.replace(",", "_").replace(".", ",").replace("_", ".")
    return s


def display_table(df: pd.DataFrame) -> None:
    """Display numeric columns nicely formatted."""
    if df.empty:
        st.warning("Tidak ada data untuk ditampilkan.")
        return
    display = df.copy()
    for col in ["Nominal Invoice", "Nominal T-Summary", "Selisih"]:
        if col in display.columns:
            display[col] = display[col].apply(fmt_currency)
    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )


# ---------- UI ----------
st.set_page_config(page_title="Rekonsiliasi Naik/Turun Golongan", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan")

with st.sidebar:
    st.header("1) Upload File")
    f_inv = st.file_uploader("📄 File Invoice (CSV/XLSX)", type=["csv", "xlsx", "xls"])
    f_tik = st.file_uploader("🎫 File Tiket Summary (CSV/XLSX)", type=["csv", "xlsx", "xls"])
    st.divider()
    st.caption("Catatan: Jika ada beberapa Nomor Invoice di file, nilai akan di-*sum* per Nomor Invoice.")

df_inv = load_dataframe(f_inv)
df_tik = load_dataframe(f_tik)

if not df_inv.empty:
    st.subheader("Preview: Invoice")
    st.dataframe(df_inv.head(10), use_container_width=True, hide_index=True)
if not df_tik.empty:
    st.subheader("Preview: Tiket Summary")
    st.dataframe(df_tik.head(10), use_container_width=True, hide_index=True)

if df_inv.empty or df_tik.empty:
    st.info("Unggah kedua file untuk melanjutkan.")
    st.stop()

st.divider()
st.subheader("2) Pemetaan Kolom")

invoice_key_guess = guess_column(
    df_inv.columns,
    ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"],
)
invoice_amt_guess = guess_column(
    df_inv.columns,
    ["harga", "nilai", "amount", "nominal", "total", "grand total"],
)
tiket_key_guess = guess_column(
    df_tik.columns,
    ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"],
)
tiket_amt_guess = guess_column(
    df_tik.columns,
    ["tarif", "harga", "nilai", "amount", "nominal", "total", "grand total"],
)

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Invoice**")
    inv_key = st.selectbox("Kolom Nomor Invoice (Invoice)", options=list(df_inv.columns), index=(
        list(df_inv.columns).index(invoice_key_guess) if invoice_key_guess in df_inv.columns else 0
    ))
    inv_amt = st.selectbox("Kolom Nominal/Harga (Invoice)", options=list(df_inv.columns), index=(
        list(df_inv.columns).index(invoice_amt_guess) if invoice_amt_guess in df_inv.columns else 0
    ))
with col2:
    st.markdown("**Tiket Summary**")
    tik_key = st.selectbox("Kolom Nomor Invoice (Tiket Summary)", options=list(df_tik.columns), index=(
        list(df_tik.columns).index(tiket_key_guess) if tiket_key_guess in df_tik.columns else 0
    ))
    tik_amt = st.selectbox("Kolom Nominal/Tarif (Tiket Summary)", options=list(df_tik.columns), index=(
        list(df_tik.columns).index(tiket_amt_guess) if tiket_amt_guess in df_tik.columns else 0
    ))

st.divider()
st.subheader("3) Proses Rekonsiliasi")
only_diff = st.checkbox("Hanya tampilkan yang berbeda (Selisih ≠ 0)", value=False)
go = st.button("🚀 Proses")

if go:
    # Validasi kolom
    missing_msgs = []
    for df, need_cols, src in [
        (df_inv, [inv_key, inv_amt], "Invoice"),
        (df_tik, [tik_key, tik_amt], "Tiket Summary"),
    ]:
        for c in need_cols:
            if c not in df.columns:
                missing_msgs.append(f"- `{c}` tidak ditemukan di {src}")
    if missing_msgs:
        st.error("Kolom tidak ditemukan:\n" + "\n".join(missing_msgs))
        st.stop()

    agg_inv, agg_tik, merged = reconcile(
        df_inv, inv_key, inv_amt, df_tik, tik_key, tik_amt, only_diff
    )

    # Ringkasan
    total_inv = float(agg_inv[agg_inv.columns[1]].sum()) if not agg_inv.empty else 0.0
    total_tik = float(agg_tik[agg_tik.columns[1]].sum()) if not agg_tik.empty else 0.0
    total_diff = float(merged["Selisih"].sum()) if not merged.empty else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nominal Invoice", fmt_currency(total_inv))
    m2.metric("Total Nominal T-Summary", fmt_currency(total_tik))
    m3.metric("Total Selisih (Invoice − T-Summary)", fmt_currency(total_diff))
    if not merged.empty:
        naik = int((merged["Kategori"] == "Naik").sum())
        turun = int((merged["Kategori"] == "Turun").sum())
        sama = int((merged["Kategori"] == "Sama").sum())
    else:
        naik = turun = sama = 0
    m4.metric("Naik / Turun / Sama", f"{naik} / {turun} / {sama}")

    st.subheader("Hasil Rekonsiliasi")
    display_table(merged)

    # Unduhan
    st.markdown("**Unduh Hasil**")
    csv_bytes = merged.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download CSV",
        data=csv_bytes,
        file_name="rekonsiliasi.csv",
        mime="text/csv",
    )
    xlsx_bytes = df_to_excel_bytes(merged)
    if xlsx_bytes:
        st.download_button(
            "⬇️ Download Excel (XLSX)",
            data=xlsx_bytes,
            file_name="rekonsiliasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.caption("Excel engine tidak tersedia—gunakan CSV atau tambahkan paket `xlsxwriter`/`openpyxl`.")

    st.success("Selesai.")
