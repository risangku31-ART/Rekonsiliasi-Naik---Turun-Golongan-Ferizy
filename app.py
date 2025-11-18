# app.py
# Rekonsiliasi Naik/Turun Golongan — CSV & PASTE first, Excel optional bila engine tersedia

import csv
import io
import re
from typing import Dict, Iterable, List, Optional, Tuple

import streamlit as st


# ----------------------------- Capability -----------------------------
def has_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False


def available_reader_engines() -> List[str]:
    # Hanya tampilkan engine jika pandas ada, karena read_excel butuh pandas.
    if not has_module("pandas"):
        return []
    engines = []
    if has_module("openpyxl"):
        engines.append("openpyxl")
    if has_module("pandas_calamine"):
        engines.append("calamine")
    if has_module("xlrd"):
        engines.append("xlrd")
    return engines


# ----------------------------- Parsers -----------------------------
def guess_delimiter(sample: str) -> str:
    if "\t" in sample:
        return "\t"
    if sample.count(";") >= sample.count(",") and ";" in sample:
        return ";"
    if "," in sample:
        return ","
    return "|"


def read_csv_file(file) -> List[Dict[str, str]]:
    """CSV robust (utf-8→cp1252 fallback)."""
    file.seek(0)
    data = file.read()
    if isinstance(data, bytes):
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("cp1252", errors="ignore")
    else:
        text = data
    try:
        dialect = csv.Sniffer().sniff(text[:2048])
        delim = dialect.delimiter
    except Exception:
        delim = guess_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    return [dict(r) for r in reader]


def read_paste(text: str) -> List[Dict[str, str]]:
    text = (text or "").strip()
    if not text:
        return []
    delim = guess_delimiter(text)
    try:
        reader = csv.DictReader(io.StringIO(text), delimiter=delim)
        return [dict(r) for r in reader]
    except Exception:
        try:
            # Fallback engine=python autodetect
            return [dict(r) for r in csv.DictReader(io.StringIO(text))]
        except Exception:
            st.error("Gagal mengurai data PASTE. Pastikan kolom dipisah TAB/;/,/|.")
            return []


def read_excel_rows(file, ext: str, forced_engine: str) -> List[Dict[str, str]]:
    """
    Baca Excel -> list of dicts. Butuh pandas + engine.
    Kenapa: hindari crash jika dependency tidak ada; tampilkan warning & skip.
    """
    name = getattr(file, "name", "uploaded")
    try:
        import pandas as pd  # import lokal agar app tetap jalan tanpa pandas
    except Exception:
        st.warning(f"Lewati `{name}`: pandas tidak tersedia. Unggah CSV atau gunakan PASTE.")
        return []

    def engine_ok(e: str) -> bool:
        try:
            __import__(e if e != "calamine" else "pandas_calamine")
            return True
        except Exception:
            return False

    try:
        if forced_engine == "Auto":
            if ext == ".xls":
                if not engine_ok("xlrd"):
                    raise RuntimeError("Butuh engine xlrd untuk .xls")
                df = pd.read_excel(file, dtype=str, engine="xlrd")
            else:
                if engine_ok("openpyxl"):
                    df = pd.read_excel(file, dtype=str, engine="openpyxl")
                elif engine_ok("calamine"):
                    df = pd.read_excel(file, dtype=str, engine="calamine")
                else:
                    raise RuntimeError("Tidak ada engine openpyxl/calamine")
        else:
            if ext == ".xls" and forced_engine != "xlrd":
                raise RuntimeError(f".xls hanya didukung xlrd, bukan {forced_engine}")
            if ext in (".xlsx", ".xlsm") and forced_engine == "xlrd":
                raise RuntimeError(".xlsx/.xlsm tidak didukung xlrd")
            if forced_engine == "openpyxl" and not engine_ok("openpyxl"):
                raise RuntimeError("engine openpyxl tidak tersedia")
            if forced_engine == "calamine" and not engine_ok("calamine"):
                raise RuntimeError("engine calamine tidak tersedia")
            df = pd.read_excel(file, dtype=str, engine=("xlrd" if forced_engine == "xlrd" else forced_engine))

        rows = df.fillna("").astype(str).to_dict(orient="records")
        return rows
    except Exception as e:
        st.warning(f"Lewati `{name}`: {e}")
        return []


def load_many(files, safe_mode: bool, forced_engine: str) -> List[Dict[str, str]]:
    if not files:
        return []
    out: List[Dict[str, str]] = []
    for f in files:
        low = (f.name or "").lower()
        rows: List[Dict[str, str]] = []
        try:
            if low.endswith(".csv"):
                rows = read_csv_file(f)
            elif low.endswith((".xls", ".xlsx", ".xlsm")):
                if safe_mode:
                    st.warning(f"Lewati `{f.name}` (Excel) karena Safe Mode aktif. Unggah CSV atau matikan Safe Mode.")
                    rows = []
                else:
                    ext = ".xls" if low.endswith(".xls") else (".xlsm" if low.endswith(".xlsm") else ".xlsx")
                    rows = read_excel_rows(f, ext, forced_engine)
            else:
                rows = read_csv_file(f)  # fallback
        except Exception as e:
            st.warning(f"Lewati `{f.name}`: {e}")
            rows = []
        for r in rows:
            r["Sumber File"] = f.name
        out.extend(rows)
    return out


# ----------------------------- Business logic -----------------------------
def normalize_colname(s: str) -> str:
    s = (s or "").lower().strip()
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


def guess_column(columns: Iterable[str], candidates: Iterable[str]) -> Optional[str]:
    cols = [c for c in columns if c is not None]
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
    return cols[0] if cols else None


def coerce_invoice_key(x: str) -> str:
    s = (x or "").strip()
    s = re.sub(r"\s+", "", s)
    return s.upper()


def parse_money(x: str) -> float:
    if x is None:
        return 0.0
    s = str(x).strip()
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
            if not re.search(r"\.[0-9]{1,2}$", s):
                s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        s2 = re.sub(r"[^\d\-]", "", s)
        try:
            return float(s2)
        except Exception:
            return 0.0


def format_idr(n: float) -> str:
    s = f"{float(n):,.2f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def union_columns(rows: List[Dict[str, str]]) -> List[str]:
    cols, seen = [], set()
    for r in rows:
        for k in r.keys():
            if k not in seen:
                cols.append(k); seen.add(k)
    return cols


def aggregate_sum(rows: List[Dict[str, str]], key_col: str, amt_col: str) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in rows:
        key = coerce_invoice_key(r.get(key_col, ""))
        val = parse_money(r.get(amt_col, "0"))
        out[key] = out.get(key, 0.0) + val
    return out


# ----------------------------- UI -----------------------------
st.set_page_config(page_title="Rekonsiliasi Naik/Turun Golongan", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan")

with st.expander("ℹ️ Mode & Engine", expanded=True):
    safe_mode = st.toggle(
        "Safe Mode (CSV & PASTE only) — aktifkan bila install dependencies gagal",
        value=True,
        help="Jika ON, file Excel di-skip. Matikan untuk coba proses Excel (butuh pandas + engine).",
    )
    avail = available_reader_engines()
    forced_engine = st.selectbox(
        "Paksa engine Excel",
        options=(["Auto"] + avail) if not safe_mode else ["Auto"],
        index=0,
        help="Auto: .xls→xlrd, .xlsx/.xlsm→openpyxl lalu calamine.",
    )
    st.write(f"**Status:** {'🟢 Safe Mode ON' if safe_mode else '🔵 Safe Mode OFF'}")

with st.sidebar:
    st.header("1) Upload File (Multiple)")
    inv_files = st.file_uploader(
        "📄 Invoice — CSV/XLS/XLSX/XLSM",
        type=["csv", "xls", "xlsx", "xlsm"],
        accept_multiple_files=True,
    )
    tik_files = st.file_uploader(
        "🎫 Tiket Summary — CSV/XLS/XLSX/XLSM",
        type=["csv", "xls", "xlsx", "xlsm"],
        accept_multiple_files=True,
    )
    st.caption("Excel akan di-skip jika Safe Mode ON atau engine tidak tersedia.")

st.subheader("Opsional: Tempel Data dari Excel")
c1, c2 = st.columns(2)
with c1:
    paste_inv = st.text_area("PASTE — Invoice (TSV/CSV)", height=160, placeholder="Tempel data Invoice di sini…")
with c2:
    paste_tik = st.text_area("PASTE — Tiket Summary (TSV/CSV)", height=160, placeholder="Tempel data Tiket Summary di sini…")

# Compose data
rows_inv = []
rows_inv.extend(load_many(inv_files, safe_mode, forced_engine))
for r in read_paste(paste_inv):
    r["Sumber File"] = "PASTE:Invoice"; rows_inv.append(r)

rows_tik = []
rows_tik.extend(load_many(tik_files, safe_mode, forced_engine))
for r in read_paste(paste_tik):
    r["Sumber File"] = "PASTE:TiketSummary"; rows_tik.append(r)

# Previews
if rows_inv:
    st.subheader(f"Preview: Invoice (gabungan {len(rows_inv)} baris)")
    st.data_editor(rows_inv[: min(10, len(rows_inv))], use_container_width=True, disabled=True, key="prev_inv")
if rows_tik:
    st.subheader(f"Preview: Tiket Summary (gabungan {len(rows_tik)} baris)")
    st.data_editor(rows_tik[: min(10, len(rows_tik))], use_container_width=True, disabled=True, key="prev_tik")

if not rows_inv or not rows_tik:
    st.info("Unggah minimal satu file atau tempel data untuk **Invoice** dan **Tiket Summary**.")
    st.stop()

# Mapping
st.divider()
st.subheader("2) Pemetaan Kolom")
inv_cols = union_columns(rows_inv); tik_cols = union_columns(rows_tik)

inv_key_guess = guess_column(inv_cols, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"])
inv_amt_guess = guess_column(inv_cols, ["harga", "nilai", "amount", "nominal", "total", "grand total"])
tik_key_guess = guess_column(tik_cols, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"])
tik_amt_guess = guess_column(tik_cols, ["tarif", "harga", "nilai", "amount", "nominal", "total", "grand total"])

c3, c4 = st.columns(2)
with c3:
    st.markdown("**Invoice**")
    inv_key = st.selectbox("Kolom Nomor Invoice (Invoice)", inv_cols, index=inv_cols.index(inv_key_guess) if inv_key_guess in inv_cols else 0)
    inv_amt = st.selectbox("Kolom Nominal/Harga (Invoice)", inv_cols, index=inv_cols.index(inv_amt_guess) if inv_amt_guess in inv_cols else 0)
with c4:
    st.markdown("**Tiket Summary**")
    tik_key = st.selectbox("Kolom Nomor Invoice (Tiket Summary)", tik_cols, index=tik_cols.index(tik_key_guess) if tik_key_guess in tik_cols else 0)
    tik_amt = st.selectbox("Kolom Nominal/Tarif (Tiket Summary)", tik_cols, index=tik_cols.index(tik_amt_guess) if tik_amt_guess in tik_cols else 0)

# Process
st.divider()
st.subheader("3) Proses Rekonsiliasi")
only_diff = st.checkbox("Hanya tampilkan yang berbeda (Selisih ≠ 0)", value=False)
go = st.button("🚀 Proses")

if go:
    agg_inv = aggregate_sum(rows_inv, inv_key, inv_amt)
    agg_tik = aggregate_sum(rows_tik, tik_key, tik_amt)

    all_keys = sorted(set(agg_inv.keys()) | set(agg_tik.keys()))
    out_rows: List[Dict[str, str]] = []
    total_inv = total_tik = total_diff = 0.0
    naik = turun = sama = 0

    for k in all_keys:
        v_inv = float(agg_inv.get(k, 0.0))
        v_tik = float(agg_tik.get(k, 0.0))
        diff = v_inv - v_tik
        cat = "Naik" if diff > 0 else ("Turun" if diff < 0 else "Sama")
        if (not only_diff) or (diff != 0):
            out_rows.append(
                {
                    "Nomor Invoice": k,
                    "Nominal Invoice": format_idr(v_inv),
                    "Nominal T-Summary": format_idr(v_tik),
                    "Selisih": format_idr(diff),
                    "Kategori": cat,
                }
            )
        total_inv += v_inv
        total_tik += v_tik
        total_diff += diff
        if cat == "Naik":
            naik += 1
        elif cat == "Turun":
            turun += 1
        else:
            sama += 1

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Nominal Invoice", format_idr(total_inv))
    m2.metric("Total Nominal T-Summary", format_idr(total_tik))
    m3.metric("Total Selisih (Invoice − T-Summary)", format_idr(total_diff))
    m4.metric("Naik / Turun / Sama", f"{naik} / {turun} / {sama}")

    st.subheader("Hasil Rekonsiliasi")
    st.data_editor(out_rows, use_container_width=True, disabled=True, key="result")

    # Download CSV
    si = io.StringIO()
    w = csv.writer(si)
    w.writerow(["Nomor Invoice", "Nominal Invoice", "Nominal T-Summary", "Selisih", "Kategori"])
    for r in out_rows:
        w.writerow([r["Nomor Invoice"], r["Nominal Invoice"], r["Nominal T-Summary"], r["Selisih"], r["Kategori"]])
    st.download_button("⬇️ Download CSV", data=si.getvalue().encode("utf-8"), file_name="rekonsiliasi.csv", mime="text/csv")
