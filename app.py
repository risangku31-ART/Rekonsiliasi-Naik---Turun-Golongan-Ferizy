# app.py
# Rekonsiliasi Naik/Turun Golongan — Mode SUMIFS (Basis: Invoice)
# Keberangkatan ← T-Summary: "Asal", Tujuan ← Invoice: "Tujuan"

import csv
import io
import re
from typing import Dict, Iterable, List, Optional, Tuple
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import streamlit as st


# ----------------------------- Capability -----------------------------
def has_module(name: str) -> bool:
    try:
        __import__(name)
        return True
    except Exception:
        return False


def available_reader_engines() -> List[str]:
    engines = ["pure-xlsx"]  # fallback .xlsx tanpa dependency
    if has_module("pandas") and has_module("openpyxl"):
        engines.append("openpyxl")
    if has_module("pandas") and has_module("pandas_calamine"):
        engines.append("calamine")
    if has_module("pandas") and has_module("xlrd"):
        engines.append("xlrd")
    return engines


# ----------------------------- CSV / PASTE Parsers -----------------------------
def guess_delimiter(sample: str) -> str:
    if "\t" in sample:
        return "\t"
    if sample.count(";") >= sample.count(",") and ";" in sample:
        return ";"
    if "," in sample:
        return ","
    return "|"


def read_csv_file(file) -> List[Dict[str, str]]:
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
        delim = csv.Sniffer().sniff(text[:2048]).delimiter
    except Exception:
        delim = guess_delimiter(text)
    return [dict(r) for r in csv.DictReader(io.StringIO(text), delimiter=delim)]


def read_paste(text: str) -> List[Dict[str, str]]:
    text = (text or "").strip()
    if not text:
        return []
    delim = guess_delimiter(text)
    try:
        return [dict(r) for r in csv.DictReader(io.StringIO(text), delimiter=delim)]
    except Exception:
        return []


# ----------------------------- Minimal XLSX Reader (pure-Python) -----------------------------
NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

def _xlsx_col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - 64)
    return n - 1

def _xlsx_ref_to_rc(ref: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        return 0, 0
    col_letters, row_str = m.group(1), m.group(2)
    return int(row_str) - 1, _xlsx_col_to_idx(col_letters)

def _xlsx_read_shared_strings(z: ZipFile) -> List[str]:
    sst = []
    try:
        with z.open("xl/sharedStrings.xml") as f:
            tree = ET.parse(f)
        for si in tree.getroot().iterfind(".//main:si", NS):
            texts = [t.text or "" for t in si.findall(".//main:t", NS)]
            sst.append("".join(texts))
    except KeyError:
        pass
    return sst

def _xlsx_find_first_sheet_path(z: ZipFile) -> Optional[str]:
    if "xl/worksheets/sheet1.xml" in z.namelist():
        return "xl/worksheets/sheet1.xml"
    try:
        with z.open("xl/workbook.xml") as f:
            wb = ET.parse(f).getroot()
        first_sheet = wb.find(".//main:sheets/main:sheet", NS)
        if first_sheet is None:
            return None
        rid = first_sheet.attrib.get(f"{{{NS['r']}}}id")
        with z.open("xl/_rels/workbook.xml.rels") as f:
            rels = ET.parse(f).getroot()
        for rel in rels.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            if rel.attrib.get("Id") == rid:
                target = rel.attrib.get("Target")
                return "xl/" + target if not target.startswith("xl/") else target
    except KeyError:
        return None
    return None

def read_xlsx_pure(bytes_data: bytes) -> List[Dict[str, str]]:
    z = ZipFile(io.BytesIO(bytes_data))
    sst = _xlsx_read_shared_strings(z)
    sheet_path = _xlsx_find_first_sheet_path(z)
    if not sheet_path:
        return []
    with z.open(sheet_path) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    rows_dict: Dict[int, Dict[int, str]] = {}
    max_col = -1
    for c in root.findall(".//main:c", NS):
        ref = c.attrib.get("r", "A1")
        t = c.attrib.get("t")
        v = c.find("main:v", NS)
        is_node = c.find("main:is", NS)
        text = ""
        if t == "s":
            idx = int(v.text) if v is not None and v.text else -1
            text = sst[idx] if 0 <= idx < len(sst) else ""
        elif t == "inlineStr" and is_node is not None:
            ts = [t.text or "" for t in is_node.findall(".//main:t", NS)]
            text = "".join(ts)
        elif t == "b":
            text = "TRUE" if (v is not None and v.text == "1") else "FALSE"
        else:
            text = (v.text or "") if v is not None else ""
        r, cidx = _xlsx_ref_to_rc(ref)
        rows_dict.setdefault(r, {})[cidx] = text
        max_col = max(max_col, cidx)

    if not rows_dict:
        return []

    matrix: List[List[str]] = []
    for r in sorted(rows_dict.keys()):
        row = ["" for _ in range(max_col + 1)]
        for cidx, val in rows_dict[r].items():
            if 0 <= cidx <= max_col:
                row[cidx] = val
        matrix.append(row)

    header: List[str] = []
    data_start = 0
    for i, row in enumerate(matrix):
        if any(cell.strip() for cell in row):
            header = row
            data_start = i + 1
            break
    if not header:
        return []

    norm = {}
    final_header = []
    for h in header:
        base = (h or "").strip() or "COL"
        name = base
        k = 2
        while name.lower() in norm:
            name = f"{base}_{k}"
            k += 1
        norm[name.lower()] = True
        final_header.append(name)

    out: List[Dict[str, str]] = []
    for row in matrix[data_start:]:
        if not any(cell.strip() for cell in row):
            continue
        rec = {}
        for j, name in enumerate(final_header):
            rec[name] = row[j] if j < len(row) else ""
        out.append(rec)
    return out


# ----------------------------- Load Many -----------------------------
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
            elif low.endswith((".xlsx", ".xlsm")):
                if safe_mode:
                    st.warning(f"Lewati `{f.name}` (Excel) karena Safe Mode aktif. Unggah CSV atau matikan Safe Mode.")
                else:
                    if forced_engine == "pure-xlsx" or (forced_engine == "Auto" and "pure-xlsx" in available_reader_engines()):
                        f.seek(0); data = f.read()
                        rows = read_xlsx_pure(data)
                    else:
                        if not has_module("pandas"):
                            st.warning(f"Lewati `{f.name}`: pandas tidak tersedia. Pilih `pure-xlsx`.")
                        else:
                            import pandas as pd
                            eng = None
                            if forced_engine == "openpyxl" and has_module("openpyxl"):
                                eng = "openpyxl"
                            elif forced_engine == "calamine" and has_module("pandas_calamine"):
                                eng = "calamine"
                            elif forced_engine == "Auto":
                                if has_module("openpyxl"):
                                    eng = "openpyxl"
                                elif has_module("pandas_calamine"):
                                    eng = "calamine"
                            if eng:
                                f.seek(0)
                                df = pd.read_excel(f, dtype=str, engine=eng)
                                rows = df.fillna("").astype(str).to_dict(orient="records")
                            else:
                                st.warning(f"Lewati `{f.name}`: Tidak ada engine openpyxl/calamine. Pilih `pure-xlsx`.")
            elif low.endswith(".xls"):
                st.warning(f"Lewati `{f.name}` (.xls lama). Konversi ke CSV atau .xlsx.")
            else:
                rows = read_csv_file(f)
        except Exception as e:
            st.warning(f"Lewati `{f.name}`: {e}")
            rows = []
        for r in rows:
            r["Sumber File"] = f.name
        out.extend(rows)
    return out


# ----------------------------- Business logic (SUMIFS) -----------------------------
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


def ordered_keys(rows: List[Dict[str, str]], key_col: str) -> List[str]:
    seen = set()
    order: List[str] = []
    for r in rows:
        k = coerce_invoice_key(r.get(key_col, ""))
        if k and k not in seen:
            seen.add(k)
            order.append(k)
    return order


def collect_first_map(rows: List[Dict[str, str]], key_col: str, val_col: Optional[str]) -> Dict[str, str]:
    if not val_col:
        return {}
    out: Dict[str, str] = {}
    for r in rows:
        k = coerce_invoice_key(r.get(key_col, ""))
        v = (r.get(val_col, "") or "").strip()
        if k and v and k not in out:
            out[k] = v
    return out


def collect_unique_join_map(rows: List[Dict[str, str]], key_col: str, val_col: Optional[str], sep: str = ", ") -> Dict[str, str]:
    if not val_col:
        return {}
    out: Dict[str, str] = {}
    seen_per_key: Dict[str, set] = {}
    for r in rows:
        k = coerce_invoice_key(r.get(key_col, ""))
        v = (r.get(val_col, "") or "").strip()
        if not k or not v:
            continue
        s = seen_per_key.setdefault(k, set())
        if v not in s:
            s.add(v)
            if k in out:
                out[k] = f"{out[k]}{sep}{v}"
            else:
                out[k] = v
    return out


# ----------------------------- UI -----------------------------
st.set_page_config(page_title="Rekonsiliasi Naik/Turun Golongan (SUMIFS)", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan — Mode SUMIFS (Basis: Invoice)")

with st.expander("ℹ️ Mode & Engine", expanded=True):
    safe_mode = st.toggle(
        "Safe Mode (CSV & PASTE only)",
        value=False,
        help="ON: lewati Excel. OFF: .xlsx diproses via engine atau pure-xlsx.",
    )
    avail = available_reader_engines()
    forced_engine = st.selectbox(
        "Paksa engine Excel",
        options=(["Auto"] + avail) if not safe_mode else ["Auto"],
        index=0,
        help="Auto: openpyxl→calamine→pure-xlsx (.xlsx). .xls disarankan konversi ke CSV/.xlsx.",
    )
    st.write(f"**Status:** {'🟢 Safe Mode ON' if safe_mode else '🔵 Safe Mode OFF'} — Engine: {', '.join(avail)}")

with st.sidebar:
    st.header("1) Upload File (Multiple)")
    inv_files = st.file_uploader(
        "📄 Invoice — CSV/XLSX/XLSM/XLS",
        type=["csv", "xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
    )
    tik_files = st.file_uploader(
        "🎫 Tiket Summary — CSV/XLSX/XLSM/XLS",
        type=["csv", "xlsx", "xlsm", "xls"],
        accept_multiple_files=True,
    )
    st.caption("Patokan = Nomor Invoice dari uploader Invoice (SUMIFS/LEFT JOIN).")

st.subheader("Opsional: Tempel Data dari Excel")
c1, c2 = st.columns(2)
with c1:
    paste_inv = st.text_area("PASTE — Invoice (TSV/CSV)", height=160, placeholder="Tempel data Invoice di sini…")
with c2:
    paste_tik = st.text_area("PASTE — Tiket Summary (TSV/CSV)", height=160, placeholder="Tempel data Tiket Summary di sini…")

# Compose data
rows_inv: List[Dict[str, str]] = []
rows_inv.extend(load_many(inv_files, safe_mode, forced_engine))
for r in read_paste(paste_inv):
    r["Sumber File"] = "PASTE:Invoice"; rows_inv.append(r)

rows_tik: List[Dict[str, str]] = []
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
    st.info("Unggah minimal satu file/PASTE untuk **Invoice** dan **Tiket Summary**.")
    st.stop()

# ----------------------------- Mapping -----------------------------
st.divider()
st.subheader("2) Pemetaan Kolom")

inv_cols = union_columns(rows_inv); tik_cols = union_columns(rows_tik)

# key & nominal
inv_key_guess = guess_column(inv_cols, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"])
inv_amt_guess = guess_column(inv_cols, ["harga", "nilai", "amount", "nominal", "total", "grand total"])
tik_key_guess = guess_column(tik_cols, ["nomor invoice", "no invoice", "invoice", "invoice number", "no faktur", "nomor faktur"])
tik_amt_guess = guess_column(tik_cols, ["tarif", "harga", "nilai", "amount", "nominal", "total", "grand total"])

col_a, col_b = st.columns(2)
with col_a:
    st.markdown("**Invoice — kolom utama**")
    inv_key = st.selectbox("Kolom Nomor Invoice (Invoice)", inv_cols, index=inv_cols.index(inv_key_guess) if inv_key_guess in inv_cols else 0)
    inv_amt = st.selectbox("Kolom Nominal/Harga (Invoice)", inv_cols, index=inv_cols.index(inv_amt_guess) if inv_amt_guess in inv_cols else 0)
with col_b:
    st.markdown("**Tiket Summary — kolom utama**")
    tik_key = st.selectbox("Kolom Nomor Invoice (T-Summary)", tik_cols, index=tik_cols.index(tik_key_guess) if tik_key_guess in tik_cols else 0)
    tik_amt = st.selectbox("Kolom Nominal/Tarif (T-Summary)", tik_cols, index=tik_cols.index(tik_amt_guess) if tik_amt_guess in tik_cols else 0)

st.markdown("**Pemetaan kolom tambahan (opsional, untuk ditampilkan di hasil):**")

def select_with_empty(label: str, options: List[str], guess: Optional[str] = None):
    opts = ["— Kosong —"] + options
    idx = 0
    if guess in options:
        idx = options.index(guess) + 1
    return st.selectbox(label, opts, index=idx)

# Guesses (disesuaikan permintaan)
inv_tgl_inv_guess   = guess_column(inv_cols, ["tanggal invoice", "tgl invoice", "tanggal", "tgl"])
inv_pay_inv_guess   = guess_column(inv_cols, ["tanggal invoice", "tgl invoice", "pembayaran", "tgl pembayaran"])
inv_tujuan_guess    = guess_column(inv_cols, ["tujuan", "destination", "destinasi"])   # ← Tujuan dari Invoice
inv_channel_guess   = guess_column(inv_cols, ["channel"])
inv_merchant_guess  = guess_column(inv_cols, ["merchant", "mid", "acquirer"])

ts_kode_booking_guess = guess_column(tik_cols, ["kode booking", "kode boking", "booking code"])
ts_nomor_tiket_guess  = guess_column(tik_cols, ["nomor tiket", "no tiket", "ticket"])
ts_pay_ts_guess       = guess_column(tik_cols, ["pembayaran", "tanggal pembayaran", "tgl bayar", "tgl pembayaran"])
ts_golongan_guess     = guess_column(tik_cols, ["golongan", "kelas", "class"])
ts_asal_guess         = guess_column(tik_cols, ["asal", "keberangkatan", "origin"])    # ← Keberangkatan dari T-Summary (Asal)
ts_cetak_bp_guess     = guess_column(tik_cols, ["cetak boarding pass", "tgl cetak", "boarding pass"])

c1, c2 = st.columns(2)
with c1:
    st.markdown("**Tambahan dari Invoice**")
    sel_inv_tgl_inv   = select_with_empty("Tanggal Invoice (Invoice)", inv_cols, inv_tgl_inv_guess)
    sel_inv_pay_inv   = select_with_empty("Tanggal Pembayaran Invoice (kolom di Invoice)", inv_cols, inv_pay_inv_guess)
    sel_inv_tujuan    = select_with_empty("Tujuan (Invoice)", inv_cols, inv_tujuan_guess)  # swapped here
    sel_inv_channel   = select_with_empty("Channel (Invoice)", inv_cols, inv_channel_guess)
    sel_inv_merchant  = select_with_empty("Merchant (Invoice)", inv_cols, inv_merchant_guess)

with c2:
    st.markdown("**Tambahan dari T-Summary**")
    sel_ts_kode_bk    = select_with_empty("Kode Booking (T-Summary)", tik_cols, ts_kode_booking_guess)
    sel_ts_no_tiket   = select_with_empty("Nomor Tiket (T-Summary)", tik_cols, ts_nomor_tiket_guess)
    sel_ts_pay_ts     = select_with_empty("Tanggal Pembayaran T-Summary (kolom 'Pembayaran')", tik_cols, ts_pay_ts_guess)
    sel_ts_golongan   = select_with_empty("Golongan (T-Summary)", tik_cols, ts_golongan_guess)
    sel_ts_asal       = select_with_empty("Keberangkatan / Asal (T-Summary)", tik_cols, ts_asal_guess)  # swapped here
    sel_ts_cetak_bp   = select_with_empty("Tgl Cetak Boarding Pass (T-Summary)", tik_cols, ts_cetak_bp_guess)

def none_if_empty(x: str) -> Optional[str]:
    return None if x == "— Kosong —" else x

inv_tgl_inv_col   = none_if_empty(sel_inv_tgl_inv)
inv_pay_inv_col   = none_if_empty(sel_inv_pay_inv)
inv_tujuan_col    = none_if_empty(sel_inv_tujuan)
inv_channel_col   = none_if_empty(sel_inv_channel)
inv_merchant_col  = none_if_empty(sel_inv_merchant)

ts_kode_bk_col    = none_if_empty(sel_ts_kode_bk)
ts_no_tiket_col   = none_if_empty(sel_ts_no_tiket)
ts_pay_ts_col     = none_if_empty(sel_ts_pay_ts)
ts_gol_col        = none_if_empty(sel_ts_golongan)
ts_asal_col       = none_if_empty(sel_ts_asal)
ts_cetak_bp_col   = none_if_empty(sel_ts_cetak_bp)

# ----------------------------- SUMIFS (basis Invoice) -----------------------------
st.divider()
st.subheader("3) Proses Rekonsiliasi — SUMIFS (Basis Invoice)")
only_diff = st.checkbox("Hanya tampilkan yang berbeda (Selisih ≠ 0)", value=False)
go = st.button("🚀 Proses")

if go:
    agg_inv = aggregate_sum(rows_inv, inv_key, inv_amt)
    agg_tik = aggregate_sum(rows_tik, tik_key, tik_amt)
    keys_ordered = ordered_keys(rows_inv, inv_key)

    # Maps tambahan
    inv_tgl_inv_map   = collect_first_map(rows_inv, inv_key, inv_tgl_inv_col)
    inv_pay_inv_map   = collect_first_map(rows_inv, inv_key, inv_pay_inv_col)
    inv_tujuan_map    = collect_first_map(rows_inv, inv_key, inv_tujuan_col)      # ← from Invoice
    inv_channel_map   = collect_first_map(rows_inv, inv_key, inv_channel_col)
    inv_merchant_map  = collect_first_map(rows_inv, inv_key, inv_merchant_col)

    ts_kode_bk_map    = collect_unique_join_map(rows_tik, tik_key, ts_kode_bk_col)
    ts_no_tiket_map   = collect_unique_join_map(rows_tik, tik_key, ts_no_tiket_col)
    ts_pay_ts_map     = collect_first_map(rows_tik, tik_key, ts_pay_ts_col)
    ts_gol_map        = collect_unique_join_map(rows_tik, tik_key, ts_gol_col)
    ts_asal_map       = collect_unique_join_map(rows_tik, tik_key, ts_asal_col)   # ← from T-Summary
    ts_cetak_bp_map   = collect_unique_join_map(rows_tik, tik_key, ts_cetak_bp_col)

    out_rows: List[Dict[str, str]] = []
    total_inv = total_tik = total_diff = 0.0
    naik = turun = sama = 0

    for k in keys_ordered:
        v_inv = float(agg_inv.get(k, 0.0))
        v_tik = float(agg_tik.get(k, 0.0))
        diff = v_inv - v_tik
        cat = "Naik" if diff > 0 else ("Turun" if diff < 0 else "Sama")

        row = {
            "Tanggal Invoice":              inv_tgl_inv_map.get(k, ""),
            "Kode Booking":                 ts_kode_bk_map.get(k, ""),
            "Nomor Tiket":                  ts_no_tiket_map.get(k, ""),
            "Tanggal Pembayaran Invoice":   inv_pay_inv_map.get(k, ""),
            "Tanggal Pembayaran T-Summary": ts_pay_ts_map.get(k, ""),
            "Golongan":                     ts_gol_map.get(k, ""),
            "Keberangkatan":                ts_asal_map.get(k, ""),       # swapped source
            "Tujuan":                       inv_tujuan_map.get(k, ""),     # swapped source
            "Tgl Cetak Boarding Pass":      ts_cetak_bp_map.get(k, ""),
            "Channel":                      inv_channel_map.get(k, ""),
            "Merchant":                     inv_merchant_map.get(k, ""),
            "Nomor Invoice":                k,
            "Nominal Invoice (SUMIFS)":     format_idr(v_inv),
            "Nominal T-Summary (SUMIFS)":   format_idr(v_tik),
            "Selisih":                      format_idr(diff),
            "Kategori":                     cat,
        }

        if (not only_diff) or (diff != 0):
            out_rows.append(row)

        total_inv += v_inv
        total_tik += v_tik
        total_diff += diff
        if cat == "Naik": naik += 1
        elif cat == "Turun": turun += 1
        else: sama += 1

    extra_tik_only = len(set(agg_tik.keys()) - set(keys_ordered))
    if extra_tik_only:
        st.caption(f"ℹ️ {extra_tik_only} Nomor Invoice hanya ada di T-Summary (diabaikan — basis Invoice).")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Invoice (SUMIFS)", format_idr(total_inv))
    m2.metric("Total T-Summary (SUMIFS)", format_idr(total_tik))
    m3.metric("Total Selisih (Inv − T)", format_idr(total_diff))
    m4.metric("Naik / Turun / Sama", f"{naik} / {turun} / {sama}")

    st.subheader("Hasil Rekonsiliasi (SUMIFS, Basis Invoice)")
    st.data_editor(out_rows, use_container_width=True, disabled=True, key="result")

    csv_cols = [
        "Tanggal Invoice",
        "Kode Booking",
        "Nomor Tiket",
        "Tanggal Pembayaran Invoice",
        "Tanggal Pembayaran T-Summary",
        "Golongan",
        "Keberangkatan",
        "Tujuan",
        "Tgl Cetak Boarding Pass",
        "Channel",
        "Merchant",
        "Nomor Invoice",
        "Nominal Invoice (SUMIFS)",
        "Nominal T-Summary (SUMIFS)",
        "Selisih",
        "Kategori",
    ]
    si = io.StringIO()
    w = csv.writer(si)
    w.writerow(csv_cols)
    for r in out_rows:
        w.writerow([r.get(c, "") for c in csv_cols])
    st.download_button("⬇️ Download CSV", data=si.getvalue().encode("utf-8"), file_name="rekonsiliasi_sumifs_basis_invoice.csv", mime="text/csv")
