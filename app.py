# app.py
# Rekonsiliasi Naik/Turun Golongan — CSV & PASTE, + fallback XLSX murni-Python (pure-xlsx)

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
    # 'pure-xlsx' selalu tersedia (untuk .xlsx). Engine lain butuh pandas.
    engines = ["pure-xlsx"]
    if has_module("pandas") and has_module("openpyxl"):
        engines.append("openpyxl")
    if has_module("pandas") and (has_module("pandas_calamine")):
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
    return n - 1  # zero-based

def _xlsx_ref_to_rc(ref: str) -> Tuple[int, int]:
    # e.g. C5 -> (row=4, col=2) zero-based
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
            # collect all <t> inside si (handle rich text)
            texts = []
            for t in si.findall(".//main:t", NS):
                texts.append(t.text or "")
            sst.append("".join(texts))
    except KeyError:
        pass  # no sharedStrings
    return sst

def _xlsx_find_first_sheet_path(z: ZipFile) -> Optional[str]:
    # Try common path first
    if "xl/worksheets/sheet1.xml" in z.namelist():
        return "xl/worksheets/sheet1.xml"
    # Robust: read workbook + rels
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
                target = rel.attrib.get("Target")  # e.g. worksheets/sheet1.xml
                if not target:
                    return None
                path = "xl/" + target if not target.startswith("xl/") else target
                return path
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

    # Collect cells into a row-wise matrix
    rows_dict: Dict[int, Dict[int, str]] = {}
    max_col = -1
    for c in root.findall(".//main:c", NS):
        ref = c.attrib.get("r", "A1")
        t = c.attrib.get("t")  # 's', 'inlineStr', 'b', 'str', etc.
        v = c.find("main:v", NS)
        is_node = c.find("main:is", NS)
        text = ""
        if t == "s":  # shared string
            idx = int(v.text) if v is not None and v.text else -1
            text = sst[idx] if 0 <= idx < len(sst) else ""
        elif t == "inlineStr" and is_node is not None:
            ts = [t.text or "" for t in is_node.findall(".//main:t", NS)]
            text = "".join(ts)
        elif t == "b":
            text = "TRUE" if (v is not None and v.text == "1") else "FALSE"
        else:
            text = (v.text or "") if v is not None else ""  # number or plain

        r, cidx = _xlsx_ref_to_rc(ref)
        row_map = rows_dict.setdefault(r, {})
        row_map[cidx] = text
        max_col = max(max_col, cidx)

    if not rows_dict:
        return []

    # Build 2D rows ordered by row index
    matrix: List[List[str]] = []
    for r in sorted(rows_dict.keys()):
        row = ["" for _ in range(max_col + 1)]
        for cidx, val in rows_dict[r].items():
            if 0 <= cidx <= max_col:
                row[cidx] = val
        matrix.append(row)

    # First non-empty row as header
    header: List[str] = []
    data_start = 0
    for i, row in enumerate(matrix):
        if any(cell.strip() for cell in row):
            header = row
            data_start = i + 1
            break
    if not header:
        return []

    # Deduplicate empty/duplicate headers
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
                    # engine selection
                    if forced_engine == "pure-xlsx" or (forced_engine == "Auto" and "pure-xlsx" in available_reader_engines()):
                        f.seek(0)
                        data = f.read()
                        rows = read_xlsx_pure(data)
                    else:
                        # coba via pandas jika tersedia
                        if not has_module("pandas"):
                            st.warning(f"Lewati `{f.name}`: pandas tidak tersedia. Gunakan engine `pure-xlsx` atau unggah CSV.")
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
        value=False,
        help="Jika ON, Excel di-skip. Matikan untuk memproses .xlsx (pure-xlsx) atau engine lain.",
    )
    avail = available_reader_engines()
    forced_engine = st.selectbox(
        "Paksa engine Excel",
        options=(["Auto"] + avail) if not safe_mode else ["Auto"],
        index=0,
        help="Auto: openpyxl→calamine→pure-xlsx untuk .xlsx; .xls disarankan konversi ke CSV/.xlsx.",
    )
    st.write(f"**Status:** {'🟢 Safe Mode ON' if safe_mode else '🔵 Safe Mode OFF'} — Engine tersedia: {', '.join(avail)}")

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
    st.caption("Jika tidak ada engine, pilih `pure-xlsx` untuk .xlsx; .xls lama akan di-skip.")

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
