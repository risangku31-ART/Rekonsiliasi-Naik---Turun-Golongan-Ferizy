# app_ultra_safe.py
# Rekonsiliasi Naik/Turun Golongan — ULTRA-SAFE (tanpa fitur Streamlit yang rawan & tanpa lib eksternal)

import csv
import io
import re
import traceback
from typing import Dict, List, Optional, Iterable, Generator, Tuple
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import streamlit as st

# ===== 1) PAGE CONFIG (WAJIB PALING AWAL) =====
st.set_page_config(page_title="Rekonsiliasi ULTRA-SAFE", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan — ULTRA-SAFE")

# Metric font kecil (why: angka panjang tidak terpotong)
st.markdown("""
<style>
div[data-testid="stMetricLabel"] { font-size: 11px !important; }
div[data-testid="stMetricValue"] { font-size: 17px !important; }
div[data-testid="stMetricValue"] > div { white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; }
</style>
""", unsafe_allow_html=True)

# ===== 2) HELPERS =====
def guess_delimiter(sample: str) -> str:
    if "\t" in sample: return "\t"
    if sample.count(";") >= sample.count(",") and ";" in sample: return ";"
    if "," in sample: return ","
    return "|"

def sniff_delimiter_from_bytes(b: bytes) -> str:
    head = b[:4096].decode("utf-8", errors="ignore")
    try:
        return csv.Sniffer().sniff(head).delimiter
    except Exception:
        return guess_delimiter(head)

def normalize(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = (s.replace("no ", "nomor ").replace("no", "nomor")
           .replace("inv", "invoice").replace("harga total", "harga")
           .replace("nilai", "harga").replace("tarip", "tarif"))
    return s

def pick_guess(headers: List[str], candidates: List[str], default: Optional[str]=None) -> str:
    if not headers: return ""
    if default and default in headers: return default
    nh = {h: normalize(h) for h in headers}
    cand = [normalize(c) for c in candidates]
    for c in cand:
        for h, n in nh.items():
            if n == c: return h
    for c in cand:
        for h, n in nh.items():
            if c in n: return h
    return headers[0]

def parse_money(x: str) -> float:
    if x is None: return 0.0
    s = str(x).strip()
    if s == "": return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") and s.count("."):
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    else:
        if "," in s and "." not in s:
            if re.search(r",[0-9]{1,2}$", s): s = s.replace(",", ".")
            else: s = s.replace(",", "")
        elif "." in s and "," not in s:
            if not re.search(r"\.[0-9]{1,2}$", s): s = s.replace(".", "")
    try:
        return float(s)
    except Exception:
        s2 = re.sub(r"[^\d\-]", "", s)
        try: return float(s2)
        except Exception: return 0.0

def format_idr(n: float) -> str:
    s = f"{float(n):,.2f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")

def coerce_key(x: str) -> str:
    return re.sub(r"\s+", "", (x or "")).upper()

# ===== 3) XLSX READER (pure-Python) =====
NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

def _xlsx_col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if "A" <= ch <= "Z":
            n = n * 26 + (ord(ch) - 64)
    return n - 1

def _xlsx_ref_to_rc(ref: str) -> Tuple[int, int]:
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m: return 0, 0
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
        if first_sheet is None: return None
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

def read_xlsx_rows(b: bytes) -> List[Dict[str, str]]:
    z = ZipFile(io.BytesIO(b))
    sst = _xlsx_read_shared_strings(z)
    sheet_path = _xlsx_find_first_sheet_path(z)
    if not sheet_path: return []
    with z.open(sheet_path) as f:
        tree = ET.parse(f)
    root = tree.getroot()

    rows_dict: Dict[int, Dict[int, str]] = {}; max_col = -1
    for c in root.findall(".//main:c", NS):
        ref = c.attrib.get("r", "A1"); t = c.attrib.get("t")
        v = c.find("main:v", NS); is_node = c.find("main:is", NS)
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

    if not rows_dict: return []
    matrix: List[List[str]] = []
    for r in sorted(rows_dict.keys()):
        row = ["" for _ in range(max_col + 1)]
        for cidx, val in rows_dict[r].items():
            if 0 <= cidx <= max_col: row[cidx] = val
        matrix.append(row)

    header: List[str] = []; data_start = 0
    for i, row in enumerate(matrix):
        if any(cell.strip() for cell in row):
            header = row; data_start = i + 1; break
    if not header: return []

    seen = {}; final_header = []
    for h in header:
        base = (h or "").strip() or "COL"; name = base; k = 2
        while name.lower() in seen:
            name = f"{base}_{k}"; k += 1
        seen[name.lower()] = True; final_header.append(name)

    out: List[Dict[str, str]] = []
    for row in matrix[data_start:]:
        if not any(cell.strip() for cell in row): continue
        rec = {final_header[j]: row[j] if j < len(row) else "" for j in range(len(final_header))}
        out.append(rec)
    return out

# ===== 4) ITERATORS CSV/XLSX/ZIP =====
def iter_csv_rows_from_bytes(b: bytes) -> Generator[Dict[str, str], None, None]:
    delim = sniff_delimiter_from_bytes(b)
    tw = io.TextIOWrapper(io.BytesIO(b), encoding="utf-8", errors="ignore")
    reader = csv.DictReader(tw, delimiter=delim)
    for row in reader:
        yield {k: (v if v is not None else "") for k, v in row.items()}

def iter_xlsx_rows_from_bytes(b: bytes) -> Generator[Dict[str, str], None, None]:
    for r in read_xlsx_rows(b): yield r

def iter_uploaded_file_rows(f, errors: List[str]) -> Generator[Dict[str, str], None, None]:
    name = (f.name or "").lower()
    try:
        if name.endswith(".csv"):
            f.seek(0); b = f.read(); yield from iter_csv_rows_from_bytes(b)
        elif name.endswith((".xlsx", ".xlsm")):
            f.seek(0); b = f.read(); yield from iter_xlsx_rows_from_bytes(b)
        elif name.endswith(".zip"):
            f.seek(0); zb = io.BytesIO(f.read())
            with ZipFile(zb) as z:
                for zi in z.infolist():
                    if zi.is_dir(): continue
                    zname = zi.filename.lower()
                    if zname.endswith(".csv"):
                        try:
                            with z.open(zi) as fh: data = fh.read()
                            yield from iter_csv_rows_from_bytes(data)
                        except Exception as e:
                            errors.append(f"ZIP CSV `{zi.filename}`: {e}")
                    elif zname.endswith((".xlsx", ".xlsm")):
                        try:
                            with z.open(zi) as fh: data = fh.read()
                            yield from iter_xlsx_rows_from_bytes(data)
                        except Exception as e:
                            errors.append(f"ZIP XLSX `{zi.filename}`: {e}")
    except Exception as e:
        errors.append(f"{f.name}: {e}")

def read_headers_from_uploaded(files: List, errors: List[str]) -> List[str]:
    for f in files or []:
        n = (f.name or "").lower()
        try:
            if n.endswith(".csv"):
                f.seek(0); b = f.read()
                delim = sniff_delimiter_from_bytes(b)
                tw = io.TextIOWrapper(io.BytesIO(b), encoding="utf-8", errors="ignore")
                reader = csv.reader(tw, delimiter=delim)
                row = next(reader, [])
                if row: return [h.strip() for h in row]
            elif n.endswith((".xlsx", ".xlsm")):
                f.seek(0); b = f.read()
                rows = read_xlsx_rows(b)
                if rows: return list(rows[0].keys())
            elif n.endswith(".zip"):
                f.seek(0); zb = io.BytesIO(f.read())
                with ZipFile(zb) as z:
                    for zi in z.infolist():
                        if zi.is_dir(): continue
                        zname = zi.filename.lower()
                        if zname.endswith(".csv"):
                            try:
                                with z.open(zi) as fh: data = fh.read()
                                delim = sniff_delimiter_from_bytes(data)
                                tw = io.TextIOWrapper(io.BytesIO(data), encoding="utf-8", errors="ignore")
                                reader = csv.reader(tw, delimiter=delim)
                                row = next(reader, [])
                                if row: return [h.strip() for h in row]
                            except Exception: pass
                        elif zname.endswith((".xlsx", ".xlsm")):
                            try:
                                with z.open(zi) as fh: data = fh.read()
                                rows = read_xlsx_rows(data)
                                if rows: return list(rows[0].keys())
                            except Exception: pass
        except Exception as e:
            errors.append(f"Header `{f.name}`: {e}")
    return []

# ===== 5) XLSX WRITER (pure-Python) =====
def _col_letters(idx: int) -> str:
    s = ""; idx += 1
    while idx: idx, r = divmod(idx - 1, 26); s = chr(65 + r) + s
    return s

def _xml_escape(t: str) -> str:
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;")

def build_xlsx(columns: List[str], rows: List[Dict[str, str]], sheet_name: str="Rekonsiliasi") -> bytes:
    ws = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>']
    r = 1
    ws.append('<row r="1">' + "".join(
        f'<c r="{_col_letters(i)}1" t="inlineStr"><is><t xml:space="preserve">{_xml_escape(c)}</t></is></c>'
        for i, c in enumerate(columns)) + "</row>")
    for row in rows:
        r += 1
        ws.append(f'<row r="{r}">' + "".join(
            f'<c r="{_col_letters(i)}{r}" t="inlineStr"><is><t xml:space="preserve">{_xml_escape(str(row.get(c,"") or ""))}</t></is></c>'
            for i, c in enumerate(columns)) + "</row>")
    ws.append("</sheetData></worksheet>")
    sheet_xml = "\n".join(ws).encode("utf-8")
    content_types = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
    rels_root = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    wb_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="{_xml_escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''.encode("utf-8")
    wb_rels = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
    styles = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs>
</styleSheet>'''
    bio = io.BytesIO()
    with ZipFile(bio, "w") as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        z.writestr("xl/styles.xml", styles)
    return bio.getvalue()

# ===== 6) UI — UPLOADERS =====
with st.sidebar:
    st.header("1) Upload (CSV / XLSX / ZIP) — multiple")
    inv_files = st.file_uploader("📄 Invoice", type=["csv","xlsx","xlsm","zip"], accept_multiple_files=True)
    ts_files  = st.file_uploader("🎫 T-Summary", type=["csv","xlsx","xlsm","zip"], accept_multiple_files=True)
    st.caption("ZIP boleh berisi banyak CSV/XLSX. Untuk Excel, hanya sheet pertama yang dipakai.")
    csv_only = st.toggle("CSV-only mode (troubleshoot)", value=False, help="Lewati Excel/ZIP jika ON.")

if not inv_files or not ts_files:
    st.info("Unggah minimal satu file untuk **Invoice** dan **T-Summary**.")
    st.stop()

# ===== 7) HEADER MAPPING =====
errors: List[str] = []
try:
    inv_candidates = [f for f in inv_files if (f.name or "").lower().endswith((".csv",".xlsx",".xlsm",".zip")) and (not csv_only or (f.name or "").lower().endswith(".csv"))]
    ts_candidates  = [f for f in ts_files  if (f.name or "").lower().endswith((".csv",".xlsx",".xlsm",".zip")) and (not csv_only or (f.name or "").lower().endswith(".csv"))]
    inv_headers = read_headers_from_uploaded(inv_candidates, errors)
    ts_headers  = read_headers_from_uploaded(ts_candidates, errors)
    if not inv_headers: st.error("Tidak bisa mendeteksi header dari Invoice."); st.stop()
    if not ts_headers:  st.error("Tidak bisa mendeteksi header dari T-Summary."); st.stop()
except Exception:
    st.error("❌ Gagal saat deteksi header.")
    st.code(traceback.format_exc()); st.stop()

st.subheader("2) Pemetaan Kolom")
c1, c2 = st.columns(2)
with c1:
    st.markdown("**Invoice**")
    inv_key = st.selectbox("Nomor Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["nomor invoice","no invoice","invoice"])) if inv_headers else 0)
    inv_amt = st.selectbox("Nominal/Harga",  inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["harga","nominal","amount","total"])) if inv_headers else 0)
    inv_tgl_inv = st.selectbox("Tanggal Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["tanggal invoice","tgl invoice","tanggal","tgl"])) if inv_headers else 0)
    inv_pay_inv = st.selectbox("Tanggal Pembayaran Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["tanggal invoice","pembayaran","tgl pembayaran"])) if inv_headers else 0)
    inv_tujuan  = st.selectbox("Tujuan", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["tujuan","destination"])) if inv_headers else 0)
    inv_channel = st.selectbox("Channel", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["channel"])) if inv_headers else 0)
    inv_merchant= st.selectbox("Merchant", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["merchant","mid"])) if inv_headers else 0)
with c2:
    st.markdown("**T-Summary**")
    ts_key   = st.selectbox("Nomor Invoice (T-Summary)", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["nomor invoice","no invoice","invoice"])) if ts_headers else 0)
    ts_amt   = st.selectbox("Nominal/Tarif", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["tarif","harga","nominal","amount","total"])) if ts_headers else 0)
    ts_kode  = st.selectbox("Kode Booking", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["kode booking","kode boking"])) if ts_headers else 0)
    ts_tiket = st.selectbox("Nomor Tiket", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["nomor tiket","no tiket","ticket"])) if ts_headers else 0)
    ts_pay   = st.selectbox("Tanggal Pembayaran T-Summary", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["pembayaran","tanggal pembayaran","tgl pembayaran"])) if ts_headers else 0)
    ts_gol   = st.selectbox("Golongan", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["golongan","kelas"])) if ts_headers else 0)
    ts_asal  = st.selectbox("Keberangkatan / Asal", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["asal","keberangkatan"])) if ts_headers else 0)
    ts_cetak = st.selectbox("Tgl Cetak Boarding Pass", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["cetak boarding pass","tgl cetak"])) if ts_headers else 0)

only_diff  = st.checkbox("Hanya Selisih ≠ 0", value=False)
show_table = st.checkbox("Tampilkan Tabel (opsional)", value=False)
go = st.button("🚀 Proses")

# ===== 8) PROSES (dengan guard error) =====
if go:
    try:
        # Aggregates
        agg_inv: Dict[str, float] = {}
        agg_ts:  Dict[str, float] = {}
        keys_order: List[str] = []
        seen_keys = set()

        inv_first = {"tgl_inv": {}, "pay_inv": {}, "tujuan": {}, "channel": {}, "merchant": {}}
        ts_first  = {"pay_ts": {}}
        ts_join   = {"kode": {}, "tiket": {}, "gol": {}, "asal": {}, "cetak": {}}

        def add_join(store: Dict[str, set], k: str, v: str):
            if not v: return
            s = store.setdefault(k, set()); s.add(v)

        # Invoice input
        for f in inv_candidates:
            for row in iter_uploaded_file_rows(f, errors):
                key = coerce_key(row.get(inv_key, ""))
                if not key: continue
                if key not in seen_keys:
                    seen_keys.add(key); keys_order.append(key)
                agg_inv[key] = agg_inv.get(key, 0.0) + parse_money(row.get(inv_amt, "0"))
                if key not in inv_first["tgl_inv"] and row.get(inv_tgl_inv, ""): inv_first["tgl_inv"][key] = row.get(inv_tgl_inv, "")
                if key not in inv_first["pay_inv"] and row.get(inv_pay_inv, ""): inv_first["pay_inv"][key] = row.get(inv_pay_inv, "")
                if key not in inv_first["tujuan"]  and row.get(inv_tujuan, ""):  inv_first["tujuan"][key]  = row.get(inv_tujuan, "")
                if key not in inv_first["channel"] and row.get(inv_channel, ""): inv_first["channel"][key] = row.get(inv_channel, "")
                if key not in inv_first["merchant"]and row.get(inv_merchant, ""):inv_first["merchant"][key]= row.get(inv_merchant, "")

        # T-Summary input
        for f in ts_candidates:
            for row in iter_uploaded_file_rows(f, errors):
                key = coerce_key(row.get(ts_key, ""))
                if not key: continue
                agg_ts[key] = agg_ts.get(key, 0.0) + parse_money(row.get(ts_amt, "0"))
                if key not in ts_first["pay_ts"] and row.get(ts_pay, ""): ts_first["pay_ts"][key] = row.get(ts_pay, "")
                add_join(ts_join["kode"], key, row.get(ts_kode, ""))
                add_join(ts_join["tiket"], key, row.get(ts_tiket, ""))
                add_join(ts_join["gol"], key, row.get(ts_gol, ""))
                add_join(ts_join["asal"], key, row.get(ts_asal, ""))
                add_join(ts_join["cetak"], key, row.get(ts_cetak, ""))

        # Build result
        out_rows: List[Dict[str, str]] = []
        total_inv = total_ts = total_diff = 0.0
        naik = turun = sama = 0

        for k in keys_order:
            v_inv = float(agg_inv.get(k, 0.0))
            v_ts  = float(agg_ts.get(k, 0.0))
            diff = v_inv - v_ts

            # Kategori: Invoice > T-Summary => "Turun"
            if v_inv > v_ts: cat = "Turun"
            elif v_inv < v_ts: cat = "Naik"
            else: cat = "Sama"

            row = {
                "Tanggal Invoice":              inv_first["tgl_inv"].get(k, ""),
                "Nomor Invoice":                k,
                "Kode Booking":                 ", ".join(sorted(ts_join["kode"].get(k, []))),
                "Nomor Tiket":                  ", ".join(sorted(ts_join["tiket"].get(k, []))),
                "Nominal Invoice (SUMIFS)":     format_idr(v_inv),
                "Tanggal Pembayaran Invoice":   inv_first["pay_inv"].get(k, ""),
                "Nominal T-Summary (SUMIFS)":   format_idr(v_ts),
                "Tanggal Pembayaran T-Summary": ts_first["pay_ts"].get(k, ""),
                "Golongan":                     ", ".join(sorted(ts_join["gol"].get(k, []))),
                "Keberangkatan":                ", ".join(sorted(ts_join["asal"].get(k, []))),  # Asal dari T-Summary
                "Tujuan":                       inv_first["tujuan"].get(k, ""),                 # Tujuan dari Invoice
                "Tgl Cetak Boarding Pass":      ", ".join(sorted(ts_join["cetak"].get(k, []))),
                "Channel":                      inv_first["channel"].get(k, ""),
                "Merchant":                     inv_first["merchant"].get(k, ""),
                "Selisih":                      format_idr(diff),
                "Kategori":                     cat,
            }
            if (not only_diff) or (diff != 0):
                out_rows.append(row)

            total_inv += v_inv; total_ts += v_ts; total_diff += diff
            if cat == "Naik": naik += 1
            elif cat == "Turun": turun += 1
            else: sama += 1

        # Metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Invoice (SUMIFS)", format_idr(total_inv))
        m2.metric("Total T-Summary (SUMIFS)", format_idr(total_ts))
        m3.metric("Total Selisih (Inv − T)", format_idr(total_diff))
        m4.metric("Naik / Turun / Sama", f"{naik} / {turun} / {sama}")

        display_cols = [
            "Tanggal Invoice","Nomor Invoice","Kode Booking","Nomor Tiket",
            "Nominal Invoice (SUMIFS)","Tanggal Pembayaran Invoice",
            "Nominal T-Summary (SUMIFS)","Tanggal Pembayaran T-Summary",
            "Golongan","Keberangkatan","Tujuan","Tgl Cetak Boarding Pass",
            "Channel","Merchant","Selisih","Kategori",
        ]

        if show_table:
            st.data_editor(out_rows, use_container_width=True, disabled=True, column_order=display_cols)

        # Download Excel
        xlsx_bytes = build_xlsx(display_cols, out_rows, sheet_name="Rekonsiliasi")
        st.download_button("⬇️ Download Excel (.xlsx)", data=xlsx_bytes,
                           file_name="rekonsiliasi_ultra_safe.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Notes / diagnostics
        if errors:
            with st.expander("⚠️ Catatan pembacaan file", expanded=False):
                for e in errors: st.caption(f"– {e}")

    except Exception:
        st.error("❌ Terjadi error saat menjalankan proses.")
        st.code(traceback.format_exc())
