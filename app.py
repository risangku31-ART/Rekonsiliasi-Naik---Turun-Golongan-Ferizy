# file: app_xlsb_only_auto_header.py
# Rekonsiliasi Naik/Turun Golongan — .xlsb only + auto header + sheet picker + export XLSX
# Requirements minimal:
#   streamlit>=1.26
#   pyxlsb>=1.0.10

import io
import re
import tempfile
import traceback
from typing import Dict, List, Tuple, Optional
import streamlit as st

# ---------- Page ----------
st.set_page_config(page_title="Rekonsiliasi (.xlsb only + Auto Header)", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan — (.xlsb only + Auto Header)")

st.markdown("""
<style>
div[data-testid="stMetricLabel"] { font-size: 11px !important; }
div[data-testid="stMetricValue"] { font-size: 17px !important; }
div[data-testid="stMetricValue"] > div { white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; }
.small-note { font-size: 12px; opacity: .8 }
</style>
""", unsafe_allow_html=True)

# ---------- Utils ----------
def pyxlsb_available() -> bool:
    try:
        import pyxlsb  # noqa
        return True
    except Exception:
        return False

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

def normalize(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = (s.replace("no ", "nomor ").replace("no", "nomor")
           .replace("inv", "invoice").replace("harga total", "harga")
           .replace("nilai", "harga").replace("tarip", "tarif"))
    return s

def pick_guess(headers: List[str], candidates: List[str]) -> str:
    if not headers: return ""
    nh = {h: normalize(h) for h in headers}
    for c in [normalize(c) for c in candidates]:
        for h, n in nh.items():
            if n == c or c in n: return h
    return headers[0]

# ---------- XLSB Reader + Auto Header ----------
class SheetMode:
    AUTO = "Auto (sheet pertama berisi data)"
    BY_NAME = "Nama sheet"
    BY_INDEX = "Index (1-based)"

def _xlsb_rows_from_sheet_bytes(b: bytes, sheet_mode: str,
                                sheet_name: Optional[str],
                                sheet_index: Optional[int]) -> List[List[str]]:
    """Baca semua baris non-kosong dari sheet terpilih."""
    try:
        import pyxlsb  # type: ignore
    except Exception:
        return []
    rows: List[List[str]] = []
    with tempfile.NamedTemporaryFile(delete=True, suffix=".xlsb") as tmp:
        tmp.write(b); tmp.flush()
        import pyxlsb  # type: ignore
        with pyxlsb.open_workbook(tmp.name) as wb:
            def read_sheet(sh) -> List[List[str]]:
                out = []
                with wb.get_sheet(sh) as sheet:
                    for r in sheet.rows():
                        vals = [(c.v if c is not None else "") for c in r]
                        if any(str(x).strip() for x in vals):
                            out.append([str(x or "") for x in vals])
                return out

            if sheet_mode == SheetMode.BY_NAME and sheet_name:
                try: rows = read_sheet(sheet_name)
                except Exception: rows = []
            elif sheet_mode == SheetMode.BY_INDEX and sheet_index and sheet_index > 0:
                try: rows = read_sheet(sheet_index)
                except Exception: rows = []
            else:
                for i in range(1, 41):
                    try:
                        tmp_rows = read_sheet(i)
                        if tmp_rows:
                            rows = tmp_rows; break
                    except Exception:
                        continue
    return rows

def _is_numeric_like(val: str) -> bool:
    return bool(re.fullmatch(r"\s*[\d\.\,\-\(\)]+\s*", str(val or "")))

def _infer_header_and_records(ne_rows: List[List[str]]) -> Tuple[List[str], List[Dict[str, str]]]:
    """Auto header: pilih baris pertama yang tidak numeric-heavy; fallback ke baris non-kosong pertama."""
    if not ne_rows: return [], []
    header_idx = 0
    for i, row in enumerate(ne_rows):
        nonempty = [c for c in row if str(c).strip() != ""]
        if len(nonempty) < 2:  # sangat kecil, loncati
            continue
        nums = sum(1 for c in nonempty if _is_numeric_like(c))
        if nums <= len(nonempty) // 2:  # mayoritas bukan angka -> kandidat header
            header_idx = i
            break
    header_raw = ne_rows[header_idx]
    header = [str(x).strip() or f"COL{j+1}" for j, x in enumerate(header_raw)]
    # pastikan unik
    seen: Dict[str, int] = {}
    cols: List[str] = []
    for h in header:
        key = h or "COL"
        seen[key] = seen.get(key, 0) + 1
        cols.append(key if seen[key] == 1 else f"{key}_{seen[key]}")

    records: List[Dict[str, str]] = []
    for r in ne_rows[header_idx + 1:]:
        if not any(str(x).strip() for x in r):
            continue
        row = [str(x or "") for x in r] + [""] * (len(cols) - len(r))
        records.append({cols[j]: row[j] for j in range(len(cols))})
    return cols, records

def read_xlsb_records(f, sheet_mode: str, sheet_name: Optional[str], sheet_index: Optional[int], errors: List[str]) -> Tuple[List[str], List[Dict[str, str]]]:
    try:
        f.seek(0); b = f.read()
        ne_rows = _xlsb_rows_from_sheet_bytes(b, sheet_mode, sheet_name, sheet_index)
        if not ne_rows:
            errors.append(f"{f.name}: sheet kosong / tidak terbaca.")
            return [], []
        return _infer_header_and_records(ne_rows)
    except Exception as e:
        errors.append(f"{f.name}: {e}")
        return [], []

# ---------- XLSX Writer (Download) ----------
def _col_letters(idx: int) -> str:
    s = ""; idx += 1
    while idx: idx, r = divmod(idx - 1, 26); s = chr(65 + r) + s
    return s

def _xml_escape(t: str) -> str:
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;")

def build_xlsx(columns: List[str], rows: List[Dict[str, str]], sheet_name: str="Rekonsiliasi") -> bytes:
    from zipfile import ZipFile
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

# ---------- Uploaders (HANYA .xlsb) ----------
with st.sidebar:
    st.header("1) Upload (.xlsb) — multiple")
    inv_files = st.file_uploader("📄 Invoice (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    ts_files  = st.file_uploader("🎫 T-Summary (.xlsb)", type=["xlsb"], accept_multiple_files=True)

# ---------- Sheet Settings ----------
class _SM:  # kecilkan namespace di widget
    AUTO = SheetMode.AUTO
    NAME = SheetMode.BY_NAME
    IDX  = SheetMode.BY_INDEX

st.subheader("2) Pengaturan Sheet")
csm1, csm2 = st.columns(2)
with csm1:
    inv_sheet_mode = st.selectbox("Mode Sheet (Invoice)", [_SM.AUTO, _SM.NAME, _SM.IDX], index=0)
    inv_sheet_name = st.text_input("Nama Sheet (Invoice)", value="") if inv_sheet_mode == _SM.NAME else ""
    inv_sheet_idx  = st.number_input("Index Sheet (Invoice)", min_value=1, value=1, step=1) if inv_sheet_mode == _SM.IDX else None
with csm2:
    ts_sheet_mode  = st.selectbox("Mode Sheet (T-Summary)", [_SM.AUTO, _SM.NAME, _SM.IDX], index=0)
    ts_sheet_name  = st.text_input("Nama Sheet (T-Summary)", value="") if ts_sheet_mode == _SM.NAME else ""
    ts_sheet_idx   = st.number_input("Index Sheet (T-Summary)", min_value=1, value=1, step=1) if ts_sheet_mode == _SM.IDX else None

# ---------- Guard: pyxlsb must exist ----------
if not pyxlsb_available():
    st.error("`pyxlsb` belum terpasang. Tambahkan ke `requirements.txt`:\n\n```\nstreamlit>=1.26\npyxlsb>=1.0.10\n```")
    st.download_button("⬇️ Download requirements.txt",
                       data=b"streamlit>=1.26\npyxlsb>=1.0.10\n",
                       file_name="requirements.txt", mime="text/plain")
    st.stop()

if not inv_files or not ts_files:
    st.info("Unggah minimal satu file **Invoice (.xlsb)** dan satu file **T-Summary (.xlsb)**.")
    st.stop()

# ---------- Header Detection (AUTO) ----------
errors: List[str] = []
def detect_headers(files, sheet_mode, sheet_name, sheet_index) -> List[str]:
    for f in files or []:
        hdr, _ = read_xlsb_records(f, sheet_mode, sheet_name, sheet_index, errors)
        if hdr: return hdr
    return []

inv_headers = detect_headers(inv_files, inv_sheet_mode, inv_sheet_name, inv_sheet_idx)
ts_headers  = detect_headers(ts_files,  ts_sheet_mode,  ts_sheet_name,  ts_sheet_idx)

if not inv_headers:
    st.error("Tidak bisa mendeteksi header dari Invoice (.xlsb). Periksa pilihan sheet.")
    st.stop()
if not ts_headers:
    st.error("Tidak bisa mendeteksi header dari T-Summary (.xlsb). Periksa pilihan sheet.")
    st.stop()

# ---------- Mapping ----------
st.subheader("3) Pemetaan Kolom")
c3, c4 = st.columns(2)
with c3:
    st.markdown("**Invoice**")
    inv_key = st.selectbox("Nomor Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["nomor invoice","no invoice","invoice"])) if inv_headers else 0)
    inv_amt = st.selectbox("Nominal/Harga",  inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["harga","nominal","amount","total"])) if inv_headers else 0)
    inv_tgl_inv = st.selectbox("Tanggal Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["tanggal invoice","tgl invoice","tanggal","tgl"])) if inv_headers else 0)
    inv_pay_inv = st.selectbox("Tanggal Pembayaran Invoice", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["pembayaran","tgl pembayaran","tanggal invoice"])) if inv_headers else 0)
    inv_tujuan  = st.selectbox("Tujuan", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["tujuan","destination"])) if inv_headers else 0)
    inv_channel = st.selectbox("Channel", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["channel"])) if inv_headers else 0)
    inv_merchant= st.selectbox("Merchant", inv_headers, index=inv_headers.index(pick_guess(inv_headers, ["merchant","mid"])) if inv_headers else 0)
with c4:
    st.markdown("**T-Summary**")
    ts_key   = st.selectbox("Nomor Invoice (T-Summary)", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["nomor invoice","no invoice","invoice"])) if ts_headers else 0)
    ts_amt   = st.selectbox("Nominal/Tarif", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["tarif","harga","nominal","amount","total"])) if ts_headers else 0)
    ts_kode  = st.selectbox("Kode Booking", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["kode booking","kode boking"])) if ts_headers else 0)
    ts_tiket = st.selectbox("Nomor Tiket", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["nomor tiket","no tiket","ticket"])) if ts_headers else 0)
    ts_pay   = st.selectbox("Tanggal Pembayaran T-Summary", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["pembayaran","tanggal pembayaran","tgl pembayaran"])) if ts_headers else 0)
    ts_gol   = st.selectbox("Golongan", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["golongan","kelas"])) if ts_headers else 0)
    ts_asal  = st.selectbox("Keberangkatan / Asal", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["asal","keberangkatan"])) if ts_headers else 0)
    ts_cetak = st.selectbox("Tgl Cetak Boarding Pass", ts_headers, index=ts_headers.index(pick_guess(ts_headers, ["cetak boarding pass","tgl cetak"])) if ts_headers else 0)

# ---------- Opsi Hasil ----------
copt1, copt2 = st.columns(2)
with copt1:
    only_diff  = st.checkbox("Hanya Selisih ≠ 0", value=False)
with copt2:
    tol = st.number_input("Toleransi selisih (Rp)", min_value=0.0, value=0.0, step=1000.0, format="%.0f")

go = st.button("🚀 Proses")

# ---------- Proses ----------
def add_join(store: Dict[str, set], k: str, v: str):
    if not v: return
    store.setdefault(k, set()).add(v)

def read_all(files, sheet_mode, sheet_name, sheet_index) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    for f in files:
        hdr, recs = read_xlsb_records(f, sheet_mode, sheet_name, sheet_index, errors)
        out.extend(recs)
    return out

if go:
    try:
        inv_recs = read_all(inv_files, inv_sheet_mode, inv_sheet_name, inv_sheet_idx)
        ts_recs  = read_all(ts_files,  ts_sheet_mode,  ts_sheet_name,  ts_sheet_idx)

        agg_inv: Dict[str, float] = {}
        agg_ts:  Dict[str, float] = {}
        keys_order: List[str] = []; seen = set()

        inv_first = {"tgl_inv": {}, "pay_inv": {}, "tujuan": {}, "channel": {}, "merchant": {}}
        ts_first  = {"pay_ts": {}}
        ts_join   = {"kode": {}, "tiket": {}, "gol": {}, "asal": {}, "cetak": {}}

        # Invoice basis (patokan awal)
        for row in inv_recs:
            key = coerce_key(row.get(inv_key, ""))
            if not key: continue
            if key not in seen: seen.add(key); keys_order.append(key)
            agg_inv[key] = agg_inv.get(key, 0.0) + parse_money(row.get(inv_amt, "0"))
            if key not in inv_first["tgl_inv"] and row.get(inv_tgl_inv, ""): inv_first["tgl_inv"][key] = row.get(inv_tgl_inv, "")
            if key not in inv_first["pay_inv"] and row.get(inv_pay_inv, ""): inv_first["pay_inv"][key] = row.get(inv_pay_inv, "")
            if key not in inv_first["tujuan"]  and row.get(inv_tujuan, ""):  inv_first["tujuan"][key]  = row.get(inv_tujuan, "")
            if key not in inv_first["channel"] and row.get(inv_channel, ""): inv_first["channel"][key] = row.get(inv_channel, "")
            if key not in inv_first["merchant"]and row.get(inv_merchant, ""):inv_first["merchant"][key]= row.get(inv_merchant, "")

        # T-Summary
        for row in ts_recs:
            key = coerce_key(row.get(ts_key, ""))
            if not key: continue
            agg_ts[key] = agg_ts.get(key, 0.0) + parse_money(row.get(ts_amt, "0"))
            if key not in ts_first["pay_ts"] and row.get(ts_pay, ""): ts_first["pay_ts"][key] = row.get(ts_pay, "")
            add_join(ts_join["kode"], key, row.get(ts_kode, ""))
            add_join(ts_join["tiket"], key, row.get(ts_tiket, ""))
            add_join(ts_join["gol"], key, row.get(ts_gol, ""))
            add_join(ts_join["asal"], key, row.get(ts_asal, ""))
            add_join(ts_join["cetak"], key, row.get(ts_cetak, ""))

        # Hasil
        out_rows: List[Dict[str, str]] = []
        total_inv = total_ts = total_diff = 0.0
        naik = turun = sama = 0

        for k in keys_order:
            v_inv = float(agg_inv.get(k, 0.0))
            v_ts  = float(agg_ts.get(k, 0.0))
            diff = v_inv - v_ts
            cat = "Turun" if v_inv > v_ts else ("Naik" if v_inv < v_ts else "Sama")

            if only_diff and diff == 0: 
                continue
            if tol > 0 and abs(diff) <= tol:
                continue

            row = {
                "Tanggal Invoice":              inv_first["tgl_inv"].get(k, ""),
                "Nomor Invoice":                k,
                "Kode Booking":                 ", ".join(sorted(ts_join["kode"].get(k, []))),
                "Nomor Tiket":                  ", ".join(sorted(ts_join["tiket"].get(k, []))),
                "Invoice (Nominal; Tgl Bayar)": f"{format_idr(v_inv)}; {inv_first['pay_inv'].get(k, '')}",
                "T-Summary (Nominal; Tgl Bayar)": f"{format_idr(v_ts)}; {ts_first['pay_ts'].get(k, '')}",
                "Golongan":                     ", ".join(sorted(ts_join["gol"].get(k, []))),
                "Keberangkatan":                ", ".join(sorted(ts_join["asal"].get(k, []))),  # Asal (T-Summary)
                "Tujuan":                       inv_first["tujuan"].get(k, ""),                 # Tujuan (Invoice)
                "Tgl Cetak Boarding Pass":      ", ".join(sorted(ts_join["cetak"].get(k, []))),
                "Channel":                      inv_first["channel"].get(k, ""),
                "Merchant":                     inv_first["merchant"].get(k, ""),
                "Selisih":                      format_idr(diff),
                "Kategori":                     cat,
            }
            out_rows.append(row)

            total_inv += v_inv; total_ts += v_ts; total_diff += diff
            naik += (cat == "Naik"); turun += (cat == "Turun"); sama += (cat == "Sama")

        # Metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Invoice (SUMIFS)", format_idr(total_inv))
        m2.metric("Total T-Summary (SUMIFS)", format_idr(total_ts))
        m3.metric("Total Selisih (Inv − T)", format_idr(total_diff))
        m4.metric("Naik / Turun / Sama", f"{int(naik)} / {int(turun)} / {int(sama)}")

        # Download
        display_cols = [
            "Tanggal Invoice","Nomor Invoice","Kode Booking","Nomor Tiket",
            "Invoice (Nominal; Tgl Bayar)","T-Summary (Nominal; Tgl Bayar)",
            "Golongan","Keberangkatan","Tujuan","Tgl Cetak Boarding Pass",
            "Channel","Merchant","Selisih","Kategori",
        ]
        xlsx_bytes = build_xlsx(display_cols, out_rows, sheet_name="Rekonsiliasi")
        st.download_button("⬇️ Download Excel (.xlsx)", data=xlsx_bytes,
                           file_name="rekonsiliasi_xlsb_auto_header.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if errors:
            with st.expander("⚠️ Catatan pembacaan file", expanded=False):
                for e in errors: st.caption(f"– {e}")

    except Exception:
        st.error("❌ Terjadi error saat proses.")
        st.code(traceback.format_exc())
