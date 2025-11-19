# file: app_xlsb_only.py
# Rekonsiliasi Naik/Turun Golongan — Uploader HANYA .xlsb

import io
import re
import csv
import tempfile
from zipfile import ZipFile  # (tidak dipakai, tapi dibiarkan untuk mudah extend)
from typing import Dict, List, Generator, Tuple
import streamlit as st

# ---------------- Config UI ----------------
st.set_page_config(page_title="Rekonsiliasi (.xlsb only)", layout="wide")
st.title("🔄 Rekonsiliasi Naik/Turun Golongan — (.xlsb only)")

# kecilkan metric (why: angka panjang)
st.markdown("""
<style>
div[data-testid="stMetricLabel"] { font-size: 11px !important; }
div[data-testid="stMetricValue"] { font-size: 17px !important; }
div[data-testid="stMetricValue"] > div { white-space: nowrap !important; overflow: visible !important; text-overflow: clip !important; }
</style>
""", unsafe_allow_html=True)

# ---------------- Utils ----------------
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

def parse_header_override(s: str) -> List[str]:
    s = (s or "").strip()
    if not s: return []
    if "\t" in s: parts = [p.strip() for p in s.split("\t")]
    elif ";" in s: parts = [p.strip() for p in s.split(";")]
    else: parts = [p.strip() for p in s.split(",")]
    return [p for p in parts if p != ""]

# ---------------- XLSB Reader ----------------
def _xlsb_first_nonempty_sheet_rows(name: str, b: bytes) -> List[List[str]]:
    """Cari sheet pertama yang berisi data (why: beberapa file header ada di sheet selain index 1)."""
    try:
        import pyxlsb  # type: ignore
    except Exception:
        return []
    rows_nonempty: List[List[str]] = []
    with tempfile.NamedTemporaryFile(delete=True, suffix=".xlsb") as tmp:
        tmp.write(b); tmp.flush()
        import pyxlsb  # type: ignore
        with pyxlsb.open_workbook(tmp.name) as wb:
            sheet_names = getattr(wb, "sheets", None) or []
            for sname in sheet_names or [1]:
                try:
                    with wb.get_sheet(sname) as sh:
                        tmp_rows = []
                        for r in sh.rows():
                            vals = [(c.v if c is not None else "") for c in r]
                            if any(str(x).strip() for x in vals):
                                tmp_rows.append([str(x or "") for x in vals])
                        if tmp_rows:
                            rows_nonempty = tmp_rows
                            break
                except Exception:
                    continue
    return rows_nonempty

def iter_xlsb_with_header(name: str, b: bytes, header_override: List[str], skip_rows_before_header: int, errors: List[str]) -> Generator[Dict[str, str], None, None]:
    if not pyxlsb_available():
        errors.append(f"{name}: pyxlsb tidak tersedia.")
        return
    ne_rows = _xlsb_first_nonempty_sheet_rows(name, b)
    if not ne_rows or len(ne_rows) <= skip_rows_before_header:
        errors.append(f"{name}: tidak ada baris data terdeteksi.")
        return
    if header_override:
        header = header_override
        data_rows = ne_rows[skip_rows_before_header:]
    else:
        header = [str(x).strip() for x in ne_rows[skip_rows_before_header]]
        data_rows = ne_rows[skip_rows_before_header + 1:]
    maxw = max(len(header), max((len(r) for r in data_rows), default=0))
    header = header + [f"COL{j+1}" for j in range(len(header), maxw)]
    for r in data_rows:
        r = list(r) + [""] * (maxw - len(r))
        yield {header[j]: r[j] for j in range(maxw)}

def detect_headers_xlsb(files: List, header_override: List[str], skip_rows_before_header: int, errors: List[str]) -> List[str]:
    if header_override: return header_override
    for f in files or []:
        try:
            f.seek(0); b = f.read()
            ne = _xlsb_first_nonempty_sheet_rows(f.name, b)
            if len(ne) > skip_rows_before_header:
                return [str(x).strip() for x in ne[skip_rows_before_header]]
        except Exception as e:
            errors.append(f"Header {f.name}: {e}")
    return []

def iter_uploaded_xlsb_rows(f, errors: List[str], header_override: List[str], skip_rows_before_header: int) -> Generator[Dict[str, str], None, None]:
    name = (f.name or "")
    try:
        f.seek(0); b = f.read()
        yield from iter_xlsb_with_header(name, b, header_override, skip_rows_before_header, errors)
    except Exception as e:
        errors.append(f"{name}: {e}")

# ---------------- XLSX Writer (Export) ----------------
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

# ---------------- Uploaders (HANYA .xlsb) ----------------
with st.sidebar:
    st.header("1) Upload (.xlsb) — multiple")
    inv_files = st.file_uploader("📄 Invoice (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    ts_files  = st.file_uploader("🎫 T-Summary (.xlsb)", type=["xlsb"], accept_multiple_files=True)

st.subheader("2) Pengaturan Header")
c1, c2 = st.columns(2)
with c1:
    inv_skip = st.number_input("Lewati baris awal (Invoice)", min_value=0, value=0, step=1)
    inv_hdr_manual = parse_header_override(st.text_input("Paksa header manual (Invoice)", value=""))
with c2:
    ts_skip = st.number_input("Lewati baris awal (T-Summary)", min_value=0, value=0, step=1)
    ts_hdr_manual = parse_header_override(st.text_input("Paksa header manual (T-Summary)", value=""))

# Hard guard: pyxlsb wajib ada
if not pyxlsb_available():
    st.error("pyxlsb belum terpasang. Tambahkan ke requirements.txt:\n\n`streamlit>=1.26\\npyxlsb>=1.0.10`")
    st.download_button("⬇️ Download requirements.txt", data=b"streamlit>=1.26\npyxlsb>=1.0.10\n",
                       file_name="requirements.txt", mime="text/plain")
    st.stop()

if not inv_files or not ts_files:
    st.info("Unggah minimal satu file **Invoice (.xlsb)** dan satu file **T-Summary (.xlsb)**.")
    st.stop()

# ---------------- Header detection (xlsb only) ----------------
errors: List[str] = []
inv_headers = detect_headers_xlsb(inv_files, inv_hdr_manual, inv_skip, errors)
ts_headers  = detect_headers_xlsb(ts_files,  ts_hdr_manual,  ts_skip,  errors)

if not inv_headers:
    st.error("Tidak bisa mendeteksi header dari Invoice (.xlsb). Atur **Lewati baris** atau isi **Paksa header manual**.")
    st.stop()
if not ts_headers:
    st.error("Tidak bisa mendeteksi header dari T-Summary (.xlsb). Atur **Lewati baris** atau isi **Paksa header manual**.")
    st.stop()

# ---------------- Mapping ----------------
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

only_diff  = st.checkbox("Hanya Selisih ≠ 0", value=False)
show_table = st.checkbox("Tampilkan tabel hasil", value=False)
go = st.button("🚀 Proses")

# ---------------- Proses ----------------
def add_join(store: Dict[str, set], k: str, v: str):
    if not v: return
    store.setdefault(k, set()).add(v)

def iter_all(files, hdr_override, skip_before):
    for f in files:
        yield from iter_uploaded_xlsb_rows(f, errors, hdr_override, skip_before)

if go:
    try:
        agg_inv: Dict[str, float] = {}
        agg_ts:  Dict[str, float] = {}
        keys_order: List[str] = []; seen = set()

        inv_first = {"tgl_inv": {}, "pay_inv": {}, "tujuan": {}, "channel": {}, "merchant": {}}
        ts_first  = {"pay_ts": {}}
        ts_join   = {"kode": {}, "tiket": {}, "gol": {}, "asal": {}, "cetak": {}}

        # Invoice
        for row in iter_all(inv_files, inv_hdr_manual, inv_skip):
            key = coerce_key(row.get(inv_key, ""))
            if not key: continue
            if key not in seen:
                seen.add(key); keys_order.append(key)
            agg_inv[key] = agg_inv.get(key, 0.0) + parse_money(row.get(inv_amt, "0"))
            if key not in inv_first["tgl_inv"] and row.get(inv_tgl_inv, ""): inv_first["tgl_inv"][key] = row.get(inv_tgl_inv, "")
            if key not in inv_first["pay_inv"] and row.get(inv_pay_inv, ""): inv_first["pay_inv"][key] = row.get(inv_pay_inv, "")
            if key not in inv_first["tujuan"]  and row.get(inv_tujuan, ""):  inv_first["tujuan"][key]  = row.get(inv_tujuan, "")
            if key not in inv_first["channel"] and row.get(inv_channel, ""): inv_first["channel"][key] = row.get(inv_channel, "")
            if key not in inv_first["merchant"]and row.get(inv_merchant, ""):inv_first["merchant"][key]= row.get(inv_merchant, "")

        # T-Summary
        for row in iter_all(ts_files, ts_hdr_manual, ts_skip):
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
                "Keberangkatan":                ", ".join(sorted(ts_join["asal"].get(k, []))),  # Asal (T-Summary)
                "Tujuan":                       inv_first["tujuan"].get(k, ""),                 # Tujuan (Invoice)
                "Tgl Cetak Boarding Pass":      ", ".join(sorted(ts_join["cetak"].get(k, []))),
                "Channel":                      inv_first["channel"].get(k, ""),
                "Merchant":                     inv_first["merchant"].get(k, ""),
                "Selisih":                      format_idr(diff),
                "Kategori":                     cat,
            }
            if (not only_diff) or (diff != 0): out_rows.append(row)
            total_inv += v_inv; total_ts += v_ts; total_diff += diff
            naik += (cat == "Naik"); turun += (cat == "Turun"); sama += (cat == "Sama")

        # Metrics
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Invoice (SUMIFS)", format_idr(total_inv))
        m2.metric("Total T-Summary (SUMIFS)", format_idr(total_ts))
        m3.metric("Total Selisih (Inv − T)", format_idr(total_diff))
        m4.metric("Naik / Turun / Sama", f"{int(naik)} / {int(turun)} / {int(sama)}")

        # Tabel & Download
        display_cols = [
            "Tanggal Invoice","Nomor Invoice","Kode Booking","Nomor Tiket",
            "Nominal Invoice (SUMIFS)","Tanggal Pembayaran Invoice",
            "Nominal T-Summary (SUMIFS)","Tanggal Pembayaran T-Summary",
            "Golongan","Keberangkatan","Tujuan","Tgl Cetak Boarding Pass",
            "Channel","Merchant","Selisih","Kategori",
        ]
        if show_table:
            st.data_editor(out_rows, use_container_width=True, disabled=True, column_order=display_cols)

        xlsx_bytes = build_xlsx(display_cols, out_rows, sheet_name="Rekonsiliasi")
        st.download_button("⬇️ Download Excel (.xlsx)", data=xlsx_bytes,
                           file_name="rekonsiliasi_xlsb_only.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if errors:
            with st.expander("⚠️ Catatan pembacaan file", expanded=False):
                for e in errors: st.caption(f"– {e}")

    except Exception as e:
        st.error("❌ Terjadi error saat proses.")
        st.code(str(e))
