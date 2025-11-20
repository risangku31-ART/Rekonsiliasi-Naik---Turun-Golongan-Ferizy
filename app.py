# file: app_ultralite_sheet1_numbers.py
# Rekonsiliasi Naik/Turun Golongan — Ultra-Lite (.xlsb, Sheet 1), preview 10, filter selisih
# Export Excel: kolom Nominal & Selisih bertipe NUMBER dengan format #,##0.00
# Requirements:
#   streamlit>=1.26
#   pyxlsb>=1.0.10

import io
import re
import tempfile
from typing import Dict, List, Optional
import streamlit as st

# ---------- Setup ----------
st.set_page_config(page_title="Rekonsiliasi (.xlsb) — Ultra-Lite (Sheet 1)", layout="wide")

def _try_set_max_upload(mb: int) -> None:
    try: st.set_option("server.maxUploadSize", int(mb))  # why: kalau ditolak runtime, dibiarkan
    except Exception: pass
_try_set_max_upload(1024)

st.title("🔄 Rekonsiliasi Naik/Turun Golongan — Ultra-Lite (.xlsb, Sheet 1)")
st.caption("Baca Sheet 1, auto-header, preview 10 baris, opsi hanya selisih. Export: kolom nominal & selisih = NUMBER.")

# Metric font kecil agar tidak terpotong
st.markdown("""
<style>
div[data-testid="stMetricLabel"]{font-size:11px!important}
div[data-testid="stMetricValue"]{
  font-size:14px!important; line-height:1.1!important;
  white-space:normal!important; word-break:break-all!important;
  overflow-wrap:anywhere!important; text-overflow:clip!important;
}
</style>
""", unsafe_allow_html=True)

# ---------- Utils ----------
def pyxlsb_ok() -> bool:
    try:
        import pyxlsb  # noqa
        return True
    except Exception:
        return False

def norm(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return (s.replace("no ", "nomor ").replace("no", "nomor")
             .replace("inv", "invoice").replace("harga total", "harga")
             .replace("nilai", "harga").replace("tarip", "tarif"))

def numeric_like(v: str) -> bool:
    return bool(re.fullmatch(r"\s*[\d\.\,\-\(\)]+(\s*)", str(v or "")))

def parse_money(x: str) -> float:
    if x is None: return 0.0
    s = str(x).strip()
    if not s: return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") and s.count("."):
        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
        else: s = s.replace(",", "")
    else:
        if "," in s and "." not in s:
            s = s.replace(",", ".") if re.search(r",[0-9]{1,2}$", s) else s.replace(",", "")
        elif "." in s and "," not in s:
            if not re.search(r"\.[0-9]{1,2}$", s): s = s.replace(".", "")
    try: return float(s)
    except Exception:
        s2 = re.sub(r"[^\d\-]", "", s)
        try: return float(s2)
        except Exception: return 0.0

def format_idr(n: float) -> str:
    s = f"{float(n):,.2f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")

def coerce_key(x: str) -> str:
    return re.sub(r"\s+", "", (x or "")).upper()

def find_exact_idx_multi(headers: List[str], exact_names: List[str]) -> Optional[int]:
    targets = {norm(x) for x in exact_names}
    for i, h in enumerate(headers):
        if norm(h) in targets: return i
    return None

def pick_col_idx(headers: List[str], candidates: List[str]) -> Optional[int]:
    if not headers: return None
    H = {i: norm(h) for i, h in enumerate(headers)}
    for c in [norm(c) for c in candidates]:
        for i, h in H.items():
            if h == c or c in h: return i
    return None

# ---------- XLSX writer (NUMBER support) ----------
def _letters(idx: int) -> str:
    s=""; idx+=1
    while idx: idx,r=divmod(idx-1,26); s=chr(65+r)+s
    return s
def _xml(t:str)->str:
    return (t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
             .replace('"',"&quot;").replace("'","&apos;"))

def build_xlsx(columns: List[str], rows: List[Dict[str, object]], sheet_name="Rekonsiliasi",
               numeric_cols: Optional[List[str]] = None) -> bytes:
    """
    numeric_cols: nama kolom yang harus dipaksa NUMBER + style #,##0.00
    """
    from zipfile import ZipFile
    numeric_cols = set(numeric_cols or [])

    # worksheet xml
    ws=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>']
    # header row
    ws.append('<row r="1">'+"".join(
        f'<c r="{_letters(i)}1" t="inlineStr"><is><t xml:space="preserve">{_xml(c)}</t></is></c>'
        for i,c in enumerate(columns))+ "</row>")
    # data rows
    r=1
    for row in rows:
        r+=1
        cells=[]
        for i,c in enumerate(columns):
            v = row.get(c, "")
            if c in numeric_cols and isinstance(v, (int, float)):  # why: agar NUMBER di Excel
                cells.append(f'<c r="{_letters(i)}{r}" s="1"><v>{v}</v></c>')
            else:
                txt = _xml(str(v if v is not None else ""))
                cells.append(f'<c r="{_letters(i)}{r}" t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>')
        ws.append(f'<row r="{r}">'+"".join(cells)+ "</row>")
    ws.append("</sheetData></worksheet>")

    # build package
    with io.BytesIO() as bio:
        with ZipFile(bio,"w") as z:
            z.writestr("[Content_Types].xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>''')
            z.writestr("_rels/.rels", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')
            z.writestr("xl/workbook.xml", f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="{_xml(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''.encode("utf-8"))
            z.writestr("xl/_rels/workbook.xml.rels", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''')
            z.writestr("xl/worksheets/sheet1.xml", "\n".join(ws).encode("utf-8"))
            # styles: s=0 default, s=1 numeric custom (#,##0.00); numFmtId >= 164 untuk custom
            z.writestr("xl/styles.xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="165" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="1"><font/></fonts>
  <fills count="1"><fill/></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2">
    <xf xfId="0"/>
    <xf xfId="0" numFmtId="165" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>''')
        return bio.getvalue()

# ---------- XLSB reader (Sheet 1, auto-header) ----------
def iter_records_xlsb_sheet1(file):
    """Read hanya Sheet 1; header = baris teks pertama (angka didiskualifikasi)."""
    import pyxlsb  # wajib
    file.seek(0); b = file.read()
    with tempfile.NamedTemporaryFile(delete=True, suffix=".xlsb") as tmp:
        tmp.write(b); tmp.flush()
        with pyxlsb.open_workbook(tmp.name) as wb:
            with wb.get_sheet(1) as sheet:
                header = None
                for row in sheet.rows():
                    vals = [str((c.v if c is not None else "") or "") for c in row]
                    if not any(v.strip() for v in vals):
                        continue
                    if header is None:
                        nonempty = [v for v in vals if v.strip()]
                        nums = sum(1 for v in nonempty if numeric_like(v))
                        header = nonempty if nums <= len(nonempty)//2 else vals
                        header = [h.strip() or f"COL{j+1}" for j,h in enumerate(header)]
                        continue
                    vals = vals + [""] * (len(header) - len(vals))
                    yield {header[i]: vals[i] for i in range(len(header))}

# ---------- Sidebar ----------
with st.sidebar:
    st.header("1) Upload (.xlsb) — multiple")
    inv_files = st.file_uploader("📄 Invoice (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    ts_files  = st.file_uploader("🎫 T-Summary (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    only_diff = st.checkbox("Hanya tampilkan Selisih ≠ 0 (berlaku untuk preview & export)", value=False)

if not pyxlsb_ok():
    st.error("`pyxlsb` belum terpasang. Tambahkan ke `requirements.txt`:\n\n```\nstreamlit>=1.26\npyxlsb>=1.0.10\n```")
    st.stop()

if not inv_files or not ts_files:
    st.info("Unggah minimal satu file **Invoice** dan satu file **T-Summary** (.xlsb).")
    st.stop()

# ---------- Proses ----------
if st.button("🚀 Proses (Sheet 1)"):
    agg_inv: Dict[str, float] = {}
    agg_ts:  Dict[str, float] = {}
    order: List[str] = []; seen = set()

    inv_first = {"tgl_inv": {}, "pay_inv": {}, "tujuan": {}, "channel": {}, "merchant": {}}
    inv_original: Dict[str, str] = {}
    ts_first  = {"pay_ts": {}, "kode": {}, "tiket": {}, "gol": {}, "asal": {}, "cetak": {}}

    # Kandidat non-kunci
    C_INV_AMT = ["harga","nominal","amount","total","nilai"]
    C_INV_TGL = ["tanggal invoice","tgl invoice","tanggal","tgl"]
    C_INV_PAY = ["pembayaran","tgl pembayaran","tanggal pembayaran"]
    C_INV_TUJ = ["tujuan","destination"]
    C_INV_CH  = ["channel"]
    C_INV_MER = ["merchant","mid"]

    C_TS_KEY  = ["nomer invoice","nomor invoice","no invoice","invoice"]
    C_TS_AMT  = ["tarif","harga","nominal","amount","total","nilai"]
    C_TS_KODE = ["kode booking","kode boking","booking"]
    C_TS_TKT  = ["nomor tiket","no tiket","ticket"]
    C_TS_PAY  = ["pembayaran","tgl pembayaran","tanggal pembayaran"]
    C_TS_GOL  = ["golongan","kelas"]
    C_TS_ASAL = ["asal","keberangkatan"]
    C_TS_CTK  = ["cetak boarding pass","tgl cetak"]

    # Invoice
    for f in inv_files:
        key_idx = amt_idx = tgl_idx = pay_idx = tuj_idx = ch_idx = mer_idx = None
        for row in iter_records_xlsb_sheet1(f):
            if key_idx is None:
                hdr = list(row.keys())
                key_idx = find_exact_idx_multi(hdr, ["Nomer Invoice", "Nomor Invoice"])
                if key_idx is None:
                    st.error(f"Invoice **{f.name}** harus memiliki header **'Nomer Invoice'** (fallback 'Nomor Invoice').")
                    st.stop()
                amt_idx = pick_col_idx(hdr, C_INV_AMT)
                tgl_idx = pick_col_idx(hdr, C_INV_TGL)
                pay_idx = pick_col_idx(hdr, C_INV_PAY)
                tuj_idx = pick_col_idx(hdr, C_INV_TUJ)
                ch_idx  = pick_col_idx(hdr, C_INV_CH)
                mer_idx = pick_col_idx(hdr, C_INV_MER)
                if amt_idx is None:
                    st.error(f"Invoice **{f.name}** tidak memiliki kolom Nominal/Harga yang dikenali.")
                    st.stop()
            vals = list(row.values())
            inv_no_raw = vals[key_idx]
            key = coerce_key(inv_no_raw)
            if not key: continue
            if key not in seen: seen.add(key); order.append(key)
            inv_original.setdefault(key, inv_no_raw)
            agg_inv[key] = agg_inv.get(key, 0.0) + parse_money(vals[amt_idx])
            if tgl_idx is not None and vals[tgl_idx]: inv_first["tgl_inv"].setdefault(key, vals[tgl_idx])
            if pay_idx is not None and vals[pay_idx]: inv_first["pay_inv"].setdefault(key, vals[pay_idx])
            if tuj_idx is not None and vals[tuj_idx]: inv_first["tujuan"].setdefault(key, vals[tuj_idx])
            if ch_idx  is not None and vals[ch_idx ]: inv_first["channel"].setdefault(key, vals[ch_idx])
            if mer_idx is not None and vals[mer_idx]: inv_first["merchant"].setdefault(key, vals[mer_idx])

    # T-Summary
    for f in ts_files:
        key_idx = amt_idx = kode_idx = tkt_idx = pay_idx = gol_idx = asal_idx = ctk_idx = None
        for row in iter_records_xlsb_sheet1(f):
            if key_idx is None:
                hdr = list(row.keys())
                key_idx  = pick_col_idx(hdr, C_TS_KEY)
                amt_idx  = pick_col_idx(hdr, C_TS_AMT)
                kode_idx = pick_col_idx(hdr, C_TS_KODE)
                tkt_idx  = pick_col_idx(hdr, C_TS_TKT)
                pay_idx  = pick_col_idx(hdr, C_TS_PAY)
                gol_idx  = pick_col_idx(hdr, C_TS_GOL)
                asal_idx = pick_col_idx(hdr, C_TS_ASAL)
                ctk_idx  = pick_col_idx(hdr, C_TS_CTK)
                if key_idx is None or amt_idx is None:
                    st.error(f"T-Summary **{f.name}** tidak memiliki kolom kunci/nominal yang dikenali.")
                    st.stop()
            vals = list(row.values())
            key = coerce_key(vals[key_idx])
            if not key: continue
            agg_ts[key] = agg_ts.get(key, 0.0) + parse_money(vals[amt_idx])
            if pay_idx  is not None and vals[pay_idx ]: ts_first["pay_ts"].setdefault(key, vals[pay_idx])
            if kode_idx is not None and vals[kode_idx]: ts_first["kode" ].setdefault(key, vals[kode_idx])
            if tkt_idx  is not None and vals[tkt_idx ]: ts_first["tiket"].setdefault(key, vals[tkt_idx])
            if gol_idx  is not None and vals[gol_idx ]: ts_first["gol"  ].setdefault(key, vals[gol_idx])
            if asal_idx is not None and vals[asal_idx]: ts_first["asal" ].setdefault(key, vals[asal_idx])
            if ctk_idx  is not None and vals[ctk_idx ]: ts_first["cetak"].setdefault(key, vals[ctk_idx])

    # Hasil
    rows_preview: List[Dict[str, str]] = []
    rows_export:  List[Dict[str, object]] = []  # kolom nominal/selisih = float
    total_inv = total_ts = total_diff = 0.0
    naik = turun = sama = 0

    for k in order:
        v_inv = float(agg_inv.get(k, 0.0))
        v_ts  = float(agg_ts.get(k, 0.0))
        diff = v_inv - v_ts
        cat = "Turun" if v_inv > v_ts else ("Naik" if v_inv < v_ts else "Sama")

        # preview (string)
        rows_preview.append({
            "Tanggal Invoice":               inv_first["tgl_inv"].get(k, ""),
            "Nomer Invoice":                 inv_original.get(k, k),
            "Kode Booking":                  ts_first["kode"].get(k, ""),
            "Nomor Tiket":                   ts_first["tiket"].get(k, ""),
            "Nominal Invoice":               format_idr(v_inv),
            "Tanggal Pembayaran Invoice":    inv_first["pay_inv"].get(k, ""),
            "Nominal T-Summary":             format_idr(v_ts),
            "Tanggal Pembayaran T-Summary":  ts_first["pay_ts"].get(k, ""),
            "Golongan":                      ts_first["gol"].get(k, ""),
            "Keberangkatan":                 ts_first["asal"].get(k, ""),
            "Tujuan":                        inv_first["tujuan"].get(k, ""),
            "Tgl Cetak Boarding Pass":       ts_first["cetak"].get(k, ""),
            "Channel":                       inv_first["channel"].get(k, ""),
            "Merchant":                      inv_first["merchant"].get(k, ""),
            "Selisih":                       format_idr(diff),
            "Kategori":                      cat,
        })

        # export (number)
        rows_export.append({
            "Tanggal Invoice":               inv_first["tgl_inv"].get(k, ""),
            "Nomer Invoice":                 inv_original.get(k, k),
            "Kode Booking":                  ts_first["kode"].get(k, ""),
            "Nomor Tiket":                   ts_first["tiket"].get(k, ""),
            "Nominal Invoice":               v_inv,     # NUMBER
            "Tanggal Pembayaran Invoice":    inv_first["pay_inv"].get(k, ""),
            "Nominal T-Summary":             v_ts,      # NUMBER
            "Tanggal Pembayaran T-Summary":  ts_first["pay_ts"].get(k, ""),
            "Golongan":                      ts_first["gol"].get(k, ""),
            "Keberangkatan":                 ts_first["asal"].get(k, ""),
            "Tujuan":                        inv_first["tujuan"].get(k, ""),
            "Tgl Cetak Boarding Pass":       ts_first["cetak"].get(k, ""),
            "Channel":                       inv_first["channel"].get(k, ""),
            "Merchant":                      inv_first["merchant"].get(k, ""),
            "Selisih":                       diff,      # NUMBER
            "Kategori":                      cat,
        })

        total_inv += v_inv; total_ts += v_ts; total_diff += diff
        naik += (cat=="Naik"); turun += (cat=="Turun"); sama += (cat=="Sama")

    # Terapkan filter "Hanya Selisih ≠ 0" ke preview & export
    if only_diff:
        rows_preview = [r for r,e in zip(rows_preview, rows_export) if abs(float(e["Selisih"])) > 0.0]
        rows_export  = [e for e in rows_export if abs(float(e["Selisih"])) > 0.0]

    # Metrics (total keseluruhan)
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Invoice", format_idr(total_inv))
    c2.metric("Total T-Summary", format_idr(total_ts))
    c3.metric("Total Selisih (Inv−T)", format_idr(total_diff))
    c4.metric("Naik / Turun / Sama", f"{int(naik)} / {int(turun)} / {int(sama)}")

    # Preview 10 baris
    st.subheader("👀 Preview 10 baris" + (" — hanya Selisih ≠ 0" if only_diff else ""))
    st.dataframe(rows_preview[:10], use_container_width=True)

    # Download Excel: kolom nominal & selisih = NUMBER + style #,##0.00
    cols = [
        "Tanggal Invoice","Nomer Invoice","Kode Booking","Nomor Tiket",
        "Nominal Invoice","Tanggal Pembayaran Invoice",
        "Nominal T-Summary","Tanggal Pembayaran T-Summary",
        "Golongan","Keberangkatan","Tujuan","Tgl Cetak Boarding Pass",
        "Channel","Merchant","Selisih","Kategori",
    ]
    xlsx = build_xlsx(
        columns=cols,
        rows=rows_export,
        sheet_name="Rekonsiliasi",
        numeric_cols=["Nominal Invoice","Nominal T-Summary","Selisih"]
    )
    st.download_button(
        "⬇️ Download Excel (.xlsx)",
        data=xlsx,
        file_name=("rekonsiliasi_ultralite_selisih.xlsx" if only_diff else "rekonsiliasi_ultralite.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
