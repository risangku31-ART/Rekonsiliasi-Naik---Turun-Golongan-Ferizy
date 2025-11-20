# file: app_ultralite_diff_option.py
# Rekonsiliasi Naik/Turun Golongan — .xlsb only, ultra-lite RAM, preview 10 baris
# Fitur: Checkbox "Hanya tampilkan Selisih ≠ 0" (berlaku untuk preview & export)
# Requirements:
#   streamlit>=1.26
#   pyxlsb>=1.0.10

import io
import re
import tempfile
from typing import Dict, List, Optional
import streamlit as st

# ---------- Setup UI ----------
st.set_page_config(page_title="Rekonsiliasi (.xlsb) — Ultra-Lite + Selisih", layout="wide")
def _try_set_max_upload(mb: int) -> None:
    try: st.set_option("server.maxUploadSize", int(mb))  # bisa ditolak di runtime; diabaikan jika gagal
    except Exception: pass
_try_set_max_upload(1024)

st.title("🔄 Rekonsiliasi Naik/Turun Golongan — Ultra-Lite (.xlsb)")
st.caption("Mode hemat RAM: 1 sheet index, auto-header, preview 10 baris. Tambahan: filter hanya selisih ≠ 0.")

# ---------- Utils (ringkas) ----------
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

# ---------- XLSX writer (pure Python, ringan) ----------
def _letters(idx: int) -> str:
    s=""; idx+=1
    while idx: idx,r=divmod(idx-1,26); s=chr(65+r)+s
    return s
def _xml(t:str)->str:
    return (t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
             .replace('"',"&quot;").replace("'","&apos;"))
def build_xlsx(columns: List[str], rows: List[Dict[str, str]], sheet_name="Rekonsiliasi") -> bytes:
    from zipfile import ZipFile
    ws=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>']
    ws.append('<row r="1">'+"".join(
        f'<c r="{_letters(i)}1" t="inlineStr"><is><t xml:space="preserve">{_xml(c)}</t></is></c>'
        for i,c in enumerate(columns))+ "</row>")
    r=1
    for row in rows:
        r+=1
        ws.append(f'<row r="{r}">'+"".join(
            f'<c r="{_letters(i)}{r}" t="inlineStr"><is><t xml:space="preserve">{_xml(str(row.get(c,'') or ''))}</t></is></c>'
            for i,c in enumerate(columns))+ "</row>")
    ws.append("</sheetData></worksheet>")
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>''')
            z.writestr("xl/worksheets/sheet1.xml", "\n".join(ws).encode("utf-8"))
            z.writestr("xl/styles.xml", b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs>
</styleSheet>''')
        return bio.getvalue()

# ---------- XLSB reader (single-sheet, auto-header) ----------
def iter_records_xlsb(file, sheet_index: int = 1):
    import pyxlsb  # wajib
    file.seek(0); b = file.read()
    with tempfile.NamedTemporaryFile(delete=True, suffix=".xlsb") as tmp:
        tmp.write(b); tmp.flush()
        with pyxlsb.open_workbook(tmp.name) as wb:
            with wb.get_sheet(sheet_index) as sheet:
                header = None
                for row in sheet.rows():
                    vals = [str((c.v if c is not None else "") or "") for c in row]
                    if not any(v.strip() for v in vals):
                        continue
                    if header is None:
                        nonempty = [v for v in vals if v.strip()]
                        nums = sum(1 for v in nonempty if numeric_like(v))
                        header = nonempty if nums <= len(nonempty)//2 else vals  # hindari baris numeric jadi header
                        header = [h.strip() or f"COL{j+1}" for j,h in enumerate(header)]
                        continue
                    vals = vals + [""] * (len(header) - len(vals))
                    yield {header[i]: vals[i] for i in range(len(header))}

# ---------- Sidebar Upload ----------
with st.sidebar:
    st.header("1) Upload (.xlsb) — multiple")
    inv_files = st.file_uploader("📄 Invoice (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    ts_files  = st.file_uploader("🎫 T-Summary (.xlsb)", type=["xlsb"], accept_multiple_files=True)
    sheet_idx = st.number_input("Index sheet (1-based, dipakai untuk keduanya)", min_value=1, value=1, step=1)
    only_diff = st.checkbox("Hanya tampilkan Selisih ≠ 0 (juga saat export)", value=False)

if not pyxlsb_ok():
    st.error("`pyxlsb` belum terpasang. Tambahkan ke `requirements.txt`:\n\n```\nstreamlit>=1.26\npyxlsb>=1.0.10\n```")
    st.stop()

if not inv_files or not ts_files:
    st.info("Unggah minimal satu file **Invoice** dan satu file **T-Summary** (.xlsb).")
    st.stop()

# ---------- Proses ----------
if st.button("🚀 Proses (Ultra-Lite)"):
    # Aggregates
    agg_inv: Dict[str, float] = {}
    agg_ts:  Dict[str, float] = {}
    order: List[str] = []; seen = set()

    inv_first = {"tgl_inv": {}, "pay_inv": {}, "tujuan": {}, "channel": {}, "merchant": {}}
    inv_original: Dict[str, str] = {}  # simpan teks asli 'Nomer Invoice'
    ts_first  = {"pay_ts": {}, "kode": {}, "tiket": {}, "gol": {}, "asal": {}, "cetak": {}}  # only first values

    # Kolom non-kunci (fuzzy)
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
        for row in iter_records_xlsb(f, sheet_index=sheet_idx):
            if key_idx is None:
                hdr = list(row.keys())
                key_idx = find_exact_idx_multi(hdr, ["Nomer Invoice", "Nomor Invoice"])
                if key_idx is None:
                    st.error(f"Invoice **{f.name}** harus punya header **'Nomer Invoice'** (fallback 'Nomor Invoice').")
                    st.stop()
                amt_idx = pick_col_idx(hdr, C_INV_AMT)
                tgl_idx = pick_col_idx(hdr, C_INV_TGL)
                pay_idx = pick_col_idx(hdr, C_INV_PAY)
                tuj_idx = pick_col_idx(hdr, C_INV_TUJ)
                ch_idx  = pick_col_idx(hdr, C_INV_CH)
                mer_idx = pick_col_idx(hdr, C_INV_MER)
                if amt_idx is None:
                    st.error(f"Invoice **{f.name}** tidak punya kolom Nominal/Harga yang dikenali.")
                    st.stop()
            vals = list(row.values())
            inv_no_raw = vals[key_idx]
            key = coerce_key(inv_no_raw)
            if not key: continue
            if key not in seen: seen.add(key); order.append(key)
            inv_original.setdefault(key, inv_no_raw)  # only first
            agg_inv[key] = agg_inv.get(key, 0.0) + parse_money(vals[amt_idx])
            if tgl_idx is not None and vals[tgl_idx]: inv_first["tgl_inv"].setdefault(key, vals[tgl_idx])
            if pay_idx is not None and vals[pay_idx]: inv_first["pay_inv"].setdefault(key, vals[pay_idx])
            if tuj_idx is not None and vals[tuj_idx]: inv_first["tujuan"].setdefault(key, vals[tuj_idx])
            if ch_idx  is not None and vals[ch_idx ]: inv_first["channel"].setdefault(key, vals[ch_idx])
            if mer_idx is not None and vals[mer_idx]: inv_first["merchant"].setdefault(key, vals[mer_idx])

    # T-Summary
    for f in ts_files:
        key_idx = amt_idx = kode_idx = tkt_idx = pay_idx = gol_idx = asal_idx = ctk_idx = None
        for row in iter_records_xlsb(f, sheet_index=sheet_idx):
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
                    st.error(f"T-Summary **{f.name}** tidak punya kolom kunci/nominal yang dikenali.")
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

    # Hasil (simpan selisih numerik untuk filter)
    rows_raw: List[Dict[str,str]] = []
    total_inv = total_ts = total_diff = 0.0
    naik = turun = sama = 0

    for k in order:
        v_inv = float(agg_inv.get(k, 0.0))
        v_ts  = float(agg_ts.get(k, 0.0))
        diff = v_inv - v_ts
        cat = "Turun" if v_inv > v_ts else ("Naik" if v_inv < v_ts else "Sama")
        rows_raw.append({
            "__diff__":                      diff,  # internal numeric for filtering
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
        total_inv += v_inv; total_ts += v_ts; total_diff += diff
        naik += (cat=="Naik"); turun += (cat=="Turun"); sama += (cat=="Sama")

    # Filter selisih jika diminta
    def _strip_internal(rows: List[Dict[str,str]]) -> List[Dict[str,str]]:
        out = []
        for r in rows:
            if "__diff__" in r:
                r = dict(r)  # copy ringan
                r.pop("__diff__", None)
            out.append(r)
        return out

    rows_to_use = rows_raw
    if only_diff:
        rows_to_use = [r for r in rows_raw if abs(float(r.get("__diff__", 0.0))) > 0.0]

    rows = _strip_internal(rows_to_use)

    # Metrics (total keseluruhan, bukan filtered)
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Total Invoice", format_idr(total_inv))
    c2.metric("Total T-Summary", format_idr(total_ts))
    c3.metric("Total Selisih (Inv−T)", format_idr(total_diff))
    c4.metric("Naik / Turun / Sama", f"{int(naik)} / {int(turun)} / {int(sama)}")

    # Preview 10 baris (fixed)
    st.subheader("👀 Preview 10 baris" + (" — hanya Selisih ≠ 0" if only_diff else ""))
    st.dataframe(rows[:10], use_container_width=True)

    # Download Excel (terapkan filter yang sama)
    cols = [
        "Tanggal Invoice","Nomer Invoice","Kode Booking","Nomor Tiket",
        "Nominal Invoice","Tanggal Pembayaran Invoice",
        "Nominal T-Summary","Tanggal Pembayaran T-Summary",
        "Golongan","Keberangkatan","Tujuan","Tgl Cetak Boarding Pass",
        "Channel","Merchant","Selisih","Kategori",
    ]
    xlsx = build_xlsx(cols, rows, sheet_name="Rekonsiliasi")
    st.download_button("⬇️ Download Excel (.xlsx)", data=xlsx,
                       file_name=("rekonsiliasi_ultralite_selisih.xlsx" if only_diff else "rekonsiliasi_ultralite.xlsx"),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
