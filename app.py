import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
import os
import requests
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import pytesseract
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import base64
import sys
from pathlib import Path


def resource_path(rel_path: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
    return str(base / rel_path)


st.set_page_config(page_title="Dashboard Monitoring", page_icon="📡", layout="wide")

# TESSERACT
pytesseract.pytesseract.tesseract_cmd = resource_path("tesseract_bin/tesseract.exe")
os.environ["TESSDATA_PREFIX"] = resource_path("tesseract_bin/tessdata")


# FUNGSI GAMBAR
def img_to_base64(rel_path: str) -> str:
    p = resource_path(rel_path)
    if not os.path.exists(p):
        return ""
    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()


def asset_exists(rel_path: str) -> bool:
    return os.path.exists(resource_path(rel_path))


# LOAD ICON GLOBALLY
loading_icon_b64 = img_to_base64("assets/loading.png")
add_icon_b64 = img_to_base64("assets/add.png")

# CSS GLOBAL
st.markdown(
    f"""
    <style>
        #MainMenu {{ 
            visibility: hidden; 
        }}

        header {{ 
            visibility: hidden; 
        }}

        footer {{ 
            visibility: hidden; 
        }}

        .red-strip {{ 
            position: fixed; 
            top: 0; 
            left: 0; 
            width: 15px; 
            height: 100vh; 
            background-color: #EE2D24; 
            z-index: 9999; 
        }}
        
        .block-container {{ 
            padding-top: 1rem !important; 
            padding-left: 3.5rem !important; 
        }}

        .stProgress > div > div > div > div {{ 
            background-color: #EE2D24; 
        }}

        [data-testid="stBorderWrapper"] {{ 
            border: 2px solid #000 !important; 
            border-radius: 15px !important; 
            padding: 30px !important; 
            background-color: white; 
        }}

        .main-title {{ 
            font-size: 28px; 
            font-weight: 800; 
            color: #1a1a1a; 
            margin-bottom: 5px; 
        }}

        .sub-title {{ 
            font-size: 14px; 
            color: #555; 
            margin-bottom: 30px; 
        }}

        .upload-note {{ 
            font-size: 12px; 
            color: #d32f2f; 
            font-style: italic; 
            margin-top: -10px; 
            margin-bottom: 20px; 
        }}
        
        .card-success {{ 
            background-color: #ECFDF3; 
            border-left: 6px solid #16A34A; 
            border-radius: 14px; 
            padding: 16px 18px; 
            margin-bottom: 14px; 
        }}

        .card-failed {{ 
            background-color: #FEF2F2; 
            border-left: 6px solid #DC2626; 
            border-radius: 14px; 
            padding: 16px 18px; 
            margin-bottom: 14px; 
        }}

        .card-multiple {{ 
            background-color: #FFFBEB; 
            border-left: 6px solid #F59E0B; 
            border-radius: 14px; 
            padding: 16px 18px; 
            margin-bottom: 14px; 
        }}

        .card-title {{ 
            font-weight: 600; 
            font-size: 14px; 
            margin-bottom: 4px; 
        }}

        .card-meta {{ 
            font-size: 12px; 
            color: #555; 
        }}
        
        .metric-box {{ 
            background-color: white; 
            border: 1px solid #e5e7eb; 
            border-radius: 12px; 
            padding: 15px 10px; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.03); 
            text-align: center; 
            display: flex; 
            flex-direction: column; 
            align-items: center; 
            justify-content: center; 
            height: 100%; 
        }}

        .metric-title {{ 
            font-size: 11px; 
            font-weight: 700; 
            color: #6b7280; 
            text-transform: uppercase; 
            letter-spacing: 0.8px; 
            margin-bottom: 8px; 
        }}

        .metric-value {{ 
            font-size: 32px; 
            font-weight: 800; 
            color: #111827; 
            line-height: 1; 
        }}

        .mb-blue {{ 
            border-bottom: 4px solid #3b82f6; 
        }}

        .mb-green {{ 
            border-bottom: 4px solid #22c55e; 
        }}

        .mb-yellow {{ 
            border-bottom: 4px solid #eab308; 
        }}

        .mb-red {{ 
            border-bottom: 4px solid #ef4444; 
        }}
        
        div[role="radiogroup"] {{ 
            justify-content: center !important; 
        }}
        
        [data-testid="stFileUploader"] button {{ 
            background-color: #EE2D24 !important; 
            border: 2px solid #EE2D24 !important; 
            color: white !important; 
            border-radius: 50px !important; 
        }}

        [data-testid="stFileUploader"] section {{ 
            background-color: #f8f9fa !important; 
            border-radius: 15px !important; 
            border: 1px dashed #ccc !important; 
        }}
        
        /* CSS TOAST ANIMATION */
        @keyframes slideIn {{ 
            from {{ 
                transform: translateX(100%); 
                opacity: 0; 
            }} 
            to {{ 
                transform: translateX(0); 
                opacity: 1; 
            }} 
        }}

        @keyframes spin {{ 
            0% {{ 
                transform: rotate(0deg); 
            }} 
            100% {{ 
                transform: rotate(360deg); 
            }} 
        }}
        
        .custom-toast {{
            position: fixed; 
            top: 20px; 
            right: 20px;
            background-color: white; 
            border-left: 5px solid #EE2D24;
            padding: 15px 20px; 
            border-radius: 8px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.15);
            display: flex; 
            align-items: center; 
            gap: 15px;
            z-index: 999999; 
            animation: slideIn 0.3s ease-out forwards;
            min-width: 320px;
        }}

        .toast-icon {{
            width: 30px; 
            height: 30px;
            background-image: url('data:image/png;base64,{loading_icon_b64}'); /* INI SEKARANG AMAN */
            background-size: contain; 
            background-repeat: no-repeat; 
            background-position: center;
            animation: spin 1s linear infinite;
        }}

        .toast-content {{ 
            font-family: 'Segoe UI', sans-serif; 
            flex-grow: 1; 
        }}

        .toast-title {{ 
            font-weight: 800; 
            font-size: 14px; 
            color: #333; 
            margin-bottom: 2px; 
        }}

        .toast-desc {{ 
            font-size: 13px; 
            color: #666; 
            font-weight: 500; 
        }}

        .toast-progress {{ 
            font-size: 11px; 
            color: #999; 
            margin-top: 2px; 
        }}
    </style>

    <div class="red-strip"></div>
    """,
    unsafe_allow_html=True,
)


# BACKEND
# SCRAPPING DATA DARI WEB GRAPH
def scrape_dual_server(sid, month_idx, year):
    servers = ["http://10.62.8.136/cacti", "http://10.62.8.135/cacti"]
    last_day = calendar.monthrange(year, month_idx)[1]
    start_d = f"{year}-{month_idx:02d}-01"
    end_d = f"{year}-{month_idx:02d}-{last_day:02d}"

    combined_graphs = []
    for base in servers:
        try:
            params = {
                "action": "preview",
                "filter": sid,
                "date1": start_d,
                "date2": end_d,
            }
            resp = requests.get(f"{base}/graph_view.php", params=params, timeout=8)
            graph_ids = re.findall(r"local_graph_id=(\d+)", resp.text)
            for g_id in list(set(graph_ids)):
                combined_graphs.append(
                    {
                        "url": f"{base}/graph_image.php?action=view&local_graph_id={g_id}&rra_id=3",
                        "server": "136" if "136" in base else "135",
                    }
                )
        except:
            continue
    return combined_graphs


# HITUNG KE KBPS
def convert_to_kbps(value_str, unit):
    try:
        s = str(value_str).strip()
        if s == "" or s.lower() in ["nan", "-nan", "none"]:
            return 0.0

        s = s.replace(",", ".")
        val = float(s)


        u = (unit or "").strip()


        if u == "k" or u == "K":
            return val

        elif u == "M":
            return val * 1024

        elif u == "m":
            return val / 1000

        elif u == "G" or u == "g":
            return val * 1_000_000

        elif u == "":
            return val / 1000

        else:
            return val / 1_000_000

    except Exception as e:
        return 0.0


# OCR
def ocr_extract_data(image_path):
    try:
        img = Image.open(image_path)

        w, h = img.size
        img = img.crop((0, int(h * 0.75), w, h))

        new_w = img.width * 3
        new_h = img.height * 3
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        img = img.convert("L")
        img = img.point(lambda x: 0 if x < 170 else 255, "1")

        custom_config = r"--oem 3 --psm 6"
        text = pytesseract.image_to_string(img, config=custom_config)

        pattern_in = r"Inbound.*?Average:\s*([\d\.,]+)\s*(\S*)\s*Max"
        pattern_out = r"Outbound.*?Average:\s*([\d\.,]+)\s*(\S*)\s*Max"

        inbound_match = re.search(pattern_in, text, re.IGNORECASE | re.DOTALL)
        outbound_match = re.search(pattern_out, text, re.IGNORECASE | re.DOTALL)

        def clean_num(s):
            s = str(s).strip()
            s = s.replace(",", ".")
            s = re.sub(r"\.+", ".", s)
            s = s.strip(".")
            return s

        if inbound_match:
            raw_val = clean_num(inbound_match.group(1))
            raw_unit = inbound_match.group(2)
            in_kbps = convert_to_kbps(raw_val, raw_unit)
        else:
            in_kbps = 0.0

        if outbound_match:
            raw_val = clean_num(outbound_match.group(1))
            raw_unit = outbound_match.group(2)
            out_kbps = convert_to_kbps(raw_val, raw_unit)
        else:
            out_kbps = 0.0

        return in_kbps, out_kbps

    except Exception as e:
        return 0.0, 0.0


# GENERATE WORD
def generate_clean_word(data, month, year):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = section.left_margin = (
        section.right_margin
    ) = Inches(1)
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), "720")

    total_link = len(data)
    if asset_exists("assets/telkom.jpg"):
        logo_p = doc.add_paragraph()
        logo_p.add_run().add_picture(
            resource_path("assets/telkom.jpg"), width=Inches(1)
        )
        logo_p.alignment = 0

    title_p1 = doc.add_paragraph()
    r1 = title_p1.add_run("TRAFIK MRTG BANK JATIM")
    r1.font.bold = True
    r1.font.size = Pt(16)

    title_p2 = doc.add_paragraph()
    r2 = title_p2.add_run(f"{total_link} LINK TELKOM {month.upper()} {year}")
    r2.font.size = Pt(14)
    doc.add_paragraph("")

    for i, item in enumerate(data, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_together = True

        header = f"{i}. {item['alamat']} SID.{item['sid']} BW.{item['bw']}"
        r = p.add_run(header)
        r.font.size = Pt(7)
        r.font.bold = True

        if item.get("selected_url"):
            try:
                img = requests.get(item["selected_url"], timeout=12).content
                tmp_img = f"tmp_{i}.png"
                with open(tmp_img, "wb") as f:
                    f.write(img)
                img_p = doc.add_paragraph()
                img_p.paragraph_format.keep_together = True
                img_p.add_run().add_picture(tmp_img, width=Inches(2.3))
                os.remove(tmp_img)
            except:
                err = doc.add_paragraph()
                rr = err.add_run("[Tidak Ditemukan Grafik]")
                rr.font.size = Pt(6)
                rr.font.color.rgb = RGBColor(255, 0, 0)
                rr.font.italic = True
        else:
            err = doc.add_paragraph()
            rr = err.add_run("[Tidak Ditemukan Grafik]")
            rr.font.size = Pt(6)
            rr.font.color.rgb = RGBColor(255, 0, 0)
            rr.font.italic = True

        doc.add_paragraph("")

    out_name = f"Laporan_MRTG_{month}_{year}.docx"
    doc.save(out_name)
    return out_name


# GENERATE EXCEL
def generate_excel_report(data, month, year):
    indo_months = [
        "Januari",
        "Februari",
        "Maret",
        "April",
        "Mei",
        "Juni",
        "Juli",
        "Agustus",
        "September",
        "Oktober",
        "November",
        "Desember",
    ]
    month_idx = indo_months.index(month) + 1
    days = calendar.monthrange(year, month_idx)[1]

    rows = []
    for i, item in enumerate(data, start=1):
        in_v = float(item.get("in_kbps", 0) or 0)
        out_v = float(item.get("out_kbps", 0) or 0)
        total_avg = in_v + out_v
        est_kb = total_avg * 24 * 3600 * days

        status = (
            "Tidak Ada Grafik"
            if not item.get("selected_url")
            else "Data Nol / OCR Gagal Membaca Grafik" if total_avg == 0 else "Sukses"
        )

        rows.append(
            {
                "No": i,
                "Alamat Cabang": item["alamat"],
                "SID": item["sid"],
                "Bandwidth": item["bw"],
                "Avg Inbound (Kbps)": in_v,
                "Avg Outbound (Kbps)": out_v,
                "Total Average (Kbps)": total_avg,
                "Total Trafik (Kb)": est_kb,
                "Status": status,
            }
        )

    df = pd.DataFrame(rows)
    fname = f"Laporan_Trafik_{month}_{year}.xlsx"
    df.to_excel(fname, index=False)

    wb = load_workbook(fname)
    ws = wb.active

    yellow = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
    red = PatternFill(start_color="EA9999", end_color="EA9999", fill_type="solid")

    status_col = df.columns.get_loc("Status") + 1

    for r in range(2, ws.max_row + 1):
        val = ws.cell(r, status_col).value
        if val == "Tidak Ada Grafik":
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = yellow
        elif val == "Data Nol / OCR Gagal Membaca Grafik":
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = red

    wb.save(fname)
    return fname


# DOWNLOAD BUTTON
def make_download_button(file_path, label, css_class, mime, icon_rel_path):
    with open(file_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    with open(resource_path(icon_rel_path), "rb") as ic:
        icon_b64 = base64.b64encode(ic.read()).decode()

    filename = os.path.basename(file_path)

    return f"""
    <a class="dl-btn {css_class}"
       href="data:{mime};base64,{b64}"
       download="{filename}">
        <img src="data:image/png;base64,{icon_b64}" class="dl-icon"/>
        <span>{label}</span>
    </a>
    """


# RESET BUTTON
def make_reset_button(label, icon_rel_path):
    try:
        with open(resource_path(icon_rel_path), "rb") as ic:
            icon_b64 = base64.b64encode(ic.read()).decode()
        img_tag = f'<img src="data:image/png;base64,{icon_b64}" class="reset-icon"/>'
    except:
        img_tag = ""

    return f"""
    <a class="reset-btn" href="/?reset=1" target="_self">
        {img_tag}
        <span>{label}</span>
    </a>
    """


# ICON
def img_to_base64(rel_path: str) -> str:
    p = resource_path(rel_path)
    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()


def asset_exists(rel_path: str) -> bool:
    return os.path.exists(resource_path(rel_path))


ARROW_ICON = img_to_base64("assets/arrow.png")

# MAIN PROCESS
if asset_exists("assets/telkom.jpg"):
    st.image(resource_path("assets/telkom.jpg"), width=200)
else:
    st.markdown("### Telkom Indonesia")

params = st.query_params
if params.get("reset") == ["1"]:
    st.session_state.clear()
    st.query_params.clear()
    st.session_state.step = "input"
    st.rerun()

if "step" not in st.session_state:
    st.session_state.step = "input"

if st.session_state.step == "input":
    try:
        loading_icon_b64 = img_to_base64("assets/loading.png")
    except:
        loading_icon_b64 = ""

    st.markdown(
        f"""
        <style>
            button[kind="primary"] {{
                background-color: #EE2D24 !important; 
                border: 2px solid #EE2D24 !important; 
                color: white !important;
                border-radius: 50px !important; 
                padding: 0.5rem 2rem !important; 
                font-weight: bold !important;
                transition: all 0.2s ease-in-out !important; 
                box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
            }}

            button[kind="primary"]:hover {{
                background-color: #cc0000 !important; 
                border-color: #cc0000 !important; 
                transform: scale(1.02) !important;
            }}

            button[kind="secondary"] {{
                background-color: white !important; 
                border: 2px solid #EE2D24 !important; 
                color: #EE2D24 !important;
                border-radius: 50px !important; 
                padding: 0.5rem 2rem !important; 
                font-weight: bold !important;
                transition: all 0.2s ease-in-out !important; 
                box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
            }}

            button[kind="secondary"]:hover {{
                background-color: #cc0000 !important; 
                border-color: #cc0000 !important; 
                color: white !important; 
                transform: scale(1.02) !important;
            }}

            [data-testid="stFileUploader"] button {{
                display: inline-flex !important; 
                justify-content: center !important; 
                align-items: center !important;
                margin-top: 5px !important; 
                padding: 8px 25px !important; 
                font-weight: normal !important;
                background-color: #EE2D24 !important; 
                border: 2px solid #EE2D24 !important; 
                color: white !important;
                border-radius: 50px !important; 
                transition: all 0.2s ease-in-out !important;
            }}

            [data-testid="stFileUploader"] section {{ 
                background-color: #f8f9fa !important; 
                border-radius: 15px !important; 
                border: 1px dashed #ccc !important; 
            }}

            @keyframes slideIn {{ 
                from {{ 
                    transform: translateX(100%); 
                    opacity: 0; 
                }} 
                to {{ 
                    transform: translateX(0); 
                    opacity: 1; 
                }} 
            }}

            @keyframes spin {{ 
                0% {{ 
                    transform: rotate(0deg); 
                }} 
                100% {{ 
                    transform: rotate(360deg); 
                }} 
            }}
            
            .custom-toast {{
                position: fixed; 
                top: 20px; 
                right: 20px;
                background-color: white; 
                border-left: 5px solid #EE2D24;
                padding: 15px 20px; 
                border-radius: 8px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
                display: flex; 
                align-items: center; 
                gap: 15px;
                z-index: 999999; 
                animation: slideIn 0.3s ease-out forwards;
                min-width: 320px;
            }}

            .toast-icon {{
                width: 30px; 
                height: 30px;
                background-image: url('data:image/png;base64,{loading_icon_b64}');
                background-size: contain; 
                background-repeat: no-repeat; 
                background-position: center;
                animation: spin 1s linear infinite;
            }}

            .toast-content {{ 
                font-family: 'Segoe UI', sans-serif; 
                flex-grow: 1; 
            }}

            .toast-title {{ 
                font-weight: 800; 
                font-size: 14px; 
                color: #333; 
                margin-bottom: 2px; 
            }}

            .toast-desc {{ 
                font-size: 13px; 
                color: #666; 
                font-weight: 500; 
            }}

            .toast-progress {{ 
                font-size: 11px; 
                color: #999; 
                margin-top: 2px; 
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="main-title">Dashboard Monitoring & Otomatisasi Laporan</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="sub-title">Automasi Monitoring, Validasi Grafik, dan Generasi Laporan</div>',
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.write("**Upload File Excel**")
        file = st.file_uploader(
            "", type=["xlsx"], label_visibility="collapsed", key="file_uploader"
        )
        st.markdown(
            '<div class="upload-note">⚠️ Excel harus memiliki kolom Alamat, SID, dan Bandwidth</div>',
            unsafe_allow_html=True,
        )

        st.write("**Periode**")
        c1, c2 = st.columns(2)
        month_list = [
            "Januari",
            "Februari",
            "Maret",
            "April",
            "Mei",
            "Juni",
            "Juli",
            "Agustus",
            "September",
            "Oktober",
            "November",
            "Desember",
        ]
        m = c1.selectbox(
            "Bulan",
            options=month_list,
            index=datetime.now().month - 1,
            label_visibility="collapsed",
        )
        y = c2.selectbox(
            "Tahun",
            options=list(range(2022, 2031)),
            index=datetime.now().year - 2022,
            label_visibility="collapsed",
        )

        st.markdown("<br>", unsafe_allow_html=True)
        btn_proc, btn_cancel, _ = st.columns([1.5, 1.5, 8])

        with btn_proc:
            is_process = st.button("Process", type="primary", use_container_width=True)
        with btn_cancel:
            is_cancel = st.button("Cancel", type="secondary", use_container_width=True)

        if is_cancel:
            st.session_state.clear()
            st.rerun()

        if is_process:
            if file:
                progress_toast = st.empty()

                progress_toast.markdown(
                    f"""
                    <div class="custom-toast">
                        <div class="toast-icon"></div>
                        <div class="toast-content">
                            <div class="toast-title">Persiapan Data...</div>
                            <div class="toast-desc">Sedang membaca file Excel</div>
                        </div>
                    </div>
                """,
                    unsafe_allow_html=True,
                )

                df = pd.read_excel(file)
                df.columns = [str(c).strip().upper() for c in df.columns]

                if not all(col in df.columns for col in ["ALAMAT", "SID", "BANDWIDTH"]):
                    st.error(
                        "Kolom tidak sesuai! Pastikan ada kolom: Alamat, SID, Bandwidth"
                    )
                    st.stop()

                m_idx = month_list.index(m) + 1
                final_res = []
                total_data = len(df)

                for idx, row in df.iterrows():
                    current_num = idx + 1
                    percent = int((current_num / total_data) * 100)
                    sid_now = str(row["SID"])

                    progress_toast.markdown(
                        f"""
                        <div class="custom-toast">
                            <div class="toast-icon"></div>
                            <div class="toast-content">
                                <div class="toast-title">Sedang Memproses.. ({percent}%)</div>
                                <div class="toast-desc">Membaca data dan mengecek grafik MRTG</div>
                                <div class="toast-progress">Data ke-{current_num} dari {total_data}</div>
                            </div>
                        </div>
                    """,
                        unsafe_allow_html=True,
                    )

                    graphs = scrape_dual_server(sid_now, m_idx, y)

                    final_res.append(
                        {
                            "alamat": row["ALAMAT"],
                            "sid": row["SID"],
                            "bw": row["BANDWIDTH"],
                            "graphs": graphs,
                            "selected_url": (
                                graphs[0]["url"] if len(graphs) == 1 else None
                            ),
                        }
                    )

                st.session_state.update(
                    {
                        "results": final_res,
                        "month": m,
                        "year": y,
                        "step": "validate",
                        "current_page": 1,
                        "q_validate": "",
                        "f_validate": "Semua",
                    }
                )
                st.rerun()
            else:
                st.error("Silakan upload file Excel terlebih dahulu.")

elif st.session_state.step == "validate":
    try:
        add_icon_b64 = img_to_base64("assets/add.png")
    except:
        add_icon_b64 = ""

    try:
        loading_icon_b64 = img_to_base64("assets/loading.png")
    except:
        loading_icon_b64 = ""

    st.markdown(
        f"""
        <style>
            button[kind="primary"] {{
                background-color: #EE2D24 !important;
                border: 2px solid #EE2D24 !important;
                color: white !important;
                border-radius: 50px !important;
                padding: 0.6rem 1rem !important;
                font-weight: 800 !important;
                box-shadow: 0 4px 6px rgba(238, 45, 36, 0.2) !important;
                transition: all 0.2s ease-in-out !important;
            }}

            button[kind="primary"]:hover {{
                background-color: #cc0000 !important;
                border-color: #cc0000 !important;
                transform: scale(1.02) !important;
            }}

            button[kind="primary"]::before {{
                content: "";
                display: inline-block;
                width: 22px;
                height: 22px;
                background-image: url('data:image/png;base64,{add_icon_b64}');
                background-size: contain;
                background-repeat: no-repeat;
                background-position: center;
                filter: brightness(0) invert(1);
                flex-shrink: 0;
                margin-right: 8px;
            }}

            button[kind="secondary"] {{
                background-color: white !important;
                border: 2px solid #000 !important;
                color: #000 !important;
                border-radius: 12px !important;
                padding: 8px 16px !important;
                font-weight: 600 !important;
                transition: all 0.2s ease !important;
                height: auto !important;
            }}

            button[kind="secondary"]:hover {{
                border-color: #EE2D24 !important;
                color: #EE2D24 !important;
                background-color: #fff5f5 !important;
            }}
            
            button[kind="secondary"]:disabled {{
                border-color: #e5e7eb !important;    
                background-color: #f9fafb !important; 
                color: #9ca3af !important;            
                opacity: 0.7 !important;              
                cursor: not-allowed !important;
            }}

            div[data-testid="column"]:nth-of-type(2) button {{
                border: 2px solid #EE2D24 !important; 
                color: #EE2D24 !important;            
                background-color: white !important;
                border-radius: 10px !important;       
                padding: 0px !important;
                height: 46px !important;
                width: 100% !important;
                display: flex !important; 
                align-items: center !important; 
                justify-content: center !important;
            }}

            div[data-testid="column"]:nth-of-type(2) button:hover {{
                background-color: #EE2D24 !important; 
                color: white !important;              
                border-color: #EE2D24 !important;
                box-shadow: 0 4px 8px rgba(238, 45, 36, 0.3) !important;
                transform: translateY(-2px);
            }}

            div[data-testid="column"]:nth-of-type(2) button p {{
                font-size: 28px !important;
                font-weight: 900 !important;
                line-height: 1 !important;
                margin-top: -5px !important;
            }}

            @keyframes slideIn {{
                from {{
                    transform: translateX(100%);
                    opacity: 0;
                }}
                to {{
                    transform: translateX(0);
                    opacity: 1;
                }}
            }}

            @keyframes spin {{
                0% {{
                    transform: rotate(0deg);
                }}
                100% {{
                    transform: rotate(360deg);
                }}
            }}
            
            .custom-toast {{
                position: fixed;
                top: 20px;
                right: 20px;
                background-color: white;
                border-left: 5px solid #EE2D24;
                padding: 15px 20px;
                border-radius: 8px;
                box-shadow: 0 4px 15px rgba(0,0,0,0.15);
                display: flex;
                align-items: center;
                gap: 15px;
                z-index: 999999;
                animation: slideIn 0.5s ease-out forwards;
                min-width: 300px;
            }}

            .toast-icon {{
                width: 30px;
                height: 30px;
                background-image: url('data:image/png;base64,{loading_icon_b64}');
                background-size: contain;
                background-repeat: no-repeat;
                background-position: center;
                animation: spin 1s linear infinite; 
            }}

            .toast-content {{
                font-family: 'Segoe UI', sans-serif;
            }}

            .toast-title {{
                font-weight: 800;
                font-size: 14px;
                color: #333;
                margin-bottom: 2px;
            }}

            .toast-desc {{
                font-size: 12px;
                color: #666;
            }}

            .pag-center {{
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100%;
                font-size: 16px;
                color: #444;
                padding-top: 10px;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="main-title">Validasi Grafik</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-title">Jika ditemukan lebih dari satu grafik, silakan pilih yang paling akurat.</div>',
        unsafe_allow_html=True,
    )

    if "current_page" not in st.session_state:
        st.session_state.current_page = 1
    if "q_validate" not in st.session_state:
        st.session_state.q_validate = ""
    if "f_validate" not in st.session_state:
        st.session_state.f_validate = "Semua"

    def reset_callback():
        st.session_state.q_validate = ""
        st.session_state.f_validate = "Semua"
        st.session_state.current_page = 1

    c_search, c_reset, c_filter, c_btn = st.columns([5, 1, 3, 3], gap="small")

    with c_search:
        st.text_input(
            "Cari Cabang",
            placeholder="Cari Nama Alamat atau SID",
            label_visibility="collapsed",
            key="q_validate",
        )

    with c_reset:
        st.button(
            "↺",
            key="btn_reset_text",
            help="Reset Filter & Pencarian",
            on_click=reset_callback,
            use_container_width=True,
            type="secondary",
        )

    with c_filter:
        st.selectbox(
            "Filter Status",
            options=["Semua", "Berhasil", "Perlu Validasi", "Tidak Ada Grafik"],
            label_visibility="collapsed",
            key="f_validate",
        )

    with c_btn:
        do_print = st.button(
            "CETAK SEMUA LAPORAN", type="primary", use_container_width=True
        )

    st.markdown("<div style='margin-bottom: 20px;'></div>", unsafe_allow_html=True)

    total_data = len(st.session_state.results)
    sukses = sum(1 for x in st.session_state.results if len(x.get("graphs", [])) == 1)
    perlu_cek = sum(1 for x in st.session_state.results if len(x.get("graphs", [])) > 1)
    gagal = sum(1 for x in st.session_state.results if len(x.get("graphs", [])) == 0)

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(
            f'<div class="metric-box mb-blue"><div class="metric-title">TOTAL CABANG</div><div class="metric-value">{total_data}</div></div>',
            unsafe_allow_html=True,
        )
    with m2:
        st.markdown(
            f'<div class="metric-box mb-green"><div class="metric-title">✅ BERHASIL</div><div class="metric-value">{sukses}</div></div>',
            unsafe_allow_html=True,
        )
    with m3:
        st.markdown(
            f'<div class="metric-box mb-yellow"><div class="metric-title">⚠️ PERLU VALIDASI</div><div class="metric-value">{perlu_cek}</div></div>',
            unsafe_allow_html=True,
        )
    with m4:
        st.markdown(
            f'<div class="metric-box mb-red"><div class="metric-title">❌ TIDAK ADA GRAFIK</div><div class="metric-value">{gagal}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    def compute_status(cnt):
        if cnt == 0:
            return "Tidak Ada Grafik"
        if cnt == 1:
            return "Berhasil"
        return "Perlu Validasi"

    query = (st.session_state.q_validate or "").strip()
    f_stat = st.session_state.f_validate

    processed = []
    for i, item in enumerate(st.session_state.results):
        cnt = len(item.get("graphs", []))
        status = compute_status(cnt)
        q_ok = (
            (query == "")
            or (query.lower() in item["alamat"].lower())
            or (query in str(item["sid"]))
        )
        f_ok = (f_stat == "Semua") or (f_stat == status)
        if q_ok and f_ok:
            it = dict(item)
            it["orig_idx"] = i
            it["status"] = status
            processed.append(it)

    items_per_page = 5
    total_pages = max(1, (len(processed) + items_per_page - 1) // items_per_page)
    st.session_state.current_page = max(
        1, min(st.session_state.current_page, total_pages)
    )
    start_idx = (st.session_state.current_page - 1) * items_per_page
    paged_data = processed[start_idx : start_idx + items_per_page]

    for item in paged_data:
        idx = item["orig_idx"]
        status = item["status"]
        card_class = (
            "card-success"
            if status == "Berhasil"
            else "card-failed" if status == "Tidak Ada Grafik" else "card-multiple"
        )
        badge = (
            "✅ Berhasil"
            if status == "Berhasil"
            else (
                "❌ Tidak Ada Grafik"
                if status == "Tidak Ada Grafik"
                else "⚠️ Perlu Validasi"
            )
        )

        st.markdown(
            f"""
        <div class="{card_class}">
            <div class="card-title">{idx+1}. {item['alamat']}</div>
            <div class="card-meta">SID: {item['sid']} &nbsp;|&nbsp; BW: {item['bw']} &nbsp;|&nbsp; <b>{badge}</b></div>
        </div>
        """,
            unsafe_allow_html=True,
        )

        with st.container():
            if status == "Tidak Ada Grafik":
                st.error("Tidak Ditemukan Grafik.")
            elif status == "Berhasil":
                st.image(item["graphs"][0]["url"], width=520)
                st.session_state.results[idx]["selected_url"] = item["graphs"][0]["url"]
            else:
                graphs = item["graphs"]
                if (
                    not st.session_state.results[idx].get("selected_url")
                    and len(graphs) > 0
                ):
                    st.session_state.results[idx]["selected_url"] = graphs[0]["url"]
                current_selected = st.session_state.results[idx].get("selected_url")

                for row_start in range(0, len(graphs), 3):
                    row_graphs = graphs[row_start : row_start + 3]
                    cols = st.columns(3)
                    for col_i, g in enumerate(row_graphs):
                        with cols[col_i]:
                            st.image(g["url"], use_container_width=True)
                            is_selected = current_selected == g["url"]
                            label = (
                                "☑ Terpilih"
                                if is_selected
                                else f"Opsi {row_start + col_i + 1}"
                            )

                            if st.button(
                                label,
                                key=f"pick_{idx}_{row_start + col_i}",
                                type="secondary",
                                use_container_width=True,
                                disabled=is_selected,
                            ):
                                st.session_state.results[idx]["selected_url"] = g["url"]
                                st.rerun()
                    st.markdown(
                        "<div style='height:10px'></div>", unsafe_allow_html=True
                    )
        st.divider()

    if do_print:
        st.markdown(
            f"""
            <div class="custom-toast">
                <div class="toast-icon"></div>
                <div class="toast-content">
                    <div class="toast-title">Sedang Memproses..</div>
                    <div class="toast-desc">Membaca Grafik & Membuat Laporan</div>
                </div>
            </div>
        """,
            unsafe_allow_html=True,
        )

        for i, it in enumerate(st.session_state.results):
            if it.get("selected_url"):
                try:
                    img_data = requests.get(it["selected_url"], timeout=10).content
                    tmp_path = f"tmp_ocr_{i}.png"
                    with open(tmp_path, "wb") as f:
                        f.write(img_data)
                    in_kbps, out_kbps = ocr_extract_data(tmp_path)
                    os.remove(tmp_path)
                    st.session_state.results[i]["in_kbps"] = in_kbps
                    st.session_state.results[i]["out_kbps"] = out_kbps
                except:
                    st.session_state.results[i]["in_kbps"] = 0.0
                    st.session_state.results[i]["out_kbps"] = 0.0
            else:
                st.session_state.results[i]["in_kbps"] = 0.0
                st.session_state.results[i]["out_kbps"] = 0.0

        path_word = generate_clean_word(
            st.session_state.results, st.session_state.month, st.session_state.year
        )
        path_excel = generate_excel_report(
            st.session_state.results, st.session_state.month, st.session_state.year
        )

        st.session_state.update(
            {"final_path": path_word, "final_excel": path_excel, "step": "finish"}
        )
        st.rerun()

    c_prev, c_mid, c_next = st.columns([3, 4, 3])
    with c_prev:
        if st.button(
            "« Previous",
            type="secondary",
            use_container_width=True,
            disabled=(st.session_state.current_page == 1),
            key="pg_prev",
        ):
            st.session_state.current_page -= 1
            st.rerun()
    with c_mid:
        st.markdown(
            f"<div class='pag-center'>Halaman&nbsp;<b>{st.session_state.current_page}</b>&nbsp;dari&nbsp;<b>{total_pages}</b></div>",
            unsafe_allow_html=True,
        )
    with c_next:
        if st.button(
            "Next »",
            type="secondary",
            use_container_width=True,
            disabled=(st.session_state.current_page == total_pages),
            key="pg_next",
        ):
            st.session_state.current_page += 1
            st.rerun()

elif st.session_state.step == "finish":
    st.markdown(
        """
        <style>
            .finish-title {
                font-size: 36px;
                font-weight: 800;
                color: #000;
                text-align: center;
                margin-bottom: 20px;
                font-family: 'Segoe UI', sans-serif;
            }

            .dl-btn, .reset-btn {
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 12px;
                height: 60px; 
                border-radius: 999px;
                font-weight: 800;
                font-size: 18px;
                text-decoration: none !important;
                border: 3px solid #000;
                box-shadow: 0 4px 10px rgba(0,0,0,0.1);
                transition: all 0.2s ease;
                width: 100%;
                cursor: pointer;
            }

            .dl-btn:hover, .reset-btn:hover {
                transform: translateY(-2px);
                box-shadow: 0 6px 15px rgba(0,0,0,0.15);
            }

            .btn-blue {
                background: #2D62D3;
                color: #fff !important;
                border-color: #2D62D3 !important;
            }

            .btn-green {
                background: #1E7A3A;
                color: #fff !important;
                border-color: #1E7A3A !important;
            }

            .reset-btn {
                background: #ffffff;
                color: #000 !important;
                border-color: #000 !important;
            }

            .dl-icon, .reset-icon {
                width: 24px;
                height: 24px;
                object-fit: contain;
            }

            .download-row {
                margin-top: 40px;
            }

            .reset-row {
                margin-top: 20px;
            }

            @keyframes bounce-check {
                0%, 100% {
                    transform: translateY(0);
                }
                40% {
                    transform: translateY(-15px);
                }
                60% {
                    transform: translateY(-7px);
                }
            }

            .check-bounce {
                animation: bounce-check 1.8s ease-in-out infinite;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div style="display:flex;justify-content:center;margin-bottom:20px;margin-top:20px;">
            <div class="check-bounce" style="
                width:100px;height:100px;background:#EE2D24;border-radius:50%;
                display:flex;align-items:center;justify-content:center;
                box-shadow: 0 10px 25px rgba(238, 45, 36, 0.3);">
                <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="20 6 9 17 4 12"></polyline>
                </svg>
            </div>
        </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="finish-title">Laporan Berhasil Dibuat!</div>',
        unsafe_allow_html=True,
    )

    icon_dl = (
        "assets/downloads.png"
        if asset_exists("assets/downloads.png")
        else "assets/arrow.png"
    )

    html_word = make_download_button(
        st.session_state.final_path,
        "Download Word",
        "btn-blue",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        icon_dl,
    )
    html_excel = make_download_button(
        st.session_state.final_excel,
        "Download Excel",
        "btn-green",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon_dl,
    )

    st.markdown('<div class="download-row">', unsafe_allow_html=True)
    cL, cWord, cExcel, cR = st.columns([2, 3, 3, 2], gap="large")
    with cWord:
        st.markdown(html_word, unsafe_allow_html=True)
    with cExcel:
        st.markdown(html_excel, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="reset-row">', unsafe_allow_html=True)
    rL, rMid, rR = st.columns([3, 4, 3])
    with rMid:
        icon_reset = (
            "assets/reset.png"
            if asset_exists("assets/reset.png")
            else "assets/arrow.png"
        )
        html_reset = make_reset_button("Buat Laporan Baru", icon_reset)
        st.markdown(html_reset, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
