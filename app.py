import os
import io
import re
import csv
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set page configuration with a premium look
st.set_page_config(
    page_title="KeyPDF Highlight & Analytics Studio",
    page_icon="🖊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Curated High-Contrast 20 Color Palette for Multi-color Highlighting
COLOR_PALETTE = [
    ((1.0, 1.0, 0.0), "#FFFF00", "เหลืองนีออน (Neon Yellow)"),
    ((0.2, 0.9, 0.4), "#33E666", "เขียวมินต์ (Mint Green)"),
    ((0.0, 0.8, 1.0), "#00CCFF", "ฟ้าสว่าง (Bright Cyan)"),
    ((1.0, 0.4, 0.7), "#FF66B2", "ชมพูพาสเทล (Hot Pink)"),
    ((1.0, 0.6, 0.0), "#FF9900", "ส้มแมนดาริน (Mandarin Orange)"),
    ((0.75, 0.52, 0.99), "#C084FC", "ม่วงลาเวนเดอร์ (Lavender)"),
    ((1.0, 0.42, 0.42), "#FF6B6B", "แดงปะการัง (Coral Red)"),
    ((0.98, 0.8, 0.08), "#FACC15", "ทองคำ (Rich Gold)"),
    ((0.65, 0.93, 0.15), "#A3E635", "เขียวมะนาว (Lime Green)"),
    ((0.18, 0.83, 0.78), "#2DD4BF", "ฟ้าเทอร์ควอยซ์ (Turquoise)"),
    ((0.53, 0.43, 0.96), "#818CF8", "ม่วงเมฆา (Violet Sky)"),
    ((0.98, 0.45, 0.63), "#FB7185", "ชมพูกุหลาบ (Rose Pink)"),
    ((0.99, 0.58, 0.43), "#FB923C", "ส้มพีช (Peach Coral)"),
    ((0.06, 0.73, 0.51), "#10B981", "เขียวเอมเมอรัลด์ (Emerald)"),
    ((0.23, 0.66, 0.96), "#38BDF8", "ฟ้าคราม (Sky Blue)"),
    ((0.91, 0.3, 0.88), "#E879F9", "ม่วงพลาสม่า (Magenta)"),
    ((0.96, 0.72, 0.07), "#F59E0B", "เหลืองแอมเบอร์ (Amber)"),
    ((0.96, 0.26, 0.45), "#F43F5E", "แดงราสเบอร์รี่ (Raspberry)"),
    ((0.3, 0.85, 0.39), "#4ADE80", "เขียวแอปเปิ้ล (Spring Green)"),
    ((0.02, 0.71, 0.83), "#06B6D4", "ฟ้าครามเข้ม (Deep Cyan)")
]

# Premium Theme and Custom Styling
st.markdown("""
    <style>
    /* Main app layout */
    .stApp {
        background-color: #0F172A;
        color: #E2E8F0;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background-color: #1E293B !important;
        border-right: 1px solid #334155;
    }
    
    /* Headers styling */
    h1, h2, h3 {
        color: #F8FAFC !important;
        font-weight: 700 !important;
    }
    
    /* Title gradient */
    .title-gradient {
        background: linear-gradient(135deg, #38BDF8 0%, #34D399 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.5rem;
        text-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Card containers */
    .glass-card {
        background: rgba(30, 41, 59, 0.7);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
    }
    
    /* Highlight summary styles */
    .metric-card {
        background: rgba(15, 23, 42, 0.6);
        border-left: 4px solid #34D399;
        border-radius: 6px;
        padding: 12px 20px;
        margin: 8px 0;
    }
    
    /* Highlight labels */
    .highlight-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85em;
        background-color: rgba(52, 211, 153, 0.15);
        color: #34D399;
        border: 1px solid rgba(52, 211, 153, 0.3);
    }
    
    /* Styled buttons */
    div.stButton > button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
        color: white !important;
        border: none !important;
        padding: 10px 24px !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.3) !important;
    }
    div.stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 16px rgba(5, 150, 105, 0.4) !important;
    }
    
    /* Styled file uploader borders */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #475569 !important;
        padding: 16px !important;
        border-radius: 10px !important;
        background-color: #1E293B !important;
    }
    
    /* Footer styling */
    .footer {
        text-align: center;
        margin-top: 50px;
        font-size: 0.85rem;
        color: #64748B;
        border-top: 1px solid #334155;
        padding-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)


def load_keywords_from_bytes(file_bytes, filename):
    """
    Parses keywords from uploaded file bytes (TXT or CSV) using various encodings.
    """
    keywords = []
    content = ""
    
    encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'tis-620', 'iso-8859-11']
    for enc in encodings:
        try:
            content = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        content = file_bytes.decode('utf-8', errors='ignore')

    if filename.lower().endswith('.csv'):
        reader = csv.reader(io.StringIO(content))
        for row in reader:
            for cell in row:
                word = cell.strip()
                if word and word not in keywords:
                    keywords.append(word)
    else:
        for line in content.splitlines():
            word = line.strip()
            if word and word not in keywords:
                keywords.append(word)
                
    return keywords


def get_hex_color(color_name):
    """
    Returns float RGB values (0.0 to 1.0) and hex code for PyMuPDF highlighting.
    """
    for rgb, hex_val, name in COLOR_PALETTE:
        if color_name == name:
            return rgb, hex_val
    return COLOR_PALETTE[0][0], COLOR_PALETTE[0][1]


def process_pdf(pdf_bytes, keywords, search_flags, single_color_rgb, case_sensitive, match_whole_word=True, color_mode="อัตโนมัติ (แยกสีตามคีย์เวิร์ด)", progress_callback=None):
    """
    Performs native search and highlighting using PyMuPDF.
    Supports exact Keyword matching, 20 Multi-color palette per keyword, Co-occurrence analysis, and progress tracking.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Color assignment mapping
    kw_color_map = {}
    for idx, kw in enumerate(keywords):
        if color_mode == "อัตโนมัติ (แยกสีตามคีย์เวิร์ด)":
            rgb, hex_val, name = COLOR_PALETTE[idx % len(COLOR_PALETTE)]
            kw_color_map[kw] = {"rgb": rgb, "hex": hex_val, "name": name}
        else:
            kw_color_map[kw] = {"rgb": single_color_rgb, "hex": "#FFFF00", "name": "สีเดี่ยว"}
            
    summary_data = []
    report_data = []
    total_matches = 0
    total_pdf_words = 0
    keyword_match_counts = {kw: 0 for kw in keywords}
    
    page_kw_map = {p: set() for p in range(1, len(doc) + 1)}
    total_pages = len(doc)
    
    for page_num in range(total_pages):
        if progress_callback:
            progress_callback(page_num + 1, total_pages)
            
        page = doc[page_num]
        page_1based = page_num + 1
        
        textpage = page.get_textpage(flags=search_flags)
        words = page.get_text("words")
        total_pdf_words += len(words)
        
        # 1. Map words to lines and build line text context
        line_coords = {}
        line_words = {}
        for w in words:
            key = (w[5], w[6])  # (block_no, line_no)
            y0 = w[1]
            if key not in line_coords:
                line_coords[key] = []
                line_words[key] = []
            line_coords[key].append(y0)
            line_words[key].append(w)
        
        line_avg_y = {key: sum(y0s)/len(y0s) for key, y0s in line_coords.items()}
        sorted_lines = sorted(line_avg_y.items(), key=lambda item: item[1])
        line_map = {key: idx + 1 for idx, (key, _) in enumerate(sorted_lines)}
        
        line_text_map = {}
        for key, ws in line_words.items():
            ws_sorted = sorted(ws, key=lambda x: x[7])
            line_text_map[key] = " ".join(x[4] for x in ws_sorted)
            
        # 2. Native Keyword Search
        for kw in keywords:
            kw_clean = kw.strip()
            if not kw_clean:
                continue
                
            rects = page.search_for(kw_clean, textpage=textpage)
            
            if rects:
                if case_sensitive:
                    rects = [r for r in rects if page.get_text("text", clip=r).strip() == kw_clean]
                
                # Filter whole word matches if requested (prevents AI inside Airport / thailand)
                if match_whole_word and rects:
                    valid_rects = []
                    is_english_kw = bool(re.match(r'^[a-zA-Z0-9_\s]+$', kw_clean))
                    
                    for rect in rects:
                        overlapping = [w for w in words if not (rect & fitz.Rect(w[:4])).is_empty]
                        if overlapping:
                            word_strs = [w[4] for w in overlapping]
                            full_word_context = " ".join(word_strs)
                            
                            if is_english_kw:
                                pattern = r'(?<![a-zA-Z0-9_])' + re.escape(kw_clean) + r'(?![a-zA-Z0-9_])'
                            else:
                                pattern = re.escape(kw_clean)
                                
                            flags = 0 if case_sensitive else re.IGNORECASE
                            if re.search(pattern, full_word_context, flags):
                                valid_rects.append(rect)
                        else:
                            valid_rects.append(rect)
                    rects = valid_rects
                
                if rects:
                    count_on_page = len(rects)
                    total_matches += count_on_page
                    keyword_match_counts[kw] += count_on_page
                    page_kw_map[page_1based].add(kw)
                    
                    highlight_rgb = kw_color_map[kw]["rgb"]
                    
                    for rect in rects:
                        annot = page.add_highlight_annot(rect)
                        if annot:
                            annot.set_colors(stroke=highlight_rgb)
                            annot.update()
                            
                        matched_block_line = None
                        for w in words:
                            w_rect = fitz.Rect(w[:4])
                            intersect = rect & w_rect
                            if not intersect.is_empty and intersect.get_area() / w_rect.get_area() > 0.5:
                                matched_block_line = (w[5], w[6])
                                break
                        
                        global_line_no = "-"
                        context_text = "-"
                        if matched_block_line:
                            global_line_no = line_map.get(matched_block_line, "-")
                            context_text = line_text_map.get(matched_block_line, "-")
                            
                        report_data.append({
                            "Keyword": kw,
                            "Page": page_1based,
                            "Line": global_line_no,
                            "Context": context_text
                        })
                        
    # Calculate Co-occurrence Matrix (Shared Pages)
    co_occurrence = {}
    for kw1 in keywords:
        co_occurrence[kw1] = {}
        for kw2 in keywords:
            shared_pages = sum(1 for p, kws in page_kw_map.items() if kw1 in kws and kw2 in kws)
            co_occurrence[kw1][kw2] = shared_pages

    # Build Summary List
    for kw, count in keyword_match_counts.items():
        pct_of_pdf = (count / total_pdf_words * 100) if total_pdf_words > 0 else 0.0
        pct_of_matches = (count / total_matches * 100) if total_matches > 0 else 0.0
        
        summary_data.append({
            "คีย์เวิร์ด (Keyword)": kw,
            "สีไฮไลต์ (Color)": kw_color_map[kw]["hex"],
            "ชื่อเฉดสี (Color Name)": kw_color_map[kw]["name"],
            "จำนวนคำที่พบ (Matches)": count,
            "% เทียบคำทั้งหมดใน PDF (% of PDF Words)": f"{pct_of_pdf:.4f}%",
            "% เทียบคำสำคัญที่พบทั้งหมด (% of Matches)": f"{pct_of_matches:.2f}%"
        })
        
    output_buffer = io.BytesIO()
    doc.save(output_buffer, garbage=4, deflate=True)
    highlighted_pdf_bytes = output_buffer.getvalue()
    
    preview_images = []
    preview_limit = min(5, len(doc))
    for page_num in range(preview_limit):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        preview_images.append((page_num + 1, Image.open(io.BytesIO(img_data))))
        
    doc.close()
    
    return highlighted_pdf_bytes, summary_data, report_data, co_occurrence, total_matches, total_pdf_words, preview_images, kw_color_map


def generate_excel_bytes(summary, report_data, co_occurrence, total_pdf_words, total_matches):
    """
    Generates a rich formatted Excel workbook (.xlsx) with 3 worksheets.
    """
    wb = openpyxl.Workbook()
    
    title_font = Font(name="Calibri", size=14, bold=True, color="0F172A")
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    # Sheet 1: Summary & Density
    ws1 = wb.active
    ws1.title = "Summary & Density"
    
    ws1.append(["PDF KEYWORD ANALYTICS SUMMARY"])
    ws1["A1"].font = title_font
    ws1.append([])
    
    ws1.append(["Metric Overview", "Value"])
    for cell in ws1[3]:
        cell.fill = header_fill
        cell.font = header_font
        
    overall_density = (total_matches / total_pdf_words * 100) if total_pdf_words > 0 else 0.0
    ws1.append(["Total PDF Words (จำนวนคำทั้งหมดใน PDF)", total_pdf_words])
    ws1.append(["Total Keyword Matches Found (คำสำคัญที่พบทั้งหมด)", total_matches])
    ws1.append(["Overall Keyword Density (% สัดส่วนคำสำคัญเทียบทั้ง PDF)", f"{overall_density:.4f}%"])
    ws1.append([])
    
    ws1.append(["Keyword", "Color Name", "Color Hex", "Matches Found", "% of Total PDF Words (แบบที่ 1)", "% of Total Keyword Matches (แบบที่ 2)"])
    header_row = ws1.max_row
    for cell in ws1[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        
    for item in summary:
        ws1.append([
            item["คีย์เวิร์ด (Keyword)"],
            item["ชื่อเฉดสี (Color Name)"],
            item["สีไฮไลต์ (Color)"],
            item["จำนวนคำที่พบ (Matches)"],
            item["% เทียบคำทั้งหมดใน PDF (% of PDF Words)"],
            item["% เทียบคำสำคัญที่พบทั้งหมด (% of Matches)"]
        ])
        
    # Sheet 2: Detailed Findings
    ws2 = wb.create_sheet(title="Detailed Findings")
    ws2.append(["DETAILED FINDINGS & POSITIONS"])
    ws2["A1"].font = title_font
    ws2.append([])
    
    ws2.append(["Keyword", "Page Number", "Line Number (Top-to-Bottom)", "Context Line Text"])
    for cell in ws2[3]:
        cell.fill = header_fill
        cell.font = header_font
        
    for row in report_data:
        ws2.append([row["Keyword"], row["Page"], row["Line"], row["Context"]])
        
    # Sheet 3: Co-occurrence Matrix
    ws3 = wb.create_sheet(title="Co-occurrence Analysis")
    ws3.append(["KEYWORD CO-OCCURRENCE MATRIX (SHARED PAGES)"])
    ws3["A1"].font = title_font
    ws3.append([])
    
    if co_occurrence:
        kws = list(co_occurrence.keys())
        ws3.append(["Keyword"] + kws)
        for cell in ws3[3]:
            cell.fill = header_fill
            cell.font = header_font
            
        for kw1 in kws:
            row_vals = [kw1] + [co_occurrence[kw1][kw2] for kw2 in kws]
            ws3.append(row_vals)
            
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 80)
            
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def generate_csv_bytes(summary, report_data, total_pdf_words, total_matches):
    output = io.StringIO()
    writer = csv.writer(output)
    
    overall_density = (total_matches / total_pdf_words * 100) if total_pdf_words > 0 else 0.0
    
    writer.writerow(["--- PDF KEYWORD SEARCH SUMMARY ---"])
    writer.writerow(["Total PDF Words", total_pdf_words])
    writer.writerow(["Total Keyword Matches", total_matches])
    writer.writerow(["Overall Density", f"{overall_density:.4f}%"])
    writer.writerow([])
    
    writer.writerow(["--- KEYWORD BREAKDOWN ---"])
    writer.writerow(["Keyword", "Color Name", "Color Hex", "Matches Count", "% of Total PDF Words", "% of Total Keyword Matches"])
    for item in summary:
        writer.writerow([
            item["คีย์เวิร์ด (Keyword)"],
            item["ชื่อเฉดสี (Color Name)"],
            item["สีไฮไลต์ (Color)"],
            item["จำนวนคำที่พบ (Matches)"],
            item["% เทียบคำทั้งหมดใน PDF (% of PDF Words)"],
            item["% เทียบคำสำคัญที่พบทั้งหมด (% of Matches)"]
        ])
        
    writer.writerow([])
    
    writer.writerow(["--- DETAILED FINDINGS ---"])
    writer.writerow(["Keyword", "Page Number", "Line Number", "Context Line Text"])
    for row in report_data:
        writer.writerow([row["Keyword"], row["Page"], row["Line"], row["Context"]])
        
    return output.getvalue().encode('utf-8-sig')


# --- STREAMLIT UI ---

def main():
    st.markdown('<div class="title-gradient">KeyPDF Highlight & Analytics Studio</div>', unsafe_allow_html=True)
    st.markdown("<p style='color: #94A3B8; font-size: 1.1rem; margin-top:-10px;'>ยกระดับการตรวจทานเอกสาร PDF ค้นหาตรงตามคำเต็ม แยกสีไฮไลต์อัจฉริยะ 20 เฉดสี พร้อมวิเคราะห์สถิติส่งออก Excel</p>", unsafe_allow_html=True)
    st.write("---")
    
    st.sidebar.markdown("### 🛠️ อัปโหลดและตั้งค่า (Upload & Settings)")
    
    # 1. Upload PDF
    uploaded_pdf = st.sidebar.file_uploader(
        "1. เลือกไฟล์ PDF ที่ต้องการไฮไลต์", 
        type=["pdf"],
        help="อัปโหลดไฟล์ PDF ที่มี Text Layer อยู่แล้ว"
    )
    
    # 2. Upload/Input Keywords
    st.sidebar.write("")
    st.sidebar.markdown("**2. ระบุคีย์เวิร์ดคำสำคัญ**")
    kw_source = st.sidebar.radio(
        "วิธีการใส่คีย์เวิร์ด:",
        ["อัปโหลดไฟล์ (.txt, .csv)", "พิมพ์คีย์เวิร์ดด้วยตนเอง"],
        label_visibility="collapsed"
    )
    
    keywords = []
    if kw_source == "อัปโหลดไฟล์ (.txt, .csv)":
        uploaded_kw = st.sidebar.file_uploader(
            "อัปโหลดไฟล์คีย์เวิร์ด (.txt / .csv)", 
            type=["txt", "csv"],
            help="ไฟล์ข้อความธรรมดา (หนึ่งคำต่อบรรทัด) หรือไฟล์ CSV"
        )
        if uploaded_kw:
            keywords = load_keywords_from_bytes(uploaded_kw.read(), uploaded_kw.name)
    else:
        text_kw = st.sidebar.text_area(
            "พิมพ์คำสำคัญที่ต้องการ (หนึ่งคำต่อหนึ่งบรรทัด):",
            placeholder="ตัวอย่างเช่น:\nAI\nPython\ncontract",
            height=130
        )
        if text_kw:
            keywords = [line.strip() for line in text_kw.splitlines() if line.strip()]

    if keywords:
        st.sidebar.markdown(f"<span class='highlight-badge'>โหลดคีย์เวิร์ดสำเร็จ: {len(keywords)} คำ</span>", unsafe_allow_html=True)
    
    st.sidebar.write("")
    st.sidebar.markdown("### 🎨 ตั้งค่าการค้นหาและสีไฮไลต์ (20 เฉดสี)")
    
    # Match Whole Word option
    match_whole_word = st.sidebar.checkbox(
        "ตรงตามคำเต็มเท่านั้น (Match Whole Word Only)",
        value=True,
        help="หากเปิดใช้งาน คำว่า 'AI' จะไม่นับคำที่มี ai อยู่ภายใน เช่น 'Airport' หรือ 'thailand'"
    )
    
    # Color Modes
    color_mode = st.sidebar.radio(
        "โหมดสีของไฮไลต์ (Highlight Color Mode):",
        ["อัตโนมัติ (แยกสีตามคีย์เวิร์ด)", "สีเดี่ยว (Single Color)"],
        index=0
    )
    
    single_color_rgb = (1.0, 1.0, 0.0)
    single_color_choice = COLOR_PALETTE[0][2]
    if color_mode == "สีเดี่ยว (Single Color)":
        single_color_choice = st.sidebar.selectbox(
            "เลือกสีไฮไลต์เดี่ยว (20 เฉดสี):",
            [name for _, _, name in COLOR_PALETTE],
            index=0
        )
        single_color_rgb, _ = get_hex_color(single_color_choice)
    
    case_sensitive = st.sidebar.checkbox(
        "Case Sensitive (ตัวพิมพ์เล็ก-ใหญ่)", 
        value=False
    )
    
    dehyphenate = st.sidebar.checkbox(
        "Dehyphenate (รวมคำตัดยติภังค์ท้ายบรรทัด)", 
        value=True
    )
    
    search_flags = 2 | 1
    if dehyphenate:
        search_flags |= 16
    
    # Main panel
    col1, col2 = st.columns([1.1, 1.1])
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("📋 รายละเอียดการประมวลผล")
        
        if not uploaded_pdf:
            st.info("👋 กรุณาอัปโหลดไฟล์ PDF ในแผงควบคุมด้านซ้ายเพื่อเริ่มต้น")
        elif not keywords:
            st.warning("⚠️ กรุณาระบุหรืออัปโหลดคีย์เวิร์ดคำสำคัญ")
        else:
            pdf_name = uploaded_pdf.name
            st.write(f"📂 **ไฟล์ PDF:** `{pdf_name}`")
            st.write(f"🔑 **จำนวนคีย์เวิร์ด:** `{len(keywords)} คำ`")
            st.write(f"🎯 **โหมดการค้นหา:** `{'ตรงตามคำเต็มเท่านั้น (Match Whole Word)' if match_whole_word else 'รวมคำส่วนย่อย (Substring Match)'}`")
            st.write(f"🎨 **โหมดสี:** `{color_mode}`")
            
            run_btn = st.button("🚀 เริ่มไฮไลต์และวิเคราะห์เอกสาร")
            
            if run_btn:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                def update_progress(curr, total):
                    progress_bar.progress(curr / total)
                    status_text.text(f"⏳ กำลังประมวลผลและไฮไลต์หน้า {curr}/{total}...")
                    
                try:
                    pdf_bytes = uploaded_pdf.read()
                    
                    res_bytes, summary, report_data, co_occurrence, count, total_pdf_words, previews, kw_color_map = process_pdf(
                        pdf_bytes, 
                        keywords, 
                        search_flags, 
                        single_color_rgb,
                        case_sensitive,
                        match_whole_word,
                        color_mode,
                        progress_callback=update_progress
                    )
                    
                    excel_report_bytes = generate_excel_bytes(summary, report_data, co_occurrence, total_pdf_words, count)
                    csv_report_bytes = generate_csv_bytes(summary, report_data, total_pdf_words, count)
                    
                    st.session_state["result_pdf"] = res_bytes
                    st.session_state["result_excel"] = excel_report_bytes
                    st.session_state["result_csv"] = csv_report_bytes
                    st.session_state["summary"] = summary
                    st.session_state["report_data"] = report_data
                    st.session_state["co_occurrence"] = co_occurrence
                    st.session_state["total_matches"] = count
                    st.session_state["total_pdf_words"] = total_pdf_words
                    st.session_state["previews"] = previews
                    st.session_state["kw_color_map"] = kw_color_map
                    st.session_state["processed_name"] = f"{os.path.splitext(pdf_name)[0]}_highlighted.pdf"
                    st.session_state["excel_name"] = f"{os.path.splitext(pdf_name)[0]}_analytics_report.xlsx"
                    st.session_state["csv_name"] = f"{os.path.splitext(pdf_name)[0]}_keyword_report.csv"
                    
                    progress_bar.empty()
                    status_text.empty()
                    st.success(f"✨ ประมวลผลเสร็จสิ้น! พบคำสำคัญทั้งหมด {count} จุด จากคำใน PDF {total_pdf_words:,} คำ")
                    
                except Exception as e:
                    progress_bar.empty()
                    status_text.empty()
                    st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล PDF: {str(e)}")
                    
        # Display Results
        if "result_pdf" in st.session_state:
            st.write("---")
            st.markdown("#### 🎉 สรุปผลลัพธ์และดาวน์โหลดเอกสาร")
            
            total_pdf_words = st.session_state['total_pdf_words']
            total_matches = st.session_state['total_matches']
            overall_density = (total_matches / total_pdf_words * 100) if total_pdf_words > 0 else 0.0
            
            if total_pdf_words == 0:
                st.warning("⚠️ **ไม่พบข้อความ (Text Layer) ในไฟล์ PDF นี้** — เอกสารอาจเป็นภาพสแกน (Scanned PDF) หรือไม่รองรับการคัดลอกข้อความ แนะนำให้ผ่านระบบ OCR แปลงเป็น Searchable PDF ก่อนนำมาค้นหา")
                
            m_col1, m_col2, m_col3 = st.columns(3)
            with m_col1:
                st.metric("จำนวนคำทั้งหมดใน PDF", f"{total_pdf_words:,} คำ")
            with m_col2:
                st.metric("คำสำคัญที่พบทั้งหมด", f"{total_matches:,} จุด")
            with m_col3:
                st.metric("สัดส่วนคำสำคัญ (% PDF)", f"{overall_density:.4f}%")
                
            st.write("")
            
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                st.download_button(
                    label="📥 ดาวน์โหลด PDF ไฮไลต์",
                    data=st.session_state["result_pdf"],
                    file_name=st.session_state["processed_name"],
                    mime="application/pdf",
                    use_container_width=True
                )
            with d_col2:
                st.download_button(
                    label="📊 ดาวน์โหลดรายงาน Excel (.xlsx)",
                    data=st.session_state["result_excel"],
                    file_name=st.session_state["excel_name"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            st.write("")
            
            tab1, tab2, tab3 = st.tabs([
                "📊 ตารางแจกแจงสถิติ & สีไฮไลต์", 
                "📌 คำสำคัญที่พบร่วมกัน (Co-occurrence)", 
                "🔍 รายละเอียดตำแหน่งที่พบ (Findings)"
            ])
            
            with tab1:
                st.write("**ตารางวิเคราะห์คำสำคัญและสีไฮไลต์ประจำคำ (20 เฉดสี):**")
                df_summary = pd.DataFrame(st.session_state["summary"])[
                    ["คีย์เวิร์ด (Keyword)", "ชื่อเฉดสี (Color Name)", "สีไฮไลต์ (Color)", "จำนวนคำที่พบ (Matches)", "% เทียบคำทั้งหมดใน PDF (% of PDF Words)", "% เทียบคำสำคัญที่พบทั้งหมด (% of Matches)"]
                ]
                
                def make_color_badge(hex_code):
                    return f'<span style="background-color: {hex_code}; color: #000; padding: 3px 10px; border-radius: 4px; font-weight: bold;">■ {hex_code}</span>'
                
                df_summary_display = df_summary.copy()
                df_summary_display["สีไฮไลต์ (Color)"] = df_summary_display["สีไฮไลต์ (Color)"].apply(make_color_badge)
                
                st.write(
                    df_summary_display.to_html(escape=False, index=False),
                    unsafe_allow_html=True
                )
                
            with tab2:
                st.write("**ตารางวิเคราะห์คำสำคัญที่พบอยู่ร่วมกันในหน้าเดียวกัน (Shared Pages Matrix):**")
                co_matrix = st.session_state["co_occurrence"]
                if co_matrix:
                    df_co = pd.DataFrame(co_matrix)
                    st.dataframe(df_co, use_container_width=True)
                else:
                    st.info("ไม่พบข้อมูล Co-occurrence")
                    
            with tab3:
                st.write("**รายละเอียดตำแหน่งหน้า บรรทัด และข้อความบริบท:**")
                df_report = pd.DataFrame(st.session_state["report_data"])
                if not df_report.empty:
                    st.dataframe(df_report, use_container_width=True, hide_index=True)
                else:
                    st.info("ไม่พบคำสำคัญในเอกสาร")
                    
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("👁️ พรีวิวหน้าเอกสาร (PDF Preview)")
        
        if "previews" in st.session_state and st.session_state["previews"]:
            previews = st.session_state["previews"]
            
            if len(previews) > 1:
                page_num_sel = st.slider(
                    "เลือกหน้าที่ต้องการดูตัวอย่าง:", 
                    min_value=1, 
                    max_value=len(previews), 
                    value=1
                )
                selected_img = previews[page_num_sel - 1][1]
                st.image(
                    selected_img, 
                    caption=f"ตัวอย่างหน้าที่ {page_num_sel} (แสดงสูงสุด 5 หน้าแรก)", 
                    use_container_width=True
                )
            else:
                st.image(
                    previews[0][1], 
                    caption="ตัวอย่างหน้าแรก", 
                    use_container_width=True
                )
        else:
            st.info("💡 เมื่อกดเริ่มไฮไลต์เอกสาร ระบบจะแสดงพรีวิวหน้าแรกที่มีสีไฮไลต์ที่ตรงนี้")
            st.markdown(
                """
                <div style='text-align: center; padding: 40px; border: 2px dashed rgba(255,255,255,0.05); border-radius: 8px;'>
                    <span style='font-size: 4rem;'>📄</span>
                    <p style='color: #64748B; margin-top: 15px;'>ยังไม่มีเอกสารตัวอย่างที่จะพรีวิว</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
        st.markdown('</div>', unsafe_allow_html=True)

    # Footer
    st.markdown("""
        <div class="footer">
            KeyPDF Highlight & Analytics Studio &copy; 2026 | พัฒนาด้วย Streamlit & PyMuPDF (fitz)
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
