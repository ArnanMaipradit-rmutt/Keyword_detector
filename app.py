import os
import io
import csv
import fitz  # PyMuPDF
import streamlit as st
import pandas as pd
from PIL import Image

# Set page configuration with a premium look
st.set_page_config(
    page_title="PDF Native Keyword Highlighter",
    page_icon="🖊️",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    
    # Try multiple encodings, including Thai standards
    encodings = ['utf-8-sig', 'utf-8', 'cp1252', 'tis-620', 'iso-8859-11']
    for enc in encodings:
        try:
            content = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        # Fallback to system default with error ignoring
        content = file_bytes.decode('utf-8', errors='ignore')

    if filename.lower().endswith('.csv'):
        # Parse CSV file
        reader = csv.reader(io.StringIO(content))
        for row in reader:
            for cell in row:
                word = cell.strip()
                if word and word not in keywords:
                    keywords.append(word)
    else:
        # Parse plain text (one keyword per line)
        for line in content.splitlines():
            word = line.strip()
            if word and word not in keywords:
                keywords.append(word)
                
    return keywords


def get_hex_color(color_name):
    """
    Returns float RGB values (0.0 to 1.0) and hex code for PyMuPDF highlighting.
    """
    colors = {
        "เหลือง (Yellow)": ((1.0, 1.0, 0.0), "#FFFF00"),
        "เขียวสว่าง (Mint)": ((0.2, 0.9, 0.4), "#33E666"),
        "ฟ้าคราม (Cyan)": ((0.0, 0.8, 1.0), "#00CCFF"),
        "ชมพู (Pink)": ((1.0, 0.4, 0.7), "#FF66B2"),
        "ส้ม (Orange)": ((1.0, 0.6, 0.0), "#FF9900")
    }
    return colors.get(color_name, ((1.0, 1.0, 0.0), "#FFFF00"))


def process_pdf(pdf_bytes, keywords, search_flags, highlight_rgb, case_sensitive):
    """
    Performs native search and highlighting using PyMuPDF.
    Returns:
      - highlighted_pdf_bytes (bytes)
      - summary_data (list of dicts)
      - report_data (list of dicts with detailed positions)
      - total_matches (int)
      - preview_images (list of (page_num, PIL.Image))
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # Track statistics
    summary_data = []
    report_data = []
    total_matches = 0
    keyword_match_counts = {kw: 0 for kw in keywords}
    
    # Process page by page
    for page_num in range(len(doc)):
        page = doc[page_num]
        page_1based = page_num + 1
        
        # Optimize performance by creating TextPage once per page
        textpage = page.get_textpage(flags=search_flags)
        
        # Get all words for line-number and context-text reconstruction
        words = page.get_text("words")
        
        # 1. Map words to lines and build line text context
        line_coords = {}
        line_words = {}
        for w in words:
            # w: (x0, y0, x1, y1, word_str, block_no, line_no, word_no)
            key = (w[5], w[6])  # (block_no, line_no)
            y0 = w[1]
            if key not in line_coords:
                line_coords[key] = []
                line_words[key] = []
            line_coords[key].append(y0)
            line_words[key].append(w)
        
        # Sort unique lines by their average y0 to assign a 1-based global line number on the page
        line_avg_y = {key: sum(y0s)/len(y0s) for key, y0s in line_coords.items()}
        sorted_lines = sorted(line_avg_y.items(), key=lambda item: item[1])
        line_map = {key: idx + 1 for idx, (key, _) in enumerate(sorted_lines)}
        
        # Reconstruct full line texts
        line_text_map = {}
        for key, ws in line_words.items():
            ws_sorted = sorted(ws, key=lambda x: x[7]) # Sort by word_no
            line_text_map[key] = " ".join(x[4] for x in ws_sorted)
            
        # 2. Native text search
        for kw in keywords:
            rects = page.search_for(kw, textpage=textpage)
            
            if rects:
                # Filter rects case-sensitively if requested
                if case_sensitive:
                    valid_rects = []
                    for rect in rects:
                        text_in_rect = page.get_text("text", clip=rect).strip()
                        # Normalize whitespace for clean comparison
                        text_clean = " ".join(text_in_rect.split())
                        kw_clean = " ".join(kw.split())
                        if text_clean == kw_clean:
                            valid_rects.append(rect)
                    rects = valid_rects
                
                if rects:
                    count_on_page = len(rects)
                    total_matches += count_on_page
                    keyword_match_counts[kw] += count_on_page
                    
                    # Annotate each matched rectangle and record detailed findings
                    for rect in rects:
                        # Draw highlight annotation
                        annot = page.add_highlight_annot(rect)
                        if annot:
                            annot.set_colors(stroke=highlight_rgb)
                            annot.update()
                            
                        # Find overlapping word to associate with its line number
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
                        
    # Build summary list
    for kw, count in keyword_match_counts.items():
        summary_data.append({"Keyword": kw, "Matches Found": count})
        
    # Save the modified PDF to memory buffer
    output_buffer = io.BytesIO()
    doc.save(output_buffer, garbage=4, deflate=True)
    highlighted_pdf_bytes = output_buffer.getvalue()
    
    # Create image views of the PDF for preview
    preview_images = []
    # Generate previews for up to the first 5 pages with highlights
    preview_limit = min(5, len(doc))
    for page_num in range(preview_limit):
        page = doc[page_num]
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        preview_images.append((page_num + 1, Image.open(io.BytesIO(img_data))))
        
    doc.close()
    
    return highlighted_pdf_bytes, summary_data, report_data, total_matches, preview_images


def generate_csv_bytes(summary, report_data):
    """
    Generates CSV content as bytes with UTF-8 BOM encoding for Excel compatibility.
    """
    output = io.StringIO()
    writer = csv.writer(output)
    
    # 1. Summary Block
    writer.writerow(["--- PDF KEYWORD SEARCH SUMMARY ---"])
    writer.writerow(["Keyword", "Matches Count"])
    for item in summary:
        writer.writerow([item["Keyword"], item["Matches Found"]])
        
    writer.writerow([]) # blank line spacing
    
    # 2. Detailed Findings Block
    writer.writerow(["--- DETAILED FINDINGS ---"])
    writer.writerow(["Keyword", "Page Number", "Line Number (Top-to-Bottom)", "Context Line Text"])
    for row in report_data:
        writer.writerow([
            row["Keyword"],
            row["Page"],
            row["Line"],
            row["Context"]
        ])
        
    return output.getvalue().encode('utf-8-sig')


# --- STREAMLIT UI ---

def main():
    st.markdown('<div class="title-gradient">PDF Native Keyword Highlighter</div>', unsafe_allow_html=True)
    st.markdown("<p style='color: #94A3B8; font-size: 1.1rem; margin-top:-10px;'>ไฮไลต์คำสำคัญบนเอกสาร PDF ค้นหาและระบุตำแหน่งได้อย่างแม่นยำด้วยการค้นหาแบบ Native (ไม่ต้องผ่าน OCR)</p>", unsafe_allow_html=True)
    st.write("---")
    
    # Sidebar: Configurations and File Uploads
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
            placeholder="ตัวอย่างเช่น:\nสัญญาจ้าง\nบริษัท จำกัด\nผู้ว่าจ้าง",
            height=150
        )
        if text_kw:
            keywords = [line.strip() for line in text_kw.splitlines() if line.strip()]

    # Display status of loaded keywords
    if keywords:
        st.sidebar.markdown(f"<span class='highlight-badge'>โหลดคีย์เวิร์ดสำเร็จ: {len(keywords)} คำ</span>", unsafe_allow_html=True)
    
    st.sidebar.write("")
    st.sidebar.markdown("### 🎨 ตั้งค่าการค้นหาและสีไฮไลต์")
    
    # Case sensitivity & other flags
    case_sensitive = st.sidebar.checkbox(
        "Case Sensitive (ตรวจหาตัวพิมพ์เล็ก-ใหญ่)", 
        value=False,
        help="หากเปิดใช้งาน คำว่า 'App' และ 'app' จะมองเป็นคนละคำกัน (รองรับภาษาอังกฤษเป็นหลัก)"
    )
    
    dehyphenate = st.sidebar.checkbox(
        "Dehyphenate (รวมคำที่แยกด้วยยติภังค์ท้ายบรรทัด)", 
        value=True,
        help="พิจารณาคำที่ถูกฉีกท้ายบรรทัดโดยมีเครื่องหมาย - ให้เป็นคำเดียวกันในการค้นหา"
    )
    
    # Highlight color
    color_choice = st.sidebar.selectbox(
        "สีของไฮไลต์ (Highlight Color):",
        ["เหลือง (Yellow)", "เขียวสว่าง (Mint)", "ฟ้าคราม (Cyan)", "ชมพู (Pink)", "ส้ม (Orange)"],
        index=0
    )
    highlight_rgb, highlight_hex = get_hex_color(color_choice)
    
    # Setup search flags
    search_flags = 2 | 1 # TEXT_PRESERVE_WHITESPACE | TEXT_PRESERVE_LIGATURES
    if dehyphenate:
        search_flags |= 16  # TEXT_DEHYPHENATE
    
    # Main panel
    col1, col2 = st.columns([1, 1.2])
    
    with col1:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("📋 รายละเอียดการไฮไลต์")
        
        # Validation checks
        if not uploaded_pdf:
            st.info("👋 กรุณาอัปโหลดไฟล์ PDF ในแผงควบคุมด้านซ้ายเพื่อเริ่มต้น")
        elif not keywords:
            st.warning("⚠️ กรุณาระบุหรืออัปโหลดคีย์เวิร์ดคำสำคัญ")
        else:
            # Ready to run
            pdf_name = uploaded_pdf.name
            st.write(f"📂 **ไฟล์ PDF ที่จะประมวลผล:** `{pdf_name}`")
            st.write(f"🔑 **จำนวนคีย์เวิร์ด:** `{len(keywords)} คำ`")
            st.write(f"🎨 **สีไฮไลต์:** <span style='color:{highlight_hex}; font-weight:bold;'>■ {color_choice}</span>", unsafe_allow_html=True)
            
            run_btn = st.button("🚀 เริ่มไฮไลต์เอกสาร (Highlight Document)")
            
            # Processing state
            if run_btn:
                with st.spinner("⏳ กำลังค้นหาและไฮไลต์คำสำคัญบนหน้า PDF..."):
                    try:
                        pdf_bytes = uploaded_pdf.read()
                        
                        # Process PDF using native engine
                        res_bytes, summary, report_data, count, previews = process_pdf(
                            pdf_bytes, 
                            keywords, 
                            search_flags, 
                            highlight_rgb,
                            case_sensitive
                        )
                        
                        # Generate CSV report bytes
                        csv_report_bytes = generate_csv_bytes(summary, report_data)
                        
                        # Store in session state to maintain results between interactions
                        st.session_state["result_pdf"] = res_bytes
                        st.session_state["result_csv"] = csv_report_bytes
                        st.session_state["summary"] = summary
                        st.session_state["total_matches"] = count
                        st.session_state["previews"] = previews
                        st.session_state["processed_name"] = f"{os.path.splitext(pdf_name)[0]}_highlighted.pdf"
                        st.session_state["csv_name"] = f"{os.path.splitext(pdf_name)[0]}_keyword_report.csv"
                        
                        st.success(f"✨ ประมวลผลเสร็จสิ้น! ไฮไลต์คำสำคัญพบทั้งหมด {count} จุด")
                        
                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล PDF: {str(e)}")
                        
        # Display results if processed
        if "result_pdf" in st.session_state:
            st.write("---")
            st.markdown(f"#### 🎉 สรุปผลลัพธ์ (Total Matches: {st.session_state['total_matches']})")
            
            # 1. Download Highlighted PDF
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ PDF ที่ไฮไลต์เสร็จแล้ว",
                data=st.session_state["result_pdf"],
                file_name=st.session_state["processed_name"],
                mime="application/pdf",
                use_container_width=True
            )
            
            st.write("")
            
            # 2. Download Detailed CSV Report
            st.download_button(
                label="📊 ดาวน์โหลดรายงานรายละเอียด (CSV)",
                data=st.session_state["result_csv"],
                file_name=st.session_state["csv_name"],
                mime="text/csv",
                use_container_width=True
            )
            
            # Keyword breakdown table
            st.write("")
            st.write("**ตารางแจกแจงจำนวนคำที่พบ:**")
            df = pd.DataFrame(st.session_state["summary"])
            # Highlight non-zero counts
            st.dataframe(
                df.style.map(
                    lambda v: 'color: #34D399; font-weight: bold;' if isinstance(v, int) and v > 0 else '',
                    subset=["Matches Found"]
                ),
                use_container_width=True,
                hide_index=True
            )
            
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col2:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("👁️ พรีวิวหน้าเอกสาร (PDF Preview)")
        
        if "previews" in st.session_state and st.session_state["previews"]:
            previews = st.session_state["previews"]
            
            # If multi-page, add page selector
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
            # Show aesthetic illustration placeholder
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
            PDF Native Keyword Highlighter Application &copy; 2026 | พัฒนาด้วย Streamlit & PyMuPDF (fitz)
        </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
