import streamlit as st
from google import genai
from google.genai import types
from PIL import Image
import tempfile
import os
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor, Cm

# --- 1. CẤU HÌNH TRANG ---
st.set_page_config(page_title="Trợ lý Giáo án NLS", page_icon="📘", layout="centered")

FILE_KHUNG_NANG_LUC = "khungnanglucso.pdf"
MODEL_NAME = "gemini-2.5-flash-lite"

# --- 2. HÀM XỬ LÝ WORD ---
def add_formatted_text(paragraph, text):
    """Hàm in đậm và ép font Times New Roman"""
    paragraph.style.font.name = "Times New Roman"
    paragraph.style.font.size = Pt(14)

    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            clean = part[2:-2]
            run = paragraph.add_run(clean)
            run.bold = True
        else:
            run = paragraph.add_run(part)
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)

def create_doc_stable(content, ten_bai, lop):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(14)
    style.paragraph_format.line_spacing = 1.2

    head = doc.add_heading(f"KẾ HOẠCH BÀI DẠY: {ten_bai.upper()}", 0)
    head.alignment = 1
    for run in head.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)

    p_lop = doc.add_paragraph(f"Lớp: {lop}")
    p_lop.alignment = 1
    p_lop.runs[0].bold = True

    doc.add_paragraph("-" * 60).alignment = 1

    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("#"):
            line = line.replace("#", "").strip()

        if line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1

            if len(table_lines) >= 3:
                try:
                    valid_rows = [r for r in table_lines if "---" not in r]
                    if valid_rows:
                        cols_count = len(valid_rows[0].split("|")) - 2
                        if cols_count > 0:
                            table = doc.add_table(rows=len(valid_rows), cols=cols_count)
                            table.style = "Table Grid"
                            table.autofit = True

                            for r_idx, r_text in enumerate(valid_rows):
                                cells_data = r_text.split("|")[1:-1]
                                for c_idx, cell_text in enumerate(cells_data):
                                    if c_idx < cols_count:
                                        cell = table.cell(r_idx, c_idx)
                                        cell._element.clear_content()

                                        raw_content = cell_text.strip().replace("<br>", "\n").replace("<br/>", "\n")
                                        sub_lines = raw_content.split("\n")

                                        for sub_line in sub_lines:
                                            sub_line = sub_line.strip()
                                            if not sub_line:
                                                continue

                                            p = cell.add_paragraph()
                                            p.paragraph_format.space_before = Pt(0)
                                            p.paragraph_format.space_after = Pt(2)
                                            p.paragraph_format.line_spacing = 1.1

                                            if r_idx == 0:
                                                p.alignment = 1
                                                run = p.add_run(sub_line.replace("**", ""))
                                                run.bold = True
                                                run.font.name = "Times New Roman"
                                                run.font.size = Pt(14)
                                            else:
                                                add_formatted_text(p, sub_line)
                except Exception:
                    pass
            continue

        if not line:
            i += 1
            continue

        if re.match(r"^(I\.|II\.|III\.|IV\.|V\.)", line) or (re.match(r"^\d+\.", line) and len(line) < 50):
            clean = line.replace("**", "").strip()
            p = doc.add_paragraph(clean)
            p.runs[0].bold = True
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(14)

        elif line.startswith("- "):
            clean = line[2:].strip()
            p = doc.add_paragraph(style="List Bullet")
            add_formatted_text(p, clean)

        else:
            p = doc.add_paragraph()
            add_formatted_text(p, line)

        i += 1

    return doc

# --- 3. HÀM CHUYỂN FILE LÊN GEMINI ---
def upload_file_to_gemini(client, file_path, mime_type):
    uploaded = client.files.upload(
        file=file_path,
        config=types.UploadFileConfig(mime_type=mime_type)
    )
    return uploaded

def build_contents(prompt_instruction, uploaded_refs, extra_text=None):
    parts = [types.Part(text=prompt_instruction)]

    for ref in uploaded_refs:
        parts.append(types.Part.from_uri(file_uri=ref.uri, mime_type=ref.mime_type))

    if extra_text:
        parts.append(types.Part(text=extra_text))

    return [types.Content(role="user", parts=parts)]

# --- 4. CSS GIAO DIỆN ---
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { background-color: #f4f6f9; }

    .main-header {
        background: linear-gradient(135deg, #004e92 0%, #000428 100%);
        padding: 30px; border-radius: 15px; text-align: center; color: white !important;
        margin-bottom: 30px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .main-header h1 { color: white !important; margin: 0; font-size: 2rem; }
    .main-header p { color: #e0e0e0 !important; margin-top: 10px; font-style: italic; }

    .section-header {
        color: #004e92; border-bottom: 2px solid #ddd; padding-bottom: 5px;
        margin-top: 20px; margin-bottom: 15px; font-weight: bold;
    }

    .lesson-plan-paper {
        background-color: white; padding: 40px; border-radius: 5px;
        border: 1px solid #ccc; box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        font-family: 'Times New Roman', Times, serif !important;
        font-size: 14pt !important;
        line-height: 1.5 !important;
        color: #000000 !important;
        text-align: justify;
    }

    div.stButton > button {
        background: linear-gradient(90deg, #11998e, #38ef7d);
        color: white !important;
        border: none;
        padding: 15px 30px;
        font-weight: bold;
        border-radius: 10px;
        width: 100%;
        margin-top: 10px;
        font-size: 18px;
    }
</style>
""", unsafe_allow_html=True)

# --- 5. GIAO DIỆN CHÍNH ---
st.markdown("""
<div class="main-header">
    <h1>📘 TRỢ LÝ SOẠN GIÁO ÁN TỰ ĐỘNG (NLS)</h1>
    <p>Tác giả: La Mạnh Hùng - Trường PTDTBT Tiểu học Nà Khương - ĐT: 0388 667 404</p>
</div>
""", unsafe_allow_html=True)

api_key = ""
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    with st.sidebar:
        st.header("🔐 Cấu hình")
        api_key = st.text_input("Nhập API Key:", type="password")

client = None
if api_key:
    try:
        client = genai.Client(api_key=api_key)
    except Exception as e:
        st.error(f"Lỗi khởi tạo Gemini client: {e}")

st.markdown('<div class="section-header">📂 1. TÀI LIỆU NGUỒN</div>', unsafe_allow_html=True)

has_framework = False
if os.path.exists(FILE_KHUNG_NANG_LUC):
    st.success(f"✅ Đã tự động tích hợp: {FILE_KHUNG_NANG_LUC}")
    has_framework = True
else:
    st.info(f"ℹ️ Chưa có file '{FILE_KHUNG_NANG_LUC}'. Có thể upload thêm để dùng tính năng Năng lực số.")

uploaded_files = st.file_uploader(
    "Tải Ảnh/PDF bài dạy (Kéo thả vào đây):",
    type=["jpg", "jpeg", "png", "pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.caption("👁️ Xem trước tài liệu:")
    cols = st.columns(3)
    for i, f in enumerate(uploaded_files):
        if f.type in ["image/jpeg", "image/png"]:
            with cols[i % 3]:
                st.image(f, caption=f.name)
        else:
            with cols[i % 3]:
                st.info(f"📄 {f.name}")

st.markdown('<div class="section-header">📝 2. THÔNG TIN BÀI DẠY</div>', unsafe_allow_html=True)

c1, c2 = st.columns(2)
with c1:
    lop = st.text_input("📚 Lớp:", "Lớp 4")
with c2:
    ten_bai = st.text_input("📌 Tên bài học:", placeholder="Ví dụ: Học hát bài...")

noidung_bosung = st.text_area("✍️ Ghi chú thêm (nội dung/kiến thức):", height=100)
yeu_cau_them = st.text_input("💡 Yêu cầu đặc biệt:", placeholder="Ví dụ: Tích hợp trò chơi khởi động...")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 SOẠN GIÁO ÁN NGAY"):
    if not api_key or not client:
        st.error("Thiếu hoặc lỗi API Key.")
    elif not uploaded_files and not noidung_bosung and not has_framework:
        st.warning("Thiếu tài liệu đầu vào.")
    elif not ten_bai.strip():
        st.warning("Vui lòng nhập tên bài học.")
    else:
        temp_paths = []
        uploaded_refs = []

        try:
            with st.spinner("AI đang soạn giáo án..."):
                prompt_instruction = f"""
Đóng vai là một Giáo viên Tiểu học giỏi, am hiểu chương trình GDPT 2018.
Nhiệm vụ: Soạn Kế hoạch bài dạy (Giáo án) cho bài: "{ten_bai}" - {lop}.

DỮ LIỆU ĐẦU VÀO:
- (Nếu có) File PDF Khung năng lực số đính kèm: Hãy dùng để đối chiếu nội dung bài học và đưa vào mục Năng lực số.
- Các tài liệu hình ảnh/PDF tải lên: Phân tích để lấy nội dung kiến thức bài học.
- Ghi chú bổ sung: "{noidung_bosung}".

YÊU CẦU LUÔN LUÔN TUÂN THỦ CẤU TRÚC (CÔNG VĂN 2345):
I. Yêu cầu cần đạt:
1. Học sinh thực hiện được
2. Học sinh vận dụng được
3. Phát triển năng lực (bao gồm năng lực đặc thù, năng lực chung, phát triển năng lực số)
4. Phát triển phẩm chất
* Nội dung tích hợp (VD: Học thông qua chơi, Công dân số)

II. Đồ dùng dạy học
1. Giáo viên
2. Học sinh

III. Tiến trình dạy học
PHẢI trình bày dưới dạng bảng markdown 2 cột:

| HOẠT ĐỘNG CỦA GIÁO VIÊN | HOẠT ĐỘNG CỦA HỌC SINH |
|---|---|
| **1. Hoạt động 1 - Khởi động:**<br>- GV tổ chức... | - HS tham gia... |
| **2. Hoạt động 2 - Hình thành kiến thức mới:**<br>- GV hướng dẫn... | - HS quan sát... |
| **3. Hoạt động 3 - Thực hành - luyện tập:**<br>- GV yêu cầu... | - HS thực hiện... |
| **4. Hoạt động 4 - Vận dụng:**<br>- GV gợi mở... | - HS chia sẻ... |

YÊU CẦU CHI TIẾT:
- Chi tiết, cụ thể, đặc biệt là hoạt động của học sinh.
- Các ý bắt đầu bằng dấu gạch đầu dòng (-).
- Tích hợp Học thông qua chơi vào hoạt động phù hợp.
- Nếu có trò chơi, phải nêu rõ luật chơi.
- Chỉ có 35 phút.
- Chỉ gồm đúng 4 hoạt động.
- Không dùng ký tự # ở đầu dòng.
- Không thêm chú thích nguồn vào bài soạn.

IV. Điều chỉnh sau tiết dạy

LƯU Ý QUAN TRỌNG TỪ NGƯỜI DÙNG: {yeu_cau_them}
"""

                if has_framework and os.path.exists(FILE_KHUNG_NANG_LUC):
                    uploaded_refs.append(
                        upload_file_to_gemini(client, FILE_KHUNG_NANG_LUC, "application/pdf")
                    )

                if uploaded_files:
                    for f in uploaded_files:
                        suffix = os.path.splitext(f.name)[1] or ".tmp"
                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                            tmp.write(f.getvalue())
                            tmp_path = tmp.name
                            temp_paths.append(tmp_path)

                        mime_type = f.type if f.type else "application/octet-stream"
                        uploaded_refs.append(upload_file_to_gemini(client, tmp_path, mime_type))

                contents = build_contents(
                    prompt_instruction=prompt_instruction,
                    uploaded_refs=uploaded_refs,
                    extra_text=noidung_bosung if noidung_bosung else None
                )

                response = client.models.generate_content(
                    model=MODEL_NAME,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        temperature=0.7,
                        top_p=0.95,
                        max_output_tokens=8192
                    )
                )

                result_text = response.text if hasattr(response, "text") and response.text else "Không có nội dung trả về."

                st.markdown("### 📄 KẾT QUẢ BÀI SOẠN:")
                st.markdown(f'<div class="lesson-plan-paper">{result_text}</div>', unsafe_allow_html=True)

                doc = create_doc_stable(result_text, ten_bai, lop)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                safe_file_name = re.sub(r'[\\\\/*?:"<>|]', "_", ten_bai).strip() or "GiaoAn"
                st.download_button(
                    label="⬇️ TẢI FILE WORD CHUẨN A4",
                    data=buf,
                    file_name=f"GiaoAn_{safe_file_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary"
                )

        except Exception as e:
            st.error(f"Có lỗi xảy ra: {e}")

        finally:
            for p in temp_paths:
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666;'>© 2025 - La Mạnh Hùng - Trường PTDTBT Tiểu học Nà Khương - ĐT: 0388 667 404</div>",
    unsafe_allow_html=True
)