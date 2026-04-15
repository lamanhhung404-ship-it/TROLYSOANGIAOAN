import streamlit as st
from google import genai
from google.genai import types
import tempfile
import os
import io
import re
import json
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

# =========================
# 1. CẤU HÌNH CHUNG
# =========================
st.set_page_config(page_title="Trợ lý Giáo án NLS", page_icon="📘", layout="centered")

FILE_KHUNG_NANG_LUC = "khungnanglucso.pdf"
MODEL_NAME = "gemini-2.5-flash-lite"

TEN_TAC_GIA = "La Mạnh Hùng"
TEN_TRUONG = "Trường PTDTBT TH&THCS Nà Khương"
SO_DIEN_THOAI = "0388 667 404"
TEN_UNG_DUNG = "📘 TRỢ LÝ SOẠN GIÁO ÁN TỰ ĐỘNG (NLS)"


# =========================
# 2. HÀM TIỆN ÍCH
# =========================
def safe_html(text: str) -> str:
    if not text:
        return ""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def clean_text(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    banned_prefixes = [
        "tuyệt vời", "hy vọng", "nếu có bất kỳ", "tôi rất sẵn lòng",
        "dựa vào khung chương trình", "xin chào", "sau đây", "mình sẽ", "tôi sẽ",
    ]
    lower = s.lower()
    for bad in banned_prefixes:
        if lower.startswith(bad):
            return ""
    return s


def normalize_bullets(items):
    result = []
    if not items:
        return result
    for item in items:
        item = clean_text(item)
        if item:
            result.append(item)
    return result


def set_run_font(run, bold=False, size=14):
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    return run


def format_paragraph(paragraph, align=None, spacing_after=2, spacing_before=0, line_spacing=1.15):
    if align is not None:
        paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(spacing_before)
    paragraph.paragraph_format.space_after = Pt(spacing_after)
    paragraph.paragraph_format.line_spacing = line_spacing


def add_bullet_paragraph(container, text):
    p = container.add_paragraph(style="List Bullet")
    format_paragraph(p, spacing_after=1)
    r = p.add_run(text)
    set_run_font(r, size=14)
    return p


def add_normal_paragraph(container, text, bold=False, align=None):
    p = container.add_paragraph()
    format_paragraph(p, align=align, spacing_after=2)
    r = p.add_run(text)
    set_run_font(r, bold=bold, size=14)
    return p


def set_cell_text(cell, text, bold=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    format_paragraph(p, align=align, spacing_after=1)
    r = p.add_run(text)
    set_run_font(r, bold=bold, size=14)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return p


def extract_json_block(text: str) -> dict:
    if not text:
        raise ValueError("Model không trả về nội dung.")
    text = text.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Không tìm thấy JSON hợp lệ trong phản hồi của AI.")
    return json.loads(text[start:end + 1])


def validate_lesson_plan(data: dict):
    required_top_keys = [
        "ten_bai", "mon_hoc", "lop", "thoi_luong",
        "yeu_cau_can_dat", "do_dung_day_hoc",
        "tien_trinh_day_hoc", "dieu_chinh_sau_tiet_day"
    ]
    for key in required_top_keys:
        if key not in data:
            raise ValueError(f"Thiếu trường bắt buộc trong JSON: {key}")

    yccd = data["yeu_cau_can_dat"]
    for key in [
        "hoc_sinh_thuc_hien_duoc", "hoc_sinh_van_dung_duoc",
        "phat_trien_nang_luc", "phat_trien_pham_chat", "noi_dung_tich_hop"
    ]:
        if key not in yccd:
            raise ValueError(f"Thiếu mục trong 'yeu_cau_can_dat': {key}")

    ptnl = yccd["phat_trien_nang_luc"]
    for key in ["nang_luc_dac_thu", "nang_luc_chung", "nang_luc_so"]:
        if key not in ptnl:
            raise ValueError(f"Thiếu mục trong 'phat_trien_nang_luc': {key}")

    ndth = yccd["noi_dung_tich_hop"]
    for key in ["hoc_thong_qua_choi", "cong_dan_so"]:
        if key not in ndth:
            raise ValueError(f"Thiếu mục trong 'noi_dung_tich_hop': {key}")

    dddh = data["do_dung_day_hoc"]
    for key in ["giao_vien", "hoc_sinh"]:
        if key not in dddh:
            raise ValueError(f"Thiếu mục trong 'do_dung_day_hoc': {key}")

    activities = data["tien_trinh_day_hoc"]
    if not isinstance(activities, list) or len(activities) != 4:
        raise ValueError("Tiến trình dạy học phải có đúng 4 hoạt động.")

    expected_names = [
        "Hoạt động 1 - Khởi động",
        "Hoạt động 2 - Hình thành kiến thức mới",
        "Hoạt động 3 - Thực hành - luyện tập",
        "Hoạt động 4 - Vận dụng"
    ]
    for idx, act in enumerate(activities):
        for key in ["ten_hoat_dong", "thoi_gian", "giao_vien", "hoc_sinh"]:
            if key not in act:
                raise ValueError(f"Thiếu trường '{key}' trong hoạt động {idx+1}")
        if expected_names[idx].lower() not in act["ten_hoat_dong"].lower():
            raise ValueError(f"Tên hoạt động {idx+1} chưa đúng chuẩn.")
        if not isinstance(act["giao_vien"], list) or not act["giao_vien"]:
            raise ValueError(f"Hoạt động {idx+1} thiếu nội dung cột giáo viên.")
        if not isinstance(act["hoc_sinh"], list) or not act["hoc_sinh"]:
            raise ValueError(f"Hoạt động {idx+1} thiếu nội dung cột học sinh.")


def ensure_default_adjustments(data: dict):
    if not data.get("dieu_chinh_sau_tiet_day"):
        data["dieu_chinh_sau_tiet_day"] = [
            "Sau tiết dạy, giáo viên ghi nhận những điểm phù hợp và những nội dung cần điều chỉnh để tổ chức bài học hiệu quả hơn ở lần sau."
        ]
    return data


def upload_file_to_gemini(client, file_path, mime_type):
    return client.files.upload(
        file=file_path,
        config=types.UploadFileConfig(mime_type=mime_type)
    )


def build_contents(prompt_instruction, uploaded_refs, extra_text=None):
    parts = [types.Part(text=prompt_instruction)]
    for ref in uploaded_refs:
        parts.append(types.Part.from_uri(file_uri=ref.uri, mime_type=ref.mime_type))
    if extra_text:
        parts.append(types.Part(text=extra_text))
    return [types.Content(role="user", parts=parts)]


def render_preview_table(data: dict):
    st.markdown("### 📄 KẾT QUẢ BÀI SOẠN")
    st.markdown(f"**KẾ HOẠCH BÀI DẠY: {data.get('ten_bai', '').upper()}**")
    st.markdown(f"**Môn học:** {data.get('mon_hoc', '')}")
    st.markdown(f"**Lớp:** {data.get('lop', '')}")
    st.markdown(f"**Thời lượng:** {data.get('thoi_luong', '35 phút')}")

    yccd = data["yeu_cau_can_dat"]
    st.markdown("**I. Yêu cầu cần đạt**")
    st.markdown("**1. Học sinh thực hiện được**")
    for x in normalize_bullets(yccd["hoc_sinh_thuc_hien_duoc"]):
        st.markdown(f"- {x}")
    st.markdown("**2. Học sinh vận dụng được**")
    for x in normalize_bullets(yccd["hoc_sinh_van_dung_duoc"]):
        st.markdown(f"- {x}")
    st.markdown("**3. Phát triển năng lực**")
    st.markdown("**- Năng lực đặc thù**")
    for x in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_dac_thu"]):
        st.markdown(f"  - {x}")
    st.markdown("**- Năng lực chung**")
    for x in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_chung"]):
        st.markdown(f"  - {x}")
    st.markdown("**- Năng lực số**")
    for x in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_so"]):
        st.markdown(f"  - {x}")
    st.markdown("**4. Phát triển phẩm chất**")
    for x in normalize_bullets(yccd["phat_trien_pham_chat"]):
        st.markdown(f"- {x}")

    st.markdown("**II. Đồ dùng dạy học**")
    st.markdown("**1. Giáo viên**")
    for x in normalize_bullets(data["do_dung_day_hoc"]["giao_vien"]):
        st.markdown(f"- {x}")
    st.markdown("**2. Học sinh**")
    for x in normalize_bullets(data["do_dung_day_hoc"]["hoc_sinh"]):
        st.markdown(f"- {x}")

    st.markdown("**III. Tiến trình dạy học**")
    md = ["| HOẠT ĐỘNG CỦA GIÁO VIÊN | HOẠT ĐỘNG CỦA HỌC SINH |", "|---|---|"]
    for act in data["tien_trinh_day_hoc"]:
        left = f"**{act['ten_hoat_dong']} ({act['thoi_gian']})**<br>" + "<br>".join(
            [f"- {x}" for x in normalize_bullets(act["giao_vien"])]
        )
        right = "<br>".join([f"- {x}" for x in normalize_bullets(act["hoc_sinh"])])
        md.append(f"| {left} | {right} |")
    st.markdown("\n".join(md), unsafe_allow_html=True)

    st.markdown("**IV. Điều chỉnh sau tiết dạy**")
    for x in normalize_bullets(data["dieu_chinh_sau_tiet_day"]):
        st.markdown(f"- {x}")


def create_doc_from_json(data: dict):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(1.5)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style.paragraph_format.line_spacing = 1.15

    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = head.add_run(f"KẾ HOẠCH BÀI DẠY: {data.get('ten_bai', '').upper()}")
    set_run_font(r, bold=True, size=14)
    r.font.color.rgb = RGBColor(0, 0, 0)

    add_normal_paragraph(doc, f"Môn học: {data.get('mon_hoc', '')}")
    add_normal_paragraph(doc, f"Lớp: {data.get('lop', '')}")
    add_normal_paragraph(doc, f"Thời lượng: {data.get('thoi_luong', '35 phút')}")

    add_normal_paragraph(doc, "I. Yêu cầu cần đạt", bold=True)
    yccd = data["yeu_cau_can_dat"]

    add_normal_paragraph(doc, "1. Học sinh thực hiện được", bold=True)
    for item in normalize_bullets(yccd["hoc_sinh_thuc_hien_duoc"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "2. Học sinh vận dụng được", bold=True)
    for item in normalize_bullets(yccd["hoc_sinh_van_dung_duoc"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "3. Phát triển năng lực", bold=True)
    add_normal_paragraph(doc, "- Năng lực đặc thù", bold=True)
    for item in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_dac_thu"]):
        add_bullet_paragraph(doc, item)
    add_normal_paragraph(doc, "- Năng lực chung", bold=True)
    for item in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_chung"]):
        add_bullet_paragraph(doc, item)
    add_normal_paragraph(doc, "- Năng lực số", bold=True)
    for item in normalize_bullets(yccd["phat_trien_nang_luc"]["nang_luc_so"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "4. Phát triển phẩm chất", bold=True)
    for item in normalize_bullets(yccd["phat_trien_pham_chat"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "* Nội dung tích hợp", bold=True)
    add_normal_paragraph(doc, "- Học thông qua chơi", bold=True)
    for item in normalize_bullets(yccd["noi_dung_tich_hop"]["hoc_thong_qua_choi"]):
        add_bullet_paragraph(doc, item)
    add_normal_paragraph(doc, "- Công dân số", bold=True)
    for item in normalize_bullets(yccd["noi_dung_tich_hop"]["cong_dan_so"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "II. Đồ dùng dạy học", bold=True)
    add_normal_paragraph(doc, "1. Giáo viên", bold=True)
    for item in normalize_bullets(data["do_dung_day_hoc"]["giao_vien"]):
        add_bullet_paragraph(doc, item)
    add_normal_paragraph(doc, "2. Học sinh", bold=True)
    for item in normalize_bullets(data["do_dung_day_hoc"]["hoc_sinh"]):
        add_bullet_paragraph(doc, item)

    add_normal_paragraph(doc, "III. Tiến trình dạy học", bold=True)

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # set fixed widths for 2 columns
    for row in table.rows:
        row.cells[0].width = Cm(8.75)
        row.cells[1].width = Cm(8.75)

    hdr = table.rows[0].cells
    set_cell_text(hdr[0], "HOẠT ĐỘNG CỦA GIÁO VIÊN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(hdr[1], "HOẠT ĐỘNG CỦA HỌC SINH", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    for act in data["tien_trinh_day_hoc"]:
        row = table.add_row()
        left = row.cells[0]
        right = row.cells[1]
        left.width = Cm(8.75)
        right.width = Cm(8.75)

        left.text = ""
        right.text = ""

        # Left cell: GV
        p_left = left.paragraphs[0]
        format_paragraph(p_left, spacing_after=2)
        r_left = p_left.add_run(f"{act['ten_hoat_dong']} ({act['thoi_gian']})")
        set_run_font(r_left, bold=True, size=14)
        for item in normalize_bullets(act["giao_vien"]):
            add_bullet_paragraph(left, item)

        # Right cell: HS
        p_right = right.paragraphs[0]
        format_paragraph(p_right, spacing_after=2)
        r_right = p_right.add_run("Hoạt động của học sinh")
        set_run_font(r_right, bold=True, size=14)
        for item in normalize_bullets(act["hoc_sinh"]):
            add_bullet_paragraph(right, item)

        left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    add_normal_paragraph(doc, "IV. Điều chỉnh sau tiết dạy", bold=True)
    for item in normalize_bullets(data["dieu_chinh_sau_tiet_day"]):
        add_bullet_paragraph(doc, item)

    return doc


def generate_lesson_prompt(ten_bai, lop, noidung_bosung, yeu_cau_them):
    return f"""
Bạn là công cụ sinh giáo án tiểu học theo đúng mẫu quy định.
NHIỆM VỤ: Sinh duy nhất 1 JSON hợp lệ. KHÔNG viết lời mở đầu. KHÔNG giải thích. KHÔNG kết luận. KHÔNG dùng markdown. KHÔNG dùng code fence.

Dữ liệu bài dạy:
- Tên bài: {ten_bai}
- Lớp: {lop}
- Ghi chú bổ sung: {noidung_bosung}
- Yêu cầu thêm: {yeu_cau_them}

Yêu cầu bắt buộc:
1. Bám đúng cấu trúc:
   I. Yêu cầu cần đạt
   II. Đồ dùng dạy học
   III. Tiến trình dạy học
   IV. Điều chỉnh sau tiết dạy

2. Mục III phải có đúng 4 hoạt động:
   - Hoạt động 1 - Khởi động
   - Hoạt động 2 - Hình thành kiến thức mới
   - Hoạt động 3 - Thực hành - luyện tập
   - Hoạt động 4 - Vận dụng

3. Mỗi hoạt động phải có:
   - ten_hoat_dong
   - thoi_gian
   - giao_vien: danh sách các ý riêng của giáo viên
   - hoc_sinh: danh sách các ý riêng của học sinh

4. Phần III bắt buộc để app xuất thành bảng Word 2 cột:
   - Cột 1: HOẠT ĐỘNG CỦA GIÁO VIÊN
   - Cột 2: HOẠT ĐỘNG CỦA HỌC SINH

5. Không được gộp hoạt động của giáo viên và học sinh vào chung 1 danh sách.
6. Nếu có trò chơi thì luật chơi phải nằm ở cột giáo viên.
7. Tổng thời lượng 35 phút.
8. Không được sinh các câu xã giao như Tuyệt vời, Hy vọng, Nếu cần...

Chỉ trả về JSON hợp lệ theo schema:
{{
  "ten_bai": "{ten_bai}",
  "mon_hoc": "...",
  "lop": "{lop}",
  "thoi_luong": "35 phút",
  "yeu_cau_can_dat": {{
    "hoc_sinh_thuc_hien_duoc": ["..."],
    "hoc_sinh_van_dung_duoc": ["..."],
    "phat_trien_nang_luc": {{
      "nang_luc_dac_thu": ["..."],
      "nang_luc_chung": ["..."],
      "nang_luc_so": ["..."]
    }},
    "phat_trien_pham_chat": ["..."],
    "noi_dung_tich_hop": {{
      "hoc_thong_qua_choi": ["..."],
      "cong_dan_so": ["..."]
    }}
  }},
  "do_dung_day_hoc": {{
    "giao_vien": ["..."],
    "hoc_sinh": ["..."]
  }},
  "tien_trinh_day_hoc": [
    {{
      "ten_hoat_dong": "Hoạt động 1 - Khởi động",
      "thoi_gian": "5 phút",
      "giao_vien": ["..."],
      "hoc_sinh": ["..."]
    }},
    {{
      "ten_hoat_dong": "Hoạt động 2 - Hình thành kiến thức mới",
      "thoi_gian": "15 phút",
      "giao_vien": ["..."],
      "hoc_sinh": ["..."]
    }},
    {{
      "ten_hoat_dong": "Hoạt động 3 - Thực hành - luyện tập",
      "thoi_gian": "10 phút",
      "giao_vien": ["..."],
      "hoc_sinh": ["..."]
    }},
    {{
      "ten_hoat_dong": "Hoạt động 4 - Vận dụng",
      "thoi_gian": "5 phút",
      "giao_vien": ["..."],
      "hoc_sinh": ["..."]
    }}
  ],
  "dieu_chinh_sau_tiet_day": ["..."]
}}
"""


# =========================
# 3. CSS GIAO DIỆN
# =========================
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
    div.stButton > button {
        background: linear-gradient(90deg, #11998e, #38ef7d);
        color: white !important; border: none; padding: 15px 30px; font-weight: bold;
        border-radius: 10px; width: 100%; margin-top: 10px; font-size: 18px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0,0,0,0.3);
    }
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="main-header">
    <h1>{TEN_UNG_DUNG}</h1>
    <p>Tác giả: {TEN_TAC_GIA} - {TEN_TRUONG} - ĐT: {SO_DIEN_THOAI}</p>
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
    if not api_key:
        st.error("Thiếu API Key.")
    elif not client:
        st.error("Không khởi tạo được Gemini client.")
    elif not uploaded_files and not noidung_bosung and not has_framework:
        st.warning("Thiếu tài liệu đầu vào.")
    elif not ten_bai.strip():
        st.warning("Vui lòng nhập tên bài học.")
    else:
        temp_paths = []
        uploaded_refs = []
        try:
            with st.spinner("AI đang soạn giáo án theo đúng biểu mẫu..."):
                prompt_instruction = generate_lesson_prompt(ten_bai, lop, noidung_bosung, yeu_cau_them)

                if has_framework and os.path.exists(FILE_KHUNG_NANG_LUC):
                    uploaded_refs.append(upload_file_to_gemini(client, FILE_KHUNG_NANG_LUC, "application/pdf"))

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
                        temperature=0.1,
                        top_p=0.8,
                        max_output_tokens=8192
                    )
                )

                result_text = getattr(response, "text", None)
                if not result_text:
                    raise ValueError("Model không trả về nội dung.")

                lesson_plan_data = extract_json_block(result_text)
                lesson_plan_data = ensure_default_adjustments(lesson_plan_data)
                validate_lesson_plan(lesson_plan_data)

                render_preview_table(lesson_plan_data)

                doc = create_doc_from_json(lesson_plan_data)
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
    f"<div style='text-align: center; color: #666;'>© 2025 - {TEN_TAC_GIA} - {TEN_TRUONG} - ĐT: {SO_DIEN_THOAI}</div>",
    unsafe_allow_html=True
)
