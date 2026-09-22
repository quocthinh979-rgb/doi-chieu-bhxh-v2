import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import datetime

# ===== CẤU HÌNH TRANG =====
st.set_page_config(
    page_title="Đối chiếu BHXH - BHYT - BHTN",
    page_icon="📊",
    layout="wide"
)

# ===== CSS =====
st.markdown("""
<style>
    .header-box {
        background: linear-gradient(90deg, #000080 0%, #1a1a9e 100%);
        padding: 15px 20px; border-radius: 8px; color: white; margin-bottom: 20px;
    }
    .header-box h2 { color: white; margin: 0; font-size: 22px; }
    .header-box p { color: #e0e0ff; margin: 5px 0 0 0; font-size: 14px; }
    .metric-box {
        background: white; padding: 15px; border-radius: 8px;
        border: 1px solid #e0e0e0; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .metric-title { font-size: 14px; color: #666; margin-bottom: 5px; }
    .metric-value { font-size: 28px; font-weight: bold; color: #000080; }
    .row-item {
        display: flex; justify-content: space-between;
        padding: 8px 0; border-bottom: 1px dashed #eee;
    }
    .row-item span:last-child { font-weight: bold; color: #000080; }
    .warning-box {
        background: #fff9e6; border-left: 4px solid #ffc107;
        padding: 12px 15px; border-radius: 4px;
        font-size: 13px; color: #664d03;
    }
    .success-box {
        background: #e8f5e9; border-left: 4px solid #4caf50;
        padding: 12px 15px; border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ===== ĐỌC SECRETS =====
def get_secret(key, default=None):
    try:
        return st.secrets[key]
    except:
        return default

# ===== LẤY GEMINI KEY =====
def get_gemini_key():
    return get_secret("GEMINI_API_KEY", None)

# ===== HÀM ĐỌC FILE BẰNG AI (GEMINI) =====
def doc_file_bang_ai(file_bytes, file_name, loai_file):
    """
    Đọc file bất kỳ bằng Gemini AI
    loai_file: "bang_luong" hoặc "d02"
    """
    api_key = get_gemini_key()
    if not api_key:
        return {"error": "Chưa cấu hình GEMINI_API_KEY trong Secrets"}
    
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
    except Exception as e:
        return {"error": f"Lỗi cấu hình Gemini: {e}"}
    
    # Xác định MIME type
    ten_lower = file_name.lower()
    if ten_lower.endswith('.pdf'):
        mime_type = 'application/pdf'
    elif ten_lower.endswith('.xlsx'):
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif ten_lower.endswith('.xls'):
        mime_type = 'application/vnd.ms-excel'
    elif ten_lower.endswith('.docx'):
        mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif ten_lower.endswith('.csv'):
        mime_type = 'text/csv'
    else:
        mime_type = 'application/octet-stream'
    
    # Prompt cho từng loại
    if loai_file == "bang_luong":
        prompt = """Bạn là chuyên gia kế toán Việt Nam. Hãy đọc file bảng lương/thanh toán tiền lương này và trích xuất dữ liệu.

YÊU CẦU QUAN TRỌNG:
1. Tìm cột "Họ và tên" (có thể là: "Họ tên", "Tên nhân viên", "Họ và tên", "Employee Name", "Tên NV")
2. Tìm cột lương TỔNG theo thứ tự ưu tiên:
   - "Tổng tiền lương còn được lĩnh" / "Tổng tiền công được lĩnh"
   - "Tổng lương tháng" / "Tổng thu nhập" / "Tổng cộng"
   - "Cộng" (cột tổng CUỐI CÙNG trong bảng) / "Sub total" / "Total"
   - "Thực lĩnh" / "Net salary" / "Lương thực tế"
3. Bỏ qua các dòng: tiêu đề công ty, dòng trống, dòng "Tổng cộng", "Cộng", "STT"
4. Bỏ qua các cột STT, Số TT
5. CHỈ lấy những dòng có tên người thật (không phải header hay tổng cộng)

QUAN TRỌNG: File có thể có header 1-4 dòng, merge cell. Hãy đọc hiểu ngữ nghĩa.

TRẢ VỀ DUY NHẤT JSON (KHÔNG có markdown, KHÔNG giải thích):
{
  "cot_ho_ten": "tên cột họ tên đã tìm thấy",
  "cot_luong": "tên cột lương đã tìm thấy",
  "so_dong_du_lieu": 45,
  "du_lieu": [
    {"ho_ten": "NGUYỄN VĂN A", "luong": 5000000},
    {"ho_ten": "TRẦN THỊ B", "luong": 6000000}
  ]
}"""
    else:
        prompt = """Bạn là chuyên gia BHXH Việt Nam. Hãy đọc file danh sách tham gia BHXH (mẫu D02-TS) này.

YÊU CẦU QUAN TRỌNG:
1. Tìm cột "Họ và tên" (có thể: "Họ tên", "Họ và tên", "Tên")
2. Tìm cột "Tiền lương tiền công" (hoặc "Tiền lương đóng BHXH", "Tiền lương", "Lương đóng")
3. Tìm cột "Mã số BHXH" (nếu có)
4. Tìm cột "Ngày sinh" (nếu có)
5. Bỏ qua dòng header, dòng tổng cộng, dòng trống

TRẢ VỀ DUY NHẤT JSON (KHÔNG markdown, KHÔNG giải thích):
{
  "cot_ho_ten": "tên cột",
  "cot_luong": "tên cột",
  "cot_bhxh": "tên cột hoặc null",
  "cot_ngay_sinh": "tên cột hoặc null",
  "so_dong_du_lieu": 45,
  "du_lieu": [
    {"ho_ten": "NGUYỄN VĂN A", "luong": 5000000, "ma_bhxh": "0123456789", "ngay_sinh": "01/01/1990"}
  ]
}"""
    
    try:
        response = model.generate_content([
            {"mime_type": mime_type, "data": file_bytes},
            prompt
        ])
        
        text = response.text.strip()
        # Loại bỏ markdown
        text = re.sub(r'^```json\s*', '', text)
        text = re.sub(r'\s*```$', '', text)
        text = text.strip()
        
        # Tìm JSON trong text
        start = text.find('{')
        end = text.rfind('}') + 1
        if start >= 0 and end > start:
            text = text[start:end]
        
        data = json.loads(text)
        return data
    except json.JSONDecodeError as e:
        return {"error": f"AI trả về không đúng JSON: {e}", "raw": text[:500]}
    except Exception as e:
        return {"error": f"Lỗi gọi AI: {e}"}

# ===== CHUYỂN JSON AI THÀNH DATAFRAME =====
def json_to_dataframe(data, loai_file):
    if not data or "du_lieu" not in data:
        return None
    
    df = pd.DataFrame(data["du_lieu"])
    if len(df) == 0:
        return None
    
    if loai_file == "bang_luong":
        df = df.rename(columns={"ho_ten": "Họ tên", "luong": "Lương trả thực tế"})
        if "Lương trả thực tế" not in df.columns:
            return None
        df["Lương trả thực tế"] = pd.to_numeric(df["Lương trả thực tế"], errors='coerce').fillna(0)
    else:
        df = df.rename(columns={"ho_ten": "Họ tên", "luong": "Lương đóng D02",
                                "ma_bhxh": "Mã số BHXH", "ngay_sinh": "Ngày sinh"})
        if "Lương đóng D02" not in df.columns:
            return None
        df["Lương đóng D02"] = pd.to_numeric(df["Lương đóng D02"], errors='coerce').fillna(0)
    
    df = df[df["Họ tên"].notna() & (df["Họ tên"].astype(str).str.strip() != "")]
    df["Họ tên"] = df["Họ tên"].astype(str).str.strip()
    
    return df.reset_index(drop=True)

# ===== ĐỌC FILE KHÔNG DÙNG AI (FALLBACK) =====
@st.cache_data(ttl=1800, show_spinner=False)
def doc_excel_cached(file_bytes, file_name):
    try:
        ten_file = file_name.lower()
        engine = 'xlrd' if ten_file.endswith('.xls') else None
        return pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine, dtype=str)
    except:
        return None

def tim_dong_header(df):
    for i in range(min(15, len(df))):
        row = df.iloc[i]
        row_str = " ".join([str(x) for x in row.values if pd.notna(x) and str(x).strip()]).lower()
        co_ho_ten = "họ và tên" in row_str or "họ tên" in row_str
        co_stt = "số tt" in row_str or "stt" in row_str
        co_luong = "lương cơ bản" in row_str or "basic salary" in row_str
        co_ma_nv = "mã nv" in row_str or "mã nhân viên" in row_str
        if sum([co_ho_ten, co_stt, co_luong, co_ma_nv]) >= 2:
            return i
    return None

# ===== CHUẨN HÓA TÊN =====
def chuan_hoa_ten_series(series):
    s = series.astype(str).str.strip().str.lower()
    s = s.str.replace(r'\s+', ' ', regex=True)
    s = s.str.replace(r'[àáảãạăằắẳẵặâầấẩẫậ]', 'a', regex=True)
    s = s.str.replace(r'[èéẻẽẹêềếểễệ]', 'e', regex=True)
    s = s.str.replace(r'[ìíỉĩị]', 'i', regex=True)
    s = s.str.replace(r'[òóỏõọôồốổỗộơờớởỡợ]', 'o', regex=True)
    s = s.str.replace(r'[ùúủũụưừứửữự]', 'u', regex=True)
    s = s.str.replace(r'[ỳýỷỹỵ]', 'y', regex=True)
    s = s.str.replace(r'[đ]', 'd', regex=True)
    return s.str.strip()

# ===== ĐỐI CHIẾU =====
def doi_chieu(bang_luong, d02):
    bl = bang_luong.copy()
    d2 = d02.copy()
    
    bl["ten_chuan"] = chuan_hoa_ten_series(bl["Họ tên"])
    d2["ten_chuan"] = chuan_hoa_ten_series(d2["Họ tên"])
    
    merged = pd.merge(bl, d2, on="ten_chuan", how="outer",
                      suffixes=("_BL", "_D02"), indicator=True)
    
    merged["Họ tên"] = merged["Họ tên_BL"].fillna(merged["Họ tên_D02"])
    merged["Lương trả thực tế"] = pd.to_numeric(merged.get("Lương trả thực tế", 0), errors='coerce').fillna(0)
    merged["Lương đóng D02"] = pd.to_numeric(merged.get("Lương đóng D02", 0), errors='coerce').fillna(0)
    
    merged["Loại"] = None
    merged.loc[merged["_merge"] == "left_only", "Loại"] = "Truy đóng"
    merged.loc[(merged["_merge"] == "both") & (merged["Lương trả thực tế"] > merged["Lương đóng D02"]), "Loại"] = "Truy thu"
    
    kq = merged[merged["Loại"].notna()].copy()
    if len(kq) == 0:
        return pd.DataFrame()
    
    kq["Chênh lệch"] = kq["Lương trả thực tế"] - kq["Lương đóng D02"]
    kq.loc[kq["Loại"] == "Truy đóng", "Chênh lệch"] = kq.loc[kq["Loại"] == "Truy đóng", "Lương trả thực tế"]
    
    kq["STT"] = range(1, len(kq) + 1)
    kq["Khoản truy thu"] = kq["Loại"].map({
        "Truy đóng": "Chưa có tên trên D02",
        "Truy thu": "Chênh lệch tiền lương đóng BHXH"
    })
    kq["Số tháng"] = 12
    kq["Số tiền truy thu"] = kq["Chênh lệch"] * 0.32
    kq["Diễn giải pháp lý"] = kq["Loại"].map({
        "Truy đóng": "Chưa tham gia BHXH - Đăng ký theo Điều 31 Luật BHXH 2024",
        "Truy thu": "Điều 31 Luật BHXH 2024 - Tiền lương làm căn cứ đóng BHXH"
    })
    
    cols = ["STT", "Họ tên", "Mã số BHXH", "Ngày sinh", "Loại",
            "Lương đóng D02", "Lương trả thực tế", "Chênh lệch",
            "Khoản truy thu", "Số tháng", "Số tiền truy thu", "Diễn giải pháp lý"]
    for c in cols:
        if c not in kq.columns:
            kq[c] = ""
    
    return kq[cols].reset_index(drop=True)

# ===== HEADER =====
st.markdown("""
<div class="header-box">
    <h2>📊 CÔNG CỤ ĐỐI CHIẾU MỨC ĐÓNG BHXH, BHYT, BHTN</h2>
    <p>Hệ thống quản lý tự động - Căn cứ Luật BHXH 2024 (Luật số 41/2024/QH15) | Hỗ trợ AI đọc mọi định dạng file</p>
</div>
""", unsafe_allow_html=True)

# ===== KIỂM TRA API KEY =====
api_key = get_gemini_key()
if not api_key:
    st.warning("⚠️ Chưa cấu hình GEMINI_API_KEY. Vào Settings → Secrets để thêm. App sẽ dùng chế độ đọc Excel thường.")

# ===== CHỌN ĐƠN VỊ =====
danh_sach_don_vi = get_secret("danh_sach_don_vi", ["[TF0372F] CÔNG TY TNHH K&D GARMENT"])

col_a, col_b = st.columns([4, 1])
with col_a:
    don_vi = st.selectbox("🏢 Đơn vị đang chọn:", danh_sach_don_vi)
with col_b:
    st.write("")
    st.write("")
    if st.button("📁 Mở Lưu Trữ", use_container_width=True):
        st.info("Chức năng đang phát triển")

# ===== MENU TABS =====
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Đối chiếu", "📂 Nhập dữ liệu", "📜 Danh sách đơn vị",
    "💼 Mức lương tối thiểu", "📚 Thư viện pháp luật", "⚙️ Quản trị"
])

# ============================================================
# TAB 1: ĐỐI CHIẾU
# ============================================================
with tab1:
    st.markdown("### 📉 Mức Đóng & Tỷ Lệ Trích Nộp Bảo Hiểm Bắt Buộc")
    st.caption("Áp dụng cho HĐLĐ từ 01 tháng trở lên | Mức tham chiếu 2.340.000 VNĐ | Lương tối thiểu Vùng II 4.410.000 VNĐ")
    st.markdown("**Tổng trích nộp: 32%**")
    st.markdown("---")

    col_emp, col_er = st.columns(2)
    with col_emp:
        st.markdown("""
        <div class="metric-box">
            <div class="metric-title">👤 Người Lao Động (NLĐ) Trích Nộp</div>
            <div class="metric-value" style="color:#1976d2;">10.5%</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        st.markdown("""
        <div class="row-item"><span>🟢 BHXH (Hưu trí & Tử tuất)</span><span>8.0%</span></div>
        <div class="row-item"><span>❤️ BHYT</span><span>1.5%</span></div>
        <div class="row-item"><span>🟦 BHTN</span><span>1.0%</span></div>
        """, unsafe_allow_html=True)
    with col_er:
        st.markdown("""
        <div class="metric-box">
            <div class="metric-title">🏢 Người Sử Dụng Lao Động (NSDLĐ) Đóng</div>
            <div class="metric-value" style="color:#388e3c;">21.5%</div>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        st.markdown("""
        <div class="row-item"><span>🟢 BHXH (14% + 3%)</span><span>17.0%</span></div>
        <div class="row-item"><span>❤️ BHYT</span><span>3.0%</span></div>
        <div class="row-item"><span>🟦 BHTN</span><span>1.0%</span></div>
        <div class="row-item"><span>⚠️ BHTN-BNN</span><span>0.5%*</span></div>
        """, unsafe_allow_html=True)

    st.write("")
    st.markdown("""
    <div class="warning-box">
    <b>⚠️ Lưu ý:</b> Mức trần BHXH, BHYT: 46.800.000 VNĐ/tháng | Mức trần BHTN Vùng II: 88.200.000 VNĐ/tháng
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("📋 Kết quả đối chiếu - Danh sách truy thu & truy đóng")

    if "df_ket_qua" in st.session_state and st.session_state.df_ket_qua is not None:
        df = st.session_state.df_ket_qua
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        with col_s1:
            st.metric("Tổng lao động", len(df))
        with col_s2:
            st.metric("🔴 Truy thu", len(df[df["Loại"] == "Truy thu"]))
        with col_s3:
            st.metric("🟠 Truy đóng", len(df[df["Loại"] == "Truy đóng"]))
        with col_s4:
            st.metric("Tổng tiền", f"{df['Số tiền truy thu'].sum():,.0f} VNĐ")
        
        st.dataframe(df, use_container_width=True, height=400)
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 Tải kết quả (CSV)", csv,
                          f"ket_qua_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
        if st.button("🗑️ Xóa kết quả này"):
            st.session_state.df_ket_qua = None
            st.rerun()
    else:
        st.info("Chưa có dữ liệu. Vui lòng nhập ở tab 'Nhập dữ liệu'.")

# ============================================================
# TAB 2: NHẬP DỮ LIỆU
# ============================================================
with tab2:
    st.header("📂 Nhập dữ liệu")
    st.caption("Hỗ trợ AI đọc: Excel (.xlsx, .xls), CSV, Word (.docx), PDF (text + scan)")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.subheader("1️⃣ Bảng lương")
        file_luong = st.file_uploader("Chọn file bảng lương",
            type=["xlsx", "xls", "csv", "pdf", "docx"], key="file_luong")
        if file_luong:
            st.success(f"✅ Đã tải: {file_luong.name}")
    with col_up2:
        st.subheader("2️⃣ Mẫu D02-TS")
        file_d02 = st.file_uploader("Chọn file D02-TS",
            type=["xlsx", "xls", "csv", "pdf", "docx"], key="file_d02")
        if file_d02:
            st.success(f"✅ Đã tải: {file_d02.name}")

    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_ai = st.button("🤖 ĐỐI CHIẾU BẰNG AI (Khuyến nghị)", type="primary", use_container_width=True)
    with col_btn2:
        btn_thuong = st.button("📊 ĐỐI CHIẾU THƯỜNG (Excel)", use_container_width=True)
    
    if btn_ai:
        if not get_gemini_key():
            st.error("❌ Chưa cấu hình GEMINI_API_KEY. Vào Settings → Secrets để thêm.")
            st.stop()
        
        if file_luong and file_d02:
            progress = st.progress(0)
            status = st.empty()
            try:
                status.text("🤖 AI đang đọc bảng lương...")
                progress.progress(30)
                data_luong = doc_file_bang_ai(file_luong.getvalue(), file_luong.name, "bang_luong")
                
                if "error" in data_luong:
                    st.error(f"❌ {data_luong['error']}")
                    if "raw" in data_luong:
                        st.code(data_luong["raw"])
                    st.stop()
                
                bang_luong = json_to_dataframe(data_luong, "bang_luong")
                if bang_luong is None or len(bang_luong) == 0:
                    st.error("❌ AI không đọc được dữ liệu bảng lương")
                    st.stop()
                
                status.text("🤖 AI đang đọc D02-TS...")
                progress.progress(60)
                data_d02 = doc_file_bang_ai(file_d02.getvalue(), file_d02.name, "d02")
                
                if "error" in data_d02:
                    st.error(f"❌ {data_d02['error']}")
                    if "raw" in data_d02:
                        st.code(data_d02["raw"])
                    st.stop()
                
                d02 = json_to_dataframe(data_d02, "d02")
                if d02 is None or len(d02) == 0:
                    st.error("❌ AI không đọc được dữ liệu D02")
                    st.stop()
                
                st.success(f"✅ Bảng lương: {len(bang_luong)} lao động")
                st.success(f"✅ D02-TS: {len(d02)} lao động")
                
                with st.expander("👁️ Xem dữ liệu AI đã đọc"):
                    st.write(f"**Cột Họ tên AI tìm thấy:** `{data_luong.get('cot_ho_ten', 'N/A')}`")
                    st.write(f"**Cột Lương AI tìm thấy:** `{data_luong.get('cot_luong', 'N/A')}`")
                    st.dataframe(bang_luong.head(20), use_container_width=True)
                    st.write(f"**D02 - Cột Họ tên:** `{data_d02.get('cot_ho_ten', 'N/A')}`")
                    st.write(f"**D02 - Cột Lương:** `{data_d02.get('cot_luong', 'N/A')}`")
                    st.dataframe(d02.head(20), use_container_width=True)
                
                status.text("🔍 Đang đối chiếu...")
                progress.progress(90)
                ket_qua = doi_chieu(bang_luong, d02)
                progress.progress(100)
                status.text("✅ Hoàn tất!")
                
                if len(ket_qua) > 0:
                    st.session_state.df_ket_qua = ket_qua
                    st.success(f"✅ Tìm thấy {len(ket_qua)} lao động cần xử lý!")
                    st.balloons()
                else:
                    st.info("✅ Không có lao động nào cần truy thu/truy đóng.")
            except Exception as e:
                st.error(f"Lỗi: {e}")
                st.exception(e)
        else:
            st.warning("⚠️ Vui lòng tải lên cả 2 file.")
    
    if btn_thuong:
        if file_luong and file_d02:
            with st.spinner("Đang đọc file Excel..."):
                try:
                    df_raw_l = doc_excel_cached(file_luong.getvalue(), file_luong.name)
                    if df_raw_l is None:
                        st.error("❌ Không đọc được bảng lương")
                        st.stop()
                    
                    st.info("💡 Chế độ đọc Excel thường chỉ hỗ trợ file .xlsx/.xls có cấu trúc chuẩn")
                    st.warning("⚠️ Nếu file lộn xộn, vui lòng dùng chế độ AI")
                except Exception as e:
                    st.error(f"Lỗi: {e}")

# ============================================================
# TAB 3-6
# ============================================================
with tab3:
    st.header("📜 Danh sách đơn vị đã đối chiếu")
    if "danh_sach_da_doi_chieu" not in st.session_state:
        st.session_state.danh_sach_da_doi_chieu = []
    if st.session_state.danh_sach_da_doi_chieu:
        st.dataframe(pd.DataFrame(st.session_state.danh_sach_da_doi_chieu), use_container_width=True)
        if st.button("🗑️ Xóa toàn bộ"):
            st.session_state.danh_sach_da_doi_chieu = []
            st.rerun()
    else:
        st.info("Chưa có đơn vị nào.")

with tab4:
    st.header("💼 Mức lương tối thiểu vùng")
    st.caption("Căn cứ Nghị định 74/2024/NĐ-CP")
    col_v1, col_v2, col_v3, col_v4 = st.columns(4)
    with col_v1:
        st.number_input("Vùng I (đ/tháng)", value=4960000, step=100000)
    with col_v2:
        st.number_input("Vùng II (đ/tháng)", value=4410000, step=100000)
    with col_v3:
        st.number_input("Vùng III (đ/tháng)", value=3860000, step=100000)
    with col_v4:
        st.number_input("Vùng IV (đ/tháng)", value=3450000, step=100000)

with tab5:
    st.header("📚 Thư viện căn cứ pháp luật")
    with st.expander("📖 Luật BHXH 2024"):
        st.markdown("- **Điều 31**: Tiền lương làm căn cứ đóng BHXH\n- **Điều 32**: Tỷ lệ đóng")
    with st.expander("📖 Nghị định 74/2024/NĐ-CP"):
        st.markdown("- Quy định mức lương tối thiểu vùng")

with tab6:
    st.header("⚙️ Quản trị hệ thống")
    mk = st.text_input("Nhập mật khẩu admin:", type="password")
    mat_khau_admin = get_secret("mat_khau_admin", "admin123")
    if mk == mat_khau_admin:
        st.success("✅ Đăng nhập thành công")
        st.markdown("---")
        st.subheader("🗑️ Xóa dữ liệu")
        st.warning("⚠️ Thao tác này không thể hoàn tác!")
        col_x1, col_x2, col_x3 = st.columns(3)
        with col_x1:
            if st.button("🗑️ Xóa kết quả"):
                st.session_state.df_ket_qua = None
                st.success("Đã xóa!")
                st.rerun()
        with col_x2:
            if st.button("🗑️ Xóa danh sách đơn vị"):
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa!")
                st.rerun()
        with col_x3:
            if st.button("🗑️ XÓA TOÀN BỘ"):
                st.session_state.df_ket_qua = None
                st.session_state.danh_sach_da_doi_chieu = []
                st.success("Đã xóa!")
                st.rerun()
    elif mk:
        st.error("❌ Sai mật khẩu")

st.markdown("---")
st.caption("📌 Công cụ hỗ trợ đối chiếu tự động dựa trên Luật BHXH 2024.")