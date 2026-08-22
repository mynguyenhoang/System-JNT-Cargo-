# -*- coding: utf-8 -*-
"""
QC OPERATIONS DASHBOARD - HOAN TOAN MOI (SUPABASE + STREAMLIT)
"""
import datetime
import io
import pandas as pd
import psycopg2
import streamlit as st

# 1. Cấu hình giao diện trang
st.set_page_config(
    page_title="QC Operations Hub",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS tùy chỉnh màu sắc, giao diện cho đẹp mắt
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1a3a6b 0%, #2c5282 100%);
        padding: 18px 24px;
        border-radius: 8px;
        color: white;
        margin-bottom: 20px;
    }
    .main-header h1 {
        font-size: 22px;
        margin: 0;
        font-weight: 700;
    }
    .main-header p {
        font-size: 13px;
        margin: 4px 0 0 0;
        color: #cbd5e0;
    }
</style>
""", unsafe_allow_html=True)

# 2. Lấy cấu hình kết nối Supabase từ Streamlit Secrets
DB_URL = st.secrets.get("DB_URL", "")
if not DB_URL:
    st.error("Chưa khai báo `DB_URL` trong Streamlit Secrets. Vui lòng cấu hình lại.")
    st.stop()

def get_connection():
    return psycopg2.connect(DB_URL)

# Tiêu đề ứng dụng
st.markdown("""
<div class="main-header">
    <h1>HỆ THỐNG VẬN HÀNH QC & SUPABASE</h1>
    <p>Giao diện tra cứu báo cáo Ontime, dữ liệu quét thô và tiến độ xếp hàng thời gian thực</p>
</div>
""", unsafe_allow_html=True)

# Hàm chuyển đổi DataFrame sang Excel để người dùng tải về
def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# 3. Thanh điều hướng Sidebar
menu = st.sidebar.selectbox(
    "Chọn danh mục",
    ["📊 Báo cáo Tổng hợp (Ontime/Late)", "📦 Tra cứu Quét Thô", "🚛 Tiến độ Xếp Hàng", "⚙️ Quản lý & Cào Dữ Liệu"]
)

# =========================================================================
# TAB 1: BÁO CÁO TỔNG HỢP (Đọc từ bảng `bao_cao`)
# =========================================================================
if menu == "📊 Báo cáo Tổng hợp (Ontime/Late)":
    st.subheader("📋 Báo cáo Nghiệp vụ Tổng Hợp (Ontime, Late, Backlog, Giải trình)")
    
    with st.form("filter_baocao"):
        col1, col2, col3 = st.columns(3)
        with col1:
            hub_filter = st.selectbox("Lọc Hub", ["Tất cả", "BN HUB", "HCM HUB", "SH DC"])
        with col2:
            status_filter = st.selectbox("Trạng thái xử lý", ["Tất cả", "Ontime", "Late", "Backlog", "Thiếu cutoff"])
        with col3:
            limit_rows = st.number_input("Giới hạn số dòng hiển thị", min_value=100, max_value=20000, value=1000, step=100)
            
        submitted = st.form_submit_button("Tải dữ liệu báo cáo", type="primary", use_container_width=True)

    if submitted or "df_bc_loaded" not in st.session_state:
        try:
            conn = get_connection()
            query = 'SELECT * FROM bao_cao WHERE 1=1'
            params = []
            
            if hub_filter != "Tất cả":
                query += ' AND "Hub" = %s'
                params.append(hub_filter)
            if status_filter != "Tất cả":
                query += ' AND "Trạng thái xử lý" = %s'
                params.append(status_filter)
                
            query += ' ORDER BY "Ngày vận hành" DESC LIMIT %s'
            params.append(limit_rows)
            
            df_bc = pd.read_sql(query, conn, params=params)
            conn.close()
            st.session_state["df_bc_loaded"] = df_bc
        except Exception as e:
            st.error(f"Lỗi truy vấn cơ sở dữ liệu: {e}")
            st.session_state["df_bc_loaded"] = pd.DataFrame()

    df_bc = st.session_state.get("df_bc_loaded", pd.DataFrame())

    if not df_bc.empty:
        # Hiển thị số liệu tổng quan metrics phía trên
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Tổng số đơn", f"{len(df_bc):,}")
        if "Trạng thái xử lý" in df_bc.columns:
            ontime_cnt = len(df_bc[df_bc["Trạng thái xử lý"] == "Ontime"])
            late_cnt = len(df_bc[df_bc["Trạng thái xử lý"] == "Late"])
            m2.metric("Số lượng Ontime", f"{ontime_cnt:,}")
            m3.metric("Số lượng Late", f"{late_cnt:,}")
            ratio = (ontime_cnt / len(df_bc) * 100) if len(df_bc) > 0 else 0
            m4.metric("Tỷ lệ Ontime", f"{ratio:.2f}%")

        st.dataframe(df_bc, use_container_width=True, height=450)
        
        st.download_button(
            label="📥 Tải xuống file Excel báo cáo",
            data=convert_df_to_excel(df_bc),
            file_name=f"bao_cao_tong_hop_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("Không tìm thấy dữ liệu báo cáo phù hợp hoặc bảng đang trống.")

# =========================================================================
# TAB 2: TRA CỨU QUÉT THÔ (Đọc từ bảng `quet_hang`)
# =========================================================================
elif menu == "📦 Tra cứu Quét Thô":
    st.subheader("🔍 Tra cứu dữ liệu Quét Thô (Lên xe / Xuống xe)")
    
    with st.form("filter_quet"):
        c1, c2, c3 = st.columns(3)
        with c1:
            hub_q = st.selectbox("Hub", ["Tất cả", "BN", "HCM", "SHDC"])
        with c2:
            loai_q = st.selectbox("Loại quét", ["Tất cả", "Dỡ xuống xe", "Xếp lên xe"])
        with c3:
            mvd = st.text_input("Mã vận đơn (nhập chính xác hoặc một phần)").strip()
            
        sub_quet = st.form_submit_button("Truy vấn quét hàng", type="primary", use_container_width=True)

    if sub_quet:
        try:
            conn = get_connection()
            query = 'SELECT * FROM quet_hang WHERE 1=1'
            params = []
            
            if hub_q != "Tất cả":
                query += ' AND "Hub" = %s'
                params.append(hub_q)
            if loai_q != "Tất cả":
                query += ' AND "Loại quét" = %s'
                params.append(loai_q)
            if mvd:
                query += ' AND "Mã vận đơn" LIKE %s'
                params.append(f"%{mvd}%")
                
            query += ' ORDER BY "Thời gian quét" DESC LIMIT 1000'
            df_q = pd.read_sql(query, conn, params=params)
            conn.close()
            
            st.success(f"Đã tìm thấy {len(df_q):,} bản ghi.")
            st.dataframe(df_q, use_container_width=True, height=450)
            
            if not df_q.empty:
                st.download_button("📥 Tải Excel dữ liệu quét", data=convert_df_to_excel(df_q), file_name="quet_hang.xlsx")
        except Exception as e:
            st.error(f"Lỗi truy vấn: {e}")

# =========================================================================
# TAB 3: TIẾN ĐỘ XẾP HÀNG (Đọc từ bảng `xep_hang`)
# =========================================================================
elif menu == "🚛 Tiến độ Xếp Hàng":
    st.subheader("🚛 Tiến độ Xếp Hàng (HCM Hub)")
    
    try:
        conn = get_connection()
        df_xh = pd.read_sql('SELECT * FROM xep_hang ORDER BY "Thời gian bắt đầu xếp hàng" DESC LIMIT 1000', conn)
        conn.close()
        
        if not df_xh.empty:
            st.metric("Tổng nhiệm vụ xếp hàng gần đây", f"{len(df_xh):,}")
            st.dataframe(df_xh, use_container_width=True, height=500)
            st.download_button("📥 Tải Excel Tiến độ xếp hàng", data=convert_df_to_excel(df_xh), file_name="xep_hang.xlsx")
        else:
            st.info("Chưa có dữ liệu xếp hàng trên Supabase.")
    except Exception as e:
        st.error(f"Lỗi truy vấn: {e}")

# =========================================================================
# TAB 4: QUẢN LÝ & CÀO DỮ LIỆU TRỰC TIẾP
# =========================================================================
elif menu == "⚙️ Quản lý & Cào Dữ Liệu":
    st.subheader("⚙️ Điều khiển Pipeline Cào Dữ Liệu")
    st.info("Tính năng này giúp bạn kích hoạt trực tiếp script cào dữ liệu từ J&T đẩy vào Supabase ngay trên giao diện web.")
    
    if st.button("🚀 Chạy Pipeline Hàng Ngày Ngay Bây Giờ", type="primary", use_container_width=True):
        try:
            from jfs_pipeline import run_full_pipeline
            with st.spinner("Đang kết nối API J&T và đồng bộ dữ liệu, vui lòng không tắt trang..."):
                run_full_pipeline()
            st.success("Tuyệt vời! Đã chạy xong pipeline và cập nhật dữ liệu thành công lên Supabase.")
        except Exception as e:
            st.error(f"Lỗi thực thi pipeline: {e}")
