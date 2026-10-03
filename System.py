# -*- coding: utf-8 -*-
import datetime
import io
import pandas as pd
import psycopg2
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG & CSS
# ==========================================
st.set_page_config(
    page_title="QC Operations Hub",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    
    [data-testid="metric-container"] {
        background: #ffffff;
        border: 1px solid #e0e6ed;
        border-radius: 8px;
        padding: 1.2rem !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    .stButton button[kind="primary"], .stDownloadButton button {
        background-color: #1a3a6b !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1a3a6b 0%, #2a5298 100%);
        padding: 20px 30px;
        border-radius: 8px;
        margin-bottom: 25px;
        color: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 { color: white; margin: 0; font-size: 24px; font-weight: 700; }
    .main-header p { color: #93b4da; margin: 5px 0 0 0; font-size: 14px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CÁC HÀM TIỆN ÍCH CƠ BẢN
# ==========================================
DB_URL = "postgresql://postgres.hpjxaxspjgsnsoxhvskm:07736215400394219723@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"

def ket_noi():
    try:
        return psycopg2.connect(DB_URL)
    except Exception as e:
        st.error(f"Lỗi kết nối CSDL: {e}")
        st.stop()

def _chuan_hoa_ma(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.replace(r"\s+", "", regex=True).str.upper()
    return s.replace({"NAN": "", "NONE": ""})

def divider_label(text):
    st.markdown(
        f'<div style="font-size:13px;font-weight:bold;color:#1a3a6b;'
        f'text-transform:uppercase;border-bottom:2px solid #1a3a6b;padding-bottom:5px;margin:25px 0 15px;">'
        f'{text}</div>',
        unsafe_allow_html=True,
    )

@st.cache_data(ttl=300)
def xuat_excel(df_dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        for ten, d in df_dict.items():
            d.to_excel(wr, sheet_name=ten, index=False)
    return buf.getvalue()

# ==========================================
# 3. LOGIC TRUY VẤN DỮ LIỆU THÔ (ĐÃ FIX CHUẨN XÁC NGUỒN VÀ CHECK TRÙNG)
# ==========================================
@st.cache_data(ttl=300, show_spinner="Đang truy vấn dữ liệu từ Supabase...")
def truy_van_du_lieu_tho(hub_chon, loai_chon, tu, den, tim):
    df_list = []
    safe_date_ib = """(CASE WHEN "Ngày vận hành" LIKE '%%/%%' THEN TO_DATE("Ngày vận hành", 'MM/DD/YYYY') ELSE "Ngày vận hành"::date END)"""

    # A. Truy vấn IB từ kpi_base
    if not loai_chon or "Dỡ xuống xe" in loai_chon:
        sql_ib = f"""
            SELECT "Mã vận đơn", "Thời gian quét", "Bưu cục quét" AS "Hub", 
                   "Trọng lượng", "Ngày vận hành", 'Dỡ xuống xe' AS "Loại quét"
            FROM kpi_base
            WHERE 1=1
        """
        dk_ib, p_ib = [], []
        if hub_chon:
            dk_ib.append(f'"Bưu cục quét" IN ({",".join(["%s"] * len(hub_chon))})')
            p_ib += list(hub_chon)
        if tu:
            dk_ib.append(f"{safe_date_ib} >= %s")
            p_ib.append(tu)
        if den:
            dk_ib.append(f"{safe_date_ib} <= %s")
            p_ib.append(den)
        if tim:
            dk_ib.append(f'"Mã vận đơn" LIKE %s')
            p_ib.append(f"%{tim}%")
        if dk_ib:
            sql_ib += " AND " + " AND ".join(dk_ib)
            
        con = ket_noi()
        df_ib = pd.read_sql(sql_ib, con, params=p_ib)
        con.close()
        df_list.append(df_ib)

    # B. Truy vấn OB từ raw_quet_hang_xep_len_xe (Dùng Thời gian quét để lọc ngày)
    if not loai_chon or "Xếp lên xe" in loai_chon:
        sql_ob = """
            SELECT "Mã vận đơn", "Thời gian quét", "Bưu cục quét" AS "Hub", 
                   "Trọng lượng", SUBSTRING("Thời gian quét", 1, 10) AS "Ngày vận hành", 
                   'Xếp lên xe' AS "Loại quét"
            FROM raw_quet_hang_xep_len_xe
            WHERE 1=1
        """
        dk_ob, p_ob = [], []
        if hub_chon:
            dk_ob.append(f'"Bưu cục quét" IN ({",".join(["%s"] * len(hub_chon))})')
            p_ob += list(hub_chon)
        if tu:
            dk_ob.append('SUBSTRING("Thời gian quét", 1, 10) >= %s')
            p_ob.append(tu)
        if den:
            dk_ob.append('SUBSTRING("Thời gian quét", 1, 10) <= %s')
            p_ob.append(den)
        if tim:
            dk_ob.append(f'"Mã vận đơn" LIKE %s')
            p_ob.append(f"%{tim}%")
        if dk_ob:
            sql_ob += " AND " + " AND ".join(dk_ob)
            
        con = ket_noi()
        df_ob = pd.read_sql(sql_ob, con, params=p_ob)
        con.close()
        df_list.append(df_ob)

    if not df_list:
        return pd.DataFrame()
        
    df = pd.concat(df_list, ignore_index=True)
    
    if not df.empty:
        df["Mã chuẩn"] = _chuan_hoa_ma(df["Mã vận đơn"])
        danh_sach_mvd = list(set(df["Mã chuẩn"].tolist()) - {"", "\n"})

        if danh_sach_mvd:
            df_hist_list = []
            chunk_size = 5000
            con2 = ket_noi()
            for i in range(0, len(danh_sach_mvd), chunk_size):
                chunk = tuple(danh_sach_mvd[i : i + chunk_size])
                
                # Gom lịch sử đầy đủ từ cả 2 bảng để tính đợt và loại trừ trùng chuẩn xác
                sql_hist = """
                    SELECT "Mã vận đơn", "Thời gian quét", 'Dỡ xuống xe' AS "Loại quét"
                    FROM kpi_base WHERE "Mã vận đơn" IN %s
                    UNION ALL
                    SELECT "Mã vận đơn", "Thời gian quét", 'Xếp lên xe' AS "Loại quét"
                    FROM raw_quet_hang_xep_len_xe WHERE "Mã vận đơn" IN %s
                """
                df_hist_list.append(pd.read_sql(sql_hist, con2, params=(chunk, chunk)))
            con2.close()

            df_hist = pd.concat(df_hist_list, ignore_index=True)
            df_hist["Mã chuẩn"] = _chuan_hoa_ma(df_hist["Mã vận đơn"])
            df_hist["Thời gian quét_dt"] = pd.to_datetime(df_hist["Thời gian quét"], errors="coerce")
            
            # Sắp xếp lịch sử toàn cục theo thời gian tăng dần
            df_hist = df_hist.sort_values(by=["Mã chuẩn", "Thời gian quét_dt"])

            # Đánh số đợt dựa vào dỡ xuống xe
            la_do_xuong = (df_hist["Loại quét"] == "Dỡ xuống xe").astype(int)
            df_hist["Đợt"] = la_do_xuong.groupby(df_hist["Mã chuẩn"]).cumsum()
            df_hist["Mã nối OB"] = df_hist["Mã chuẩn"] + "_Đợt" + df_hist["Đợt"].astype(str)

            # Check trùng OB trong lịch sử
            la_xep_len = df_hist["Loại quét"] == "Xếp lên xe"
            df_hist["Trùng OB (sai thao tác)"] = False
            df_hist.loc[la_xep_len, "Trùng OB (sai thao tác)"] = df_hist[la_xep_len].duplicated(subset=["Mã nối OB"], keep="first")

            df_hist["Key_map"] = df_hist["Mã chuẩn"] + "_" + df_hist["Thời gian quét_dt"].dt.strftime('%Y%m%d%H%M%S')
            
            df["Thời gian quét_dt"] = pd.to_datetime(df["Thời gian quét"], errors="coerce")
            df["Key_map"] = df["Mã chuẩn"] + "_" + df["Thời gian quét_dt"].dt.strftime('%Y%m%d%H%M%S')

            df_hist_unique = df_hist.drop_duplicates(subset=["Key_map"], keep="last")
            map_ma_noi = df_hist_unique.set_index("Key_map")["Mã nối OB"].to_dict()
            map_trung = df_hist_unique.set_index("Key_map")["Trùng OB (sai thao tác)"].to_dict()

            df["Mã nối OB"] = df["Key_map"].map(map_ma_noi)
            df["Trùng OB (sai thao tác)"] = df["Key_map"].map(map_trung).fillna(False)
            
            df = df.drop(columns=["Mã chuẩn", "Thời gian quét_dt", "Key_map"])
            df = df.sort_values(by="Thời gian quét", ascending=False)

    return df

# ==========================================
# 4. GIAO DIỆN TỪNG TRANG
# ==========================================
def page_du_lieu_tho():
    st.markdown('<div class="main-header"><h1>TRUY VẤN QUÉT HÀNG</h1><p>Dữ liệu IB lấy từ kpi_base | Dữ liệu OB lấy từ raw_quet_hang_xep_len_xe</p></div>', unsafe_allow_html=True)
    
    with st.expander("BỘ LỌC TÌM KIẾM", expanded=True):
        with st.form("bo_loc"):
            c1, c2, c3 = st.columns(3)
            hub_chon = c1.multiselect("Hub (Bưu cục quét)", ["HCM HUB", "BN HUB", "SH DC", "CTO SC"], placeholder="Tất cả")
            loai_chon = c2.multiselect("Loại quét", ["Dỡ xuống xe", "Xếp lên xe"], placeholder="Tất cả (IB & OB)")
            tim = c3.text_input("Tìm mã vận đơn (Tuỳ chọn)").strip()
            
            c4, c5 = st.columns(2)
            tu = c4.date_input("Từ ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            den = c5.date_input("Đến ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            
            submit = st.form_submit_button("TRUY VẤN DỮ LIỆU", type="primary", use_container_width=True)

    if submit:
        st.session_state["kq_tho"] = truy_van_du_lieu_tho(
            tuple(hub_chon), tuple(loai_chon), str(tu), str(den), tim
        )

    if "kq_tho" in st.session_state:
        df = st.session_state["kq_tho"]
        divider_label("Kết quả truy vấn")

        if not df.empty:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Số dòng", f"{len(df):,}")
            
            if "Mã nối OB" in df.columns:
                la_ob_dem = df["Loại quét"] == "Xếp lên xe"
                khoa_dem = df["Mã vận đơn"].astype(str).where(~la_ob_dem, df["Mã nối OB"])
                so_van_don = khoa_dem.nunique()
            else:
                so_van_don = df["Mã vận đơn"].nunique()
            m2.metric("Số vận đơn", f"{so_van_don:,}")

            df["Trọng lượng"] = pd.to_numeric(df["Trọng lượng"], errors="coerce").fillna(0)
            mask_ob_w = df["Loại quét"] == "Xếp lên xe"
            mask_ib_w = df["Loại quét"] == "Dỡ xuống xe"

            if "Mã nối OB" in df.columns:
                trong_luong_ob = df[mask_ob_w].drop_duplicates(subset="Mã nối OB")["Trọng lượng"].sum()
            else:
                trong_luong_ob = df.loc[mask_ob_w, "Trọng lượng"].sum()
            trong_luong_ib = df[mask_ib_w].drop_duplicates(subset="Mã vận đơn")["Trọng lượng"].sum()
            
            m3.metric("Trọng lượng (kg)", f"{(trong_luong_ob + trong_luong_ib):,.0f}")
            
            so_trung_ob = df["Trùng OB (sai thao tác)"].sum() if "Trùng OB (sai thao tác)" in df.columns else 0
            m4.metric("Trùng OB (sai thao tác)", f"{so_trung_ob:,}")

            st.dataframe(df.head(500), use_container_width=True, height=400)

            if "Mã nối OB" in df.columns:
                la_ob_xuat = df["Loại quét"] == "Xếp lên xe"
                khoa_xuat = df["Mã vận đơn"].astype(str).where(~la_ob_xuat, df["Mã nối OB"])
                df_xuat = df.loc[~khoa_xuat.duplicated(keep="first")]
            else:
                df_xuat = df.drop_duplicates(subset="Mã vận đơn")
                
            cot_tai, _ = st.columns(2)
            with cot_tai:
                st.download_button(
                    f"Tải Excel ({len(df_xuat):,} số vận đơn sau khi lọc)",
                    xuat_excel({"Du lieu": df_xuat}),
                    file_name=f"quet_hang_{datetime.date.today()}.xlsx",
                    use_container_width=True,
                    type="primary",
                )
        else:
            st.warning("Không tìm thấy dữ liệu phù hợp với bộ lọc!")

def page_ontime_xep_xe():
    st.title("Đang chờ update logic...")
    st.info("Test xong trang Dữ liệu thô thì nhắn mình để update phần này nhé!")

def page_bao_cao_ontime():
    st.title("Đang chờ update logic...")
    st.info("Test xong trang Dữ liệu thô thì nhắn mình để update phần này nhé!")

# ==========================================
# 5. ĐIỀU HƯỚNG CHÍNH (SIDEBAR)
# ==========================================
def main():
    with st.sidebar:
        st.markdown("## VẬN HÀNH QC")
        st.markdown("---")
        menu_lua_chon = st.radio(
            "ĐIỀU HƯỚNG:",
            ["Dữ liệu thô", "Ontime Xếp xe", "Báo cáo Ontime"],
            label_visibility="collapsed"
        )
        st.markdown("---")
        st.caption(f"Hôm nay: {datetime.date.today().strftime('%d/%m/%Y')}")

    if menu_lua_chon == "Dữ liệu thô":
        page_du_lieu_tho()
    elif menu_lua_chon == "Ontime Xếp xe":
        page_ontime_xep_xe()
    elif menu_lua_chon == "Báo cáo Ontime":
        page_bao_cao_ontime()

if __name__ == "__main__":
    main()
