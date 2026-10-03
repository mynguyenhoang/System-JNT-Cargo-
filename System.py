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
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. CÁC HÀM TIỆN ÍCH CƠ BẢN
# ==========================================
DB_URL = "postgresql://postgres.hpjxaxspjgsnsoxhvskm:07736215400394219723@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres?sslmode=require"

# Trang "Ontime 准时报表": các Bưu cục (cột "Bưu cục") bị loại khỏi tổng IB và COT.
# So khớp không phân biệt hoa/thường và bỏ khoảng trắng ("DT TN" = "DTTN").
BUU_CUC_LOAI_TRU = {"DTTN", "SETN"}


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

def _khoa_buu_cuc(series):
    """Chuẩn hóa tên bưu cục để so khớp: IN HOA + bỏ mọi khoảng trắng."""
    return series.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

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
# 3. TRANG 1: SẢN LƯỢNG | 生产
# ==========================================
@st.cache_data(ttl=300, show_spinner="Đang truy vấn dữ liệu sản lượng...")
def truy_van_du_lieu_tho(hub_chon, loai_chon, tu, den, tim):
    df_list = []
    safe_date = """(CASE WHEN "Ngày vận hành" LIKE '%%/%%' THEN TO_DATE("Ngày vận hành", 'MM/DD/YYYY') ELSE "Ngày vận hành"::date END)"""

    if not loai_chon or "Dỡ xuống xe" in loai_chon:
        sql_ib = f"""
            SELECT "Mã vận đơn", "Thời gian quét", "Bưu cục quét" AS "Hub", 
                   "Trọng lượng", "Ngày vận hành", 'Dỡ xuống xe' AS "Loại quét"
            FROM kpi_base WHERE 1=1
        """
        dk_ib, p_ib = [], []
        if hub_chon:
            dk_ib.append(f'"Bưu cục quét" IN ({",".join(["%s"] * len(hub_chon))})')
            p_ib += list(hub_chon)
        if tu:
            dk_ib.append(f"{safe_date} >= %s")
            p_ib.append(tu)
        if den:
            dk_ib.append(f"{safe_date} <= %s")
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

    if not loai_chon or "Xếp lên xe" in loai_chon:
        sql_ob = f"""
            SELECT "Mã vận đơn", "Thời gian quét", "Bưu cục quét" AS "Hub", 
                   "Trọng lượng", "Ngày vận hành", 'Xếp lên xe' AS "Loại quét"
            FROM raw_quet_hang_xep_len_xe WHERE 1=1
        """
        dk_ob, p_ob = [], []
        if hub_chon:
            dk_ob.append(f'"Bưu cục quét" IN ({",".join(["%s"] * len(hub_chon))})')
            p_ob += list(hub_chon)
        if tu:
            dk_ob.append(f"{safe_date} >= %s")
            p_ob.append(tu)
        if den:
            dk_ob.append(f"{safe_date} <= %s")
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
            df_hist = df_hist.sort_values(by=["Mã chuẩn", "Thời gian quét_dt"])

            la_do_xuong = (df_hist["Loại quét"] == "Dỡ xuống xe").astype(int)
            df_hist["Đợt"] = la_do_xuong.groupby(df_hist["Mã chuẩn"]).cumsum()
            df_hist["Mã nối OB"] = df_hist["Mã chuẩn"] + "_Đợt" + df_hist["Đợt"].astype(str)

            df_hist["Key_map"] = df_hist["Mã chuẩn"] + "_" + df_hist["Thời gian quét_dt"].dt.strftime('%Y%m%d%H%M%S')
            df["Thời gian quét_dt"] = pd.to_datetime(df["Thời gian quét"], errors="coerce")
            df["Key_map"] = df["Mã chuẩn"] + "_" + df["Thời gian quét_dt"].dt.strftime('%Y%m%d%H%M%S')

            df_hist_unique = df_hist.drop_duplicates(subset=["Key_map"], keep="last")
            map_ma_noi = df_hist_unique.set_index("Key_map")["Mã nối OB"].to_dict()

            df["Mã nối OB"] = df["Key_map"].map(map_ma_noi)
            df = df.drop(columns=["Mã chuẩn", "Thời gian quét_dt", "Key_map"], errors="ignore")
            df = df.sort_values(by="Thời gian quét", ascending=False)

    return df

def page_du_lieu_tho():
    st.markdown('<div class="main-header"><h1>SẢN LƯỢNG | 生产</h1></div>', unsafe_allow_html=True)
    
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
            m1, m2, m3, _ = st.columns(4)
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


# ==========================================
# 4. TRANG 2: Linehaul Ontime Departure / 干线准时发车
# ==========================================
@st.cache_data(ttl=300, show_spinner="Đang truy vấn tiến độ xếp xe...")
def truy_van_ontime_xep_xe(hub_chon, tu, den):
    safe_date_check = """
        CASE 
            WHEN "Date" ~ '^[0-9]{2}/[0-9]{2}/[0-9]{4}$' THEN TO_DATE("Date", 'MM/DD/YYYY')
            WHEN "Date" ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN "Date"::date
            ELSE NULL
        END
    """
    sql = f"""
        SELECT * FROM raw_quan_ly_tien_do_xep_hang
        WHERE ("Date" ~ '^[0-9]{{2}}/[0-9]{{2}}/[0-9]{{4}}$' OR "Date" ~ '^[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}$')
    """
    dk, params = [], []
    if hub_chon:
        dk.append(f'"Bộ phận xếp hàng" IN ({",".join(["%s"] * len(hub_chon))})')
        params += list(hub_chon)
    if tu:
        dk.append(f"({safe_date_check}) >= %s")
        params.append(tu)
    if den:
        dk.append(f"({safe_date_check}) <= %s")
        params.append(den)
        
    if dk:
        sql += " AND " + " AND ".join(dk)
        
    con = ket_noi()
    df = pd.read_sql(sql, con, params=params)
    con.close()
    return df

def page_ontime_xep_xe():
    st.markdown('<div class="main-header"><h1>Linehaul Ontime Departure / 干线准时发车</h1></div>', unsafe_allow_html=True)
    
    with st.expander("BỘ LỌC TÌM KIẾM", expanded=True):
        with st.form("form_ontime_xh"):
            c1, c2, c3 = st.columns(3)
            hub_xh = c1.multiselect("Bộ phận xếp hàng (Hub)", ["HCM HUB", "BN HUB", "SH DC", "CTO SC"], placeholder="Tất cả")
            tu_xh = c2.date_input("Từ ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            den_xh = c3.date_input("Đến ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            
            submit_xh = st.form_submit_button("TRUY VẤN TIẾN ĐỘ", type="primary", use_container_width=True)

    if submit_xh:
        st.session_state["kq_ontime_xh"] = truy_van_ontime_xep_xe(
            tuple(hub_xh), str(tu_xh), str(den_xh)
        )

    if "kq_ontime_xh" in st.session_state:
        df_xh = st.session_state["kq_ontime_xh"]
        divider_label("Kết quả tiến độ chuyến xe")

        if not df_xh.empty:
            df_xh["type_clean"] = df_xh["type"].astype(str).str.strip().str.upper()
            df_xh["ontime_clean"] = df_xh["ontime"].astype(str).str.strip().str.lower()

            df_lh = df_xh[df_xh["type_clean"] == "LINEHAUL"]
            df_st = df_xh[df_xh["type_clean"] == "SHUTTLE"]

            tab_lh, tab_st = st.tabs(["LINEHAUL", "SHUTTLE"])

            with tab_lh:
                st.markdown("### Thống kê tuyến Linehaul")
                so_ontime_lh = (df_lh["ontime_clean"] == "ontime").sum()
                so_late_lh = (df_lh["ontime_clean"] == "late").sum()
                ty_le_lh = (so_ontime_lh / len(df_lh) * 100) if len(df_lh) > 0 else 0

                l1, l2, l3, l4 = st.columns(4)
                l1.metric("Tổng chuyến Linehaul", f"{len(df_lh):,}")
                l2.metric("Ontime", f"{so_ontime_lh:,}")
                l3.metric("Late", f"{so_late_lh:,}")
                l4.metric("Tỷ lệ Ontime", f"{ty_le_lh:.2f}%")

                st.dataframe(df_lh.head(500), use_container_width=True, height=380)

            with tab_st:
                st.markdown("### Thống kê tuyến Shuttle")
                so_ontime_st = (df_st["ontime_clean"] == "ontime").sum()
                so_late_st = (df_st["ontime_clean"] == "late").sum()
                ty_le_st = (so_ontime_st / len(df_st) * 100) if len(df_st) > 0 else 0

                s1, s2, s3, s4 = st.columns(4)
                s1.metric("Tổng chuyến Shuttle", f"{len(df_st):,}")
                s2.metric("Ontime", f"{so_ontime_st:,}")
                s3.metric("Late", f"{so_late_st:,}")
                s4.metric("Tỷ lệ Ontime", f"{ty_le_st:.2f}%")

                st.dataframe(df_st.head(500), use_container_width=True, height=380)

            divider_label("Xuất dữ liệu tổng hợp")
            cot_tai, _ = st.columns(2)
            with cot_tai:
                st.download_button(
                    f"Tải Excel Tổng Hợp ({len(df_xh):,} chuyến xe)",
                    xuat_excel({
                        "Linehaul": df_lh, 
                        "Shuttle": df_st, 
                        "Tat ca": df_xh
                    }),
                    file_name=f"ontime_xep_xe_{datetime.date.today()}.xlsx",
                    use_container_width=True,
                    type="primary",
                )
        else:
            st.warning("Không tìm thấy dữ liệu phù hợp với bộ lọc!")


# ==========================================
# 5. TRANG 3: Ontime 准时报表
# ==========================================
@st.cache_data(ttl=300, show_spinner="Đang truy vấn báo cáo Ontime...")
def truy_van_bao_cao_ontime(hub_chon, tu, den):
    safe_date_check = """
        CASE 
            WHEN "Ngày vận hành" ~ '^[0-9]{2}/[0-9]{2}/[0-9]{4}$' THEN TO_DATE("Ngày vận hành", 'MM/DD/YYYY')
            WHEN "Ngày vận hành" ~ '^[0-9]{4}-[0-9]{2}-[0-9]{2}$' THEN "Ngày vận hành"::date
            ELSE NULL
        END
    """
    sql = f"""
        SELECT * FROM kpi_base
        WHERE ("Ngày vận hành" ~ '^[0-9]{{2}}/[0-9]{{2}}/[0-9]{{4}}$' OR "Ngày vận hành" ~ '^[0-9]{{4}}-[0-9]{{2}}-[0-9]{{2}}$')
    """
    dk, params = [], []
    if hub_chon:
        dk.append(f'"Bưu cục quét" IN ({",".join(["%s"] * len(hub_chon))})')
        params += list(hub_chon)
    if tu:
        dk.append(f"({safe_date_check}) >= %s")
        params.append(tu)
    if den:
        dk.append(f"({safe_date_check}) <= %s")
        params.append(den)
        
    if dk:
        sql += " AND " + " AND ".join(dk)
        
    con = ket_noi()
    df = pd.read_sql(sql, con, params=params)
    con.close()
    return df

def page_bao_cao_ontime():
    st.markdown('<div class="main-header"><h1>Ontime 准时报表</h1></div>', unsafe_allow_html=True)
    
    with st.expander("BỘ LỌC TÌM KIẾM", expanded=True):
        with st.form("form_bao_cao_ontime"):
            c1, c2, c3 = st.columns(3)
            hub_bc = c1.multiselect("Bưu cục quét (Hub)", ["HCM HUB", "BN HUB", "SH DC", "CTO SC"], placeholder="Tất cả")
            tu_bc = c2.date_input("Từ ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            den_bc = c3.date_input("Đến ngày vận hành", value=datetime.date.today(), format="YYYY-MM-DD")
            
            submit_bc = st.form_submit_button("TRUY VẤN BÁO CÁO", type="primary", use_container_width=True)

    if submit_bc:
        st.session_state["kq_bao_cao_ontime"] = truy_van_bao_cao_ontime(
            tuple(hub_bc), str(tu_bc), str(den_bc)
        )

    if "kq_bao_cao_ontime" in st.session_state:
        df_bc_goc = st.session_state["kq_bao_cao_ontime"]
        divider_label("Kết quả báo cáo tổng hợp")

        if df_bc_goc.empty:
            st.warning("Không tìm thấy dữ liệu phù hợp với bộ lọc!")
            return

        # Loại trừ các đơn có Bưu cục (đích) thuộc DT TN / SETN.
        # Chỉ áp dụng ở trang này; trang SẢN LƯỢNG vẫn giữ nguyên toàn bộ số IB.
        mask_loai_tru = _khoa_buu_cuc(df_bc_goc["Bưu cục"]).isin(BUU_CUC_LOAI_TRU)
        df_bc = df_bc_goc.loc[~mask_loai_tru].copy()

        if df_bc.empty:
            st.warning("Sau khi loại trừ DT TN / SETN không còn dữ liệu nào.")
            return

        # 1. Tổng volume IB (đã bỏ DT TN, SETN; chuẩn hóa và unique mã vận đơn)
        df_bc["Mã chuẩn"] = _chuan_hoa_ma(df_bc["Mã vận đơn"])
        tong_ib = df_bc["Mã chuẩn"].nunique()

        # 2. Tổng số đơn hàng được gửi đúng COT (按COT准时出库的订单量)
        df_cot = df_bc[df_bc["Ontime"].astype(str).str.strip().str.lower() == "giao đúng cot"]
        tong_dung_cot = df_cot["Mã chuẩn"].nunique()

        # Tỷ lệ COT = đúng COT / tổng IB (đã loại trừ)
        ty_le_cot = (tong_dung_cot / tong_ib * 100) if tong_ib > 0 else 0

        # 3. Tổng lượng hàng Inbound 1AM (1AM 入库件量)
        df_1am_truoc = df_bc[df_bc["Ontime1AM"].astype(str).str.strip().str.lower() == "trước"]
        tong_1am_truoc = df_1am_truoc["Mã chuẩn"].nunique()

        # 4. Số đơn đến trước 1AM VÀ Giao đúng COT
        df_1am_cot = df_1am_truoc[df_1am_truoc["Ontime"].astype(str).str.strip().str.lower() == "giao đúng cot"]
        tong_1am_dung_cot = df_1am_cot["Mã chuẩn"].nunique()

        ty_le_1am = (tong_1am_dung_cot / tong_1am_truoc * 100) if tong_1am_truoc > 0 else 0

        st.markdown("### Chỉ số Hiệu suất Ontime")
        r1, r2, r3 = st.columns(3)
        r1.metric("Tổng volume IB", f"{tong_ib:,}")
        r2.metric("Tổng số đơn hàng được gửi đúng COT (按COT准时出库的订单量)", f"{tong_dung_cot:,}")
        r3.metric("Tỷ lệ COT", f"{ty_le_cot:.2f}%")

        st.markdown("---")
        r4, r5, r6 = st.columns(3)
        r4.metric("Tổng lượng hàng Inbound 1AM (1AM 入库件量)", f"{tong_1am_truoc:,}")
        r5.metric("Số đơn trước 1AM & Đúng COT", f"{tong_1am_dung_cot:,}")
        r6.metric("Tỷ lệ On-time 1AM", f"{ty_le_1am:.2f}%")

        # 5. Tuyến chính (Linehaul): lọc Type OB = Linehaul + Trạng thái = Đúng giờ,
        #    đếm unique Mã vận đơn (giống công thức UNIQUE/FILTER trên Feishu)
        st.markdown("---")
        if "Type OB" in df_bc.columns and "Trạng thái" in df_bc.columns:
            df_lh_dung_gio = df_bc[
                (df_bc["Type OB"].astype(str).str.strip().str.lower() == "linehaul")
                & (df_bc["Trạng thái"].astype(str).str.strip().str.lower() == "đúng giờ")
                & (df_bc["Mã chuẩn"] != "")
            ]
            # Số đơn tuyến chính đến đúng hạn
            so_lh_den_dung_han = df_lh_dung_gio["Mã chuẩn"].nunique()
            # Số đơn tuyến chính gửi đi đúng hạn (thêm điều kiện Ontime = Giao đúng COT)
            so_lh_gui_dung_han = df_lh_dung_gio[
                df_lh_dung_gio["Ontime"].astype(str).str.strip().str.lower() == "giao đúng cot"
            ]["Mã chuẩn"].nunique()

            # Tỷ lệ = Số đơn tuyến chính gửi đi đúng hạn / Số đơn tuyến chính đến đúng hạn
            ty_le_lh_gui_di = (so_lh_gui_dung_han / so_lh_den_dung_han * 100) if so_lh_den_dung_han > 0 else 0

            r7, r8, r9 = st.columns(3)
            r7.metric("Số đơn tuyến chính đến đúng hạn / 干线准时到达票数", f"{so_lh_den_dung_han:,}")
            r8.metric("Số đơn tuyến chính gửi đi đúng hạn / 干线准时发出票数", f"{so_lh_gui_dung_han:,}")
            r9.metric("Tỷ lệ tuyến chính gửi đi đúng hạn / 干线准时发出率", f"{ty_le_lh_gui_di:.2f}%")

            # 6. Tuyến nhánh (Shuttle): giống tuyến chính nhưng Type OB = Shuttle
            df_sh_dung_gio = df_bc[
                (df_bc["Type OB"].astype(str).str.strip().str.lower() == "shuttle")
                & (df_bc["Trạng thái"].astype(str).str.strip().str.lower() == "đúng giờ")
                & (df_bc["Mã chuẩn"] != "")
            ]
            # Số đơn tuyến nhánh đến đúng hạn
            so_sh_den_dung_han = df_sh_dung_gio["Mã chuẩn"].nunique()
            # Số đơn tuyến nhánh gửi đi đúng hạn (thêm điều kiện Ontime = Giao đúng COT)
            so_sh_gui_dung_han = df_sh_dung_gio[
                df_sh_dung_gio["Ontime"].astype(str).str.strip().str.lower() == "giao đúng cot"
            ]["Mã chuẩn"].nunique()
            # Tỷ lệ = gửi đi đúng hạn / đến đúng hạn
            ty_le_sh_gui_di = (so_sh_gui_dung_han / so_sh_den_dung_han * 100) if so_sh_den_dung_han > 0 else 0

            st.markdown("---")
            r10, r11, r12 = st.columns(3)
            r10.metric("Số đơn tuyến nhánh đến đúng hạn / 支线准时到达票数", f"{so_sh_den_dung_han:,}")
            r11.metric("Số đơn tuyến nhánh gửi đi đúng hạn / 支线准时发出票数", f"{so_sh_gui_dung_han:,}")
            r12.metric("Tỷ lệ tuyến nhánh gửi đi đúng hạn / 支线准时发出率", f"{ty_le_sh_gui_di:.2f}%")
        else:
            st.warning("Bảng kpi_base chưa có cột 'Type OB' / 'Trạng thái' nên chưa tính được 2 chỉ số tuyến chính.")

        divider_label("Chi tiết dữ liệu báo cáo")
        st.dataframe(df_bc.head(500), use_container_width=True, height=400)

        cot_tai, _ = st.columns(2)
        with cot_tai:
            st.download_button(
                f"Tải Excel Báo Cáo ({len(df_bc):,} dòng)",
                xuat_excel({"Bao cao Ontime": df_bc}),
                file_name=f"bao_cao_ontime_{datetime.date.today()}.xlsx",
                use_container_width=True,
                type="primary",
            )


# ==========================================
# 6. ĐIỀU HƯỚNG CHÍNH (SIDEBAR)
# ==========================================
def main():
    with st.sidebar:
        st.markdown("## VẬN HÀNH QC")
        st.markdown("---")
        menu_lua_chon = st.radio(
            "ĐIỀU HƯỚNG:",
            ["SẢN LƯỢNG | 生产", "Linehaul Ontime Departure / 干线准时发车", "Ontime 准时报表"],
            label_visibility="collapsed"
        )
        st.markdown("---")
        st.caption(f"Hôm nay: {datetime.date.today().strftime('%d/%m/%Y')}")

    if menu_lua_chon == "SẢN LƯỢNG | 生产":
        page_du_lieu_tho()
    elif menu_lua_chon == "Linehaul Ontime Departure / 干线准时发车":
        page_ontime_xep_xe()
    elif menu_lua_chon == "Ontime 准时报表":
        page_bao_cao_ontime()

if __name__ == "__main__":
    main()
