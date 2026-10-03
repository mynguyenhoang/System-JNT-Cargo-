# -*- coding: utf-8 -*-
"""
KPI calculation - bước đầu

Logic:
1) Ontime1AM từ "Thời gian quét" (>= 06:00 hoặc <= 01:00 -> "trước", còn lại -> "sau")

2) Lấy Bưu cục đích (XLOOKUP, lấy dòng đầu tiên):
   - Ưu tiên raw_quet_hang_xep_len_xe:
       khóa = Bưu cục quét + Mã vận đơn + Ngày vận hành
       -> "Trạm trước hoặc trạm tiếp theo"
   - Nếu không có -> raw_backlog:
       khóa = Bộ phận kho vận + Mã số vận đơn
       -> "Điểm giao dịch đích đến"

3) Scan OB: thời gian quét bên Xếp lên xe, khóa = Bưu cục quét + Mã vận đơn + Ngày vận hành
4) Ontime: Scan OB < DATE CUTOFF -> Giao đúng COT, ngược lại Giao trễ COT

(COT) khóa = Bưu cục quét + Bưu cục đích -> raw_cot.cot -> OB

5) Type OB: Bưu cục thuộc HCM HUB / BN HUB / CTO SC / SH DC -> "Linehaul", còn lại -> "Shuttle"
6) Thời gian xe đến kế hoạch / thực tế: tra raw_quan_ly_nhiem_vu_con theo Số nhiệm vụ
   (thực tế còn khớp thêm Bộ phận đến = Bưu cục quét)
7) Trạng thái: xe đến thực tế < kế hoạch -> "Đúng giờ", ngược lại "Trễ giờ"
"""

import sqlite3
import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent
DB = BASE / "feishu_archive_local.db"

SEP = "|"  # tránh ghép nhầm khi nối chuỗi

# Giờ chốt ngày theo Bưu cục quét (quét trước giờ này -> tính ngày quét, từ giờ này -> +1 ngày)
#   SH DC, CTO SC     : 8h hôm qua -> 8h hôm nay = 1 ngày
#   HCM HUB, BN HUB   : 6h hôm qua -> 6h hôm nay = 1 ngày
DAY_CUT_HOUR = {
    "SH DC": 8,
    "CTO SC": 8,
    "HCM HUB": 6,
    "BN HUB": 6,
}
DEFAULT_DAY_CUT_HOUR = 6  # bưu cục không có trong danh sách trên

# Type OB: cột dùng để phân loại và danh sách bưu cục tính là Linehaul
TYPE_OB_SOURCE_COL = "Bưu cục"          # đổi thành "Bưu cục quét" nếu muốn phân loại theo bưu cục quét
LINEHAUL_SET = {"HCM HUB", "BN HUB", "CTO SC", "SH DC"}


def _norm(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def _find_col(df, names):
    for name in names:
        if name in df.columns:
            return name
    return None


def _norm_date(series):
    """Chuẩn hóa Ngày vận hành về 'YYYY-MM-DD' (nhận '10/01/2026' kiểu M/D/Y hoặc datetime)."""
    d = pd.to_datetime(series, format="%m/%d/%Y", errors="coerce")
    miss = d.isna() & series.notna()
    if miss.any():
        d[miss] = pd.to_datetime(series[miss], errors="coerce")
    out = d.dt.strftime("%Y-%m-%d")
    # không đọc được ngày -> giữ nguyên chuỗi gốc để vẫn so sánh được
    return out.where(d.notna(), series.map(_norm))


def _to_datetime(series):
    """Đọc thời gian dạng "09/13/2026 04:03:03" (M/D/Y) hoặc ISO; trống/lỗi -> NaT."""
    d = pd.to_datetime(series, format="%m/%d/%Y %H:%M:%S", errors="coerce")
    miss = d.isna() & series.notna() & (series.astype(str).str.strip() != "")
    if miss.any():
        try:
            d[miss] = pd.to_datetime(series[miss], format="mixed", errors="coerce")
        except (ValueError, TypeError):
            d[miss] = pd.to_datetime(series[miss], errors="coerce")
    return d


def _make_key(df, scan_col, ma_col, ngay_col=None):
    key = df[scan_col].map(_norm) + SEP + df[ma_col].map(_norm)
    if ngay_col is not None:
        key = key + SEP + _norm_date(df[ngay_col])
    return key


def _cot_to_timedelta(series):
    """
    Chuyển cột COT thành timedelta để cộng vào ngày (giống Excel cộng số/giờ).
    Hỗ trợ: số kiểu Excel (0.0417 = 01:00), chuỗi "HH:MM[:SS]",
    hoặc datetime (chỉ lấy phần giờ).
    """
    out = pd.Series(pd.NaT, index=series.index, dtype="timedelta64[ns]")

    # 1) Số (phần ngày kiểu Excel)
    num = pd.to_numeric(series, errors="coerce")
    m = num.notna()
    out[m] = pd.to_timedelta(num[m], unit="D")

    # 2) Chuỗi giờ "HH:MM:SS"
    rest = ~m & series.notna() & (series.astype(str).str.strip() != "")
    if rest.any():
        td = pd.to_timedelta(series[rest].astype(str).str.strip(), errors="coerce")
        out[rest] = td

        # 3) Datetime -> lấy phần giờ
        still = out.isna() & rest
        if still.any():
            dt = pd.to_datetime(series[still], errors="coerce")
            out[still] = dt - dt.dt.normalize()
    return out


def load_raw():
    con = sqlite3.connect(DB)
    unload = pd.read_sql_query('SELECT * FROM "raw_quet_do_xuong_xe"', con)
    load = pd.read_sql_query('SELECT * FROM "raw_quet_hang_xep_len_xe"', con)
    backlog = pd.read_sql_query('SELECT * FROM "raw_backlog"', con)
    cot = pd.read_sql_query('SELECT * FROM "raw_cot"', con)
    nvc = pd.read_sql_query('SELECT * FROM "raw_quan_ly_nhiem_vu_con"', con)
    con.close()
    return unload, load, backlog, cot, nvc


def build_kpi_base():
    unload, load, backlog, cot, nvc = load_raw()

    # ===== 1. XÁC ĐỊNH CỘT =====
    ma_unload = _find_col(unload, ["Mã vận đơn"])
    ngay_unload = _find_col(unload, ["Ngày vận hành"])
    nhiem_vu_unload = _find_col(unload, ["Số nhiệm vụ"])
    scan_time = _find_col(unload, ["Thời gian quét"])
    scan_unload = _find_col(unload, ["Bưu cục quét"])

    ma_load = _find_col(load, ["Mã vận đơn"])
    ngay_load = _find_col(load, ["Ngày vận hành"])
    # Cột bưu cục quét bên "Quét hàng xếp lên xe" - sửa/thêm tên nếu header khác
    scan_load = _find_col(load, ["Bưu cục quét", "Bưu cục", "Trạm quét"])
    tram_tiep_theo = _find_col(load, ["Trạm trước hoặc trạm tiếp theo"])
    scan_time_load = _find_col(load, ["Thời gian quét"])

    ma_backlog = _find_col(backlog, ["Mã số vận đơn", "Mã vận đơn"])
    kho_backlog = _find_col(backlog, ["Bộ phận kho vận"])
    backlog_dich_den = _find_col(backlog, ["Điểm giao dịch đích đến"])

    nvc_ma = _find_col(nvc, ["Mã nhiệm vụ"])
    nvc_den = _find_col(nvc, ["Bộ phận đến"])
    nvc_ke_hoach = _find_col(nvc, ["Thời gian dự kiến xe đến", "Thời gian xe đến dự kiến"])
    nvc_xe_den = _find_col(nvc, ["Thời gian xe đến"])
    nvc_bd_do = _find_col(nvc, ["Thời gian bắt đầu dỡ hàng"])

    cot_key = _find_col(cot, ["cot"])
    cot_ob = _find_col(cot, ["OB"])

    required = {
        "raw_quet_do_xuong_xe.Mã vận đơn": ma_unload,
        "raw_quet_do_xuong_xe.Thời gian quét": scan_time,
        "raw_quet_do_xuong_xe.Bưu cục quét": scan_unload,
        "raw_quet_do_xuong_xe.Ngày vận hành": ngay_unload,
        "raw_quet_hang_xep_len_xe.Ngày vận hành": ngay_load,
        "raw_quet_hang_xep_len_xe.Mã vận đơn": ma_load,
        "raw_quet_hang_xep_len_xe.Bưu cục quét": scan_load,
        "raw_quet_hang_xep_len_xe.Trạm trước hoặc trạm tiếp theo": tram_tiep_theo,
        "raw_quet_hang_xep_len_xe.Thời gian quét": scan_time_load,
        "raw_backlog.Mã số vận đơn": ma_backlog,
        "raw_backlog.Bộ phận kho vận": kho_backlog,
        "raw_backlog.Điểm giao dịch đích đến": backlog_dich_den,
        "raw_quet_do_xuong_xe.Số nhiệm vụ": nhiem_vu_unload,
        "raw_quan_ly_nhiem_vu_con.Mã nhiệm vụ": nvc_ma,
        "raw_quan_ly_nhiem_vu_con.Bộ phận đến": nvc_den,
        "raw_quan_ly_nhiem_vu_con.Thời gian dự kiến xe đến": nvc_ke_hoach,
        "raw_quan_ly_nhiem_vu_con.Thời gian xe đến": nvc_xe_den,
        "raw_quan_ly_nhiem_vu_con.Thời gian bắt đầu dỡ hàng": nvc_bd_do,
        "raw_cot.cot": cot_key,
        "raw_cot.OB": cot_ob,
    }
    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(
            "Thiếu cột cần thiết trong DB:\n- " + "\n- ".join(missing)
        )

    # ===== 2. BASE =====
    df = unload.copy()

    # ===== 3. ONTIME 1AM =====
    ts = pd.to_datetime(df[scan_time], errors="coerce")
    mins = ts.dt.hour * 60 + ts.dt.minute + ts.dt.second / 60
    ontime = (mins >= 360) | (mins <= 60)

    df["Ontime1AM"] = ""
    df.loc[ts.notna() & ontime, "Ontime1AM"] = "trước"
    df.loc[ts.notna() & ~ontime, "Ontime1AM"] = "sau"

    # ===== 4. BƯU CỤC ĐÍCH =====
    # Khóa bên đơn nhận vào: Bưu cục quét + Mã vận đơn
    # Xếp lên xe: Bưu cục quét + Mã vận đơn + Ngày vận hành
    df["__key"] = _make_key(df, scan_unload, ma_unload, ngay_unload)
    # Backlog: Bưu cục quét + Mã vận đơn (không có ngày vận hành)
    df["__key_bl"] = _make_key(df, scan_unload, ma_unload)

    # --- Xếp lên xe: Bưu cục quét + Mã vận đơn + Ngày vận hành -> Trạm tiếp theo, Scan OB ---
    load_map = pd.DataFrame({
        "__key": _make_key(load, scan_load, ma_load, ngay_load),
        "__dich_load": load[tram_tiep_theo].map(_norm),
        "__scan_ob": pd.to_datetime(load[scan_time_load], errors="coerce"),
    })
    load_map = load_map[
        load[ma_load].map(_norm).values != ""
    ].drop_duplicates(subset=["__key"], keep="first")

    # --- Backlog: Bộ phận kho vận + Mã số vận đơn -> Điểm giao dịch đích đến ---
    backlog_map = pd.DataFrame({
        "__key_bl": _make_key(backlog, kho_backlog, ma_backlog),
        "__dich_backlog": backlog[backlog_dich_den].map(_norm),
    })
    backlog_map = backlog_map[
        backlog[ma_backlog].map(_norm).values != ""
    ].drop_duplicates(subset=["__key_bl"], keep="first")

    df = df.merge(load_map, on="__key", how="left", sort=False)
    df = df.merge(backlog_map, on="__key_bl", how="left", sort=False)

    dich_load = df["__dich_load"].fillna("").map(_norm)
    dich_backlog = df["__dich_backlog"].fillna("").map(_norm)

    # Scan OB = thời gian quét bên Xếp lên xe (Bưu cục quét + Mã vận đơn + Ngày vận hành)
    df["Scan OB"] = pd.to_datetime(df["__scan_ob"], errors="coerce")

    # Ưu tiên Xếp lên xe; chỉ khi không có mới dùng Backlog
    df["Bưu cục"] = dich_load.where(dich_load != "", dich_backlog)

    # ===== 5. COT / OB =====
    df["__cot_key"] = (
        df[scan_unload].fillna("").map(_norm) + df["Bưu cục"].map(_norm)
    )

    cot_map = pd.DataFrame({
        "__cot_key": cot[cot_key].map(_norm),
        "COT": cot[cot_ob],
    })
    cot_map = cot_map[cot_map["__cot_key"] != ""].drop_duplicates(
        subset=["__cot_key"], keep="first"
    )

    df = df.merge(cot_map, on="__cot_key", how="left")

    # ===== 6. DATE CUTOFF =====
    # Excel: =IFERROR((INT(tg) + IF(MOD(tg,1) < TIME(h,0,0), 0, 1)) + COT, "")
    # - h phụ thuộc Bưu cục quét: SH DC / CTO SC = 8h, HCM HUB / BN HUB = 6h
    # - Quét trước h -> tính ngày quét; từ h trở đi -> +1 ngày
    # - Cộng thêm COT; lỗi/trống -> ""
    ts_all = pd.to_datetime(df[scan_time], errors="coerce")
    time_of_day = ts_all - ts_all.dt.normalize()

    cut_hour = (
        df[scan_unload].map(_norm).str.upper()
        .map(DAY_CUT_HOUR)
        .fillna(DEFAULT_DAY_CUT_HOUR)
    )
    add_day = (time_of_day >= pd.to_timedelta(cut_hour, unit="h")).astype(int)

    date_cutoff = (
        ts_all.dt.normalize()
        + pd.to_timedelta(add_day, unit="D")
        + _cot_to_timedelta(df["COT"])
    )
    df["DATE CUTOFF"] = date_cutoff  # kiểu datetime (lỗi/trống -> NaT = ô trống)

    # ===== 7. ONTIME (Giao đúng / trễ COT) =====
    # Scan OB < DATE CUTOFF -> "Giao đúng COT", ngược lại -> "Giao trễ COT"
    # Không có Scan OB -> ""
    has_both = df["Scan OB"].notna() & df["DATE CUTOFF"].notna()
    df["Ontime"] = ""
    df.loc[has_both & (df["Scan OB"] < df["DATE CUTOFF"]), "Ontime"] = "Giao đúng COT"
    df.loc[has_both & ~(df["Scan OB"] < df["DATE CUTOFF"]), "Ontime"] = "Giao trễ COT"
    # Giống Excel/Feishu: có Scan OB nhưng không có DATE CUTOFF (bưu cục không có COT,
    # ví dụ DT TN, SETN) -> số < ô trống nên IF(S<Q) ra TRUE -> tính là đúng.
    df.loc[df["Scan OB"].notna() & df["DATE CUTOFF"].isna(), "Ontime"] = "Giao đúng COT"

    # ===== 7B. TYPE OB =====
    # Excel: =IF(OR(x="HCM HUB", x="BN HUB", x="CTO SC", x="SH DC"), "Linehaul", "Shuttle")
    df["Type OB"] = (
        df[TYPE_OB_SOURCE_COL].map(_norm).str.upper()
        .isin(LINEHAUL_SET)
        .map({True: "Linehaul", False: "Shuttle"})
    )

    # ===== 7C. THỜI GIAN XE ĐẾN (từ Quản lý nhiệm vụ con) =====
    nv_key = df[nhiem_vu_unload].map(_norm).str.upper()

    # Kế hoạch: XLOOKUP(Số nhiệm vụ, Mã nhiệm vụ, Thời gian dự kiến xe đến) - dòng đầu tiên
    plan_map = pd.DataFrame({
        "__nv": nvc[nvc_ma].map(_norm).str.upper(),
        "Thời gian xe đến kế hoạch": _to_datetime(nvc[nvc_ke_hoach]),
    })
    plan_map = plan_map[plan_map["__nv"] != ""].drop_duplicates(subset=["__nv"], keep="first")
    df["__nv"] = nv_key
    df = df.merge(plan_map, on="__nv", how="left", sort=False)

    # Thực tế: MAXIFS theo (Mã nhiệm vụ + Bộ phận đến = Bưu cục quét)
    #   ưu tiên "Thời gian xe đến"; không có thì "Thời gian bắt đầu dỡ hàng"; không có -> trống
    act = pd.DataFrame({
        "__nv": nvc[nvc_ma].map(_norm).str.upper(),
        "__den": nvc[nvc_den].map(_norm).str.upper(),
        "__xe_den": _to_datetime(nvc[nvc_xe_den]),
        "__bd_do": _to_datetime(nvc[nvc_bd_do]),
    })
    act = act[act["__nv"] != ""].groupby(["__nv", "__den"], as_index=False).max()  # max bỏ qua NaT

    df["__den"] = df[scan_unload].map(_norm).str.upper()
    df = df.merge(act, on=["__nv", "__den"], how="left", sort=False)
    df["Thời gian xe đến thực tế"] = df["__xe_den"].where(
        df["__xe_den"].notna(), df["__bd_do"]
    )

    # ===== 7D. TRẠNG THÁI (Đúng giờ / Trễ giờ) =====
    # Excel: =IF(Y2<X2, "Đúng giờ", "Trễ giờ")  (Y = xe đến thực tế, X = xe đến kế hoạch)
    # Thiếu kế hoạch hoặc thực tế -> ""
    plan = df["Thời gian xe đến kế hoạch"]
    actual = df["Thời gian xe đến thực tế"]
    has_both_xe = plan.notna() & actual.notna()
    df["Trạng thái"] = ""
    df.loc[has_both_xe & (actual < plan), "Trạng thái"] = "Đúng giờ"
    df.loc[has_both_xe & ~(actual < plan), "Trạng thái"] = "Trễ giờ"

    # Bỏ cột tạm
    df.drop(
        columns=[c for c in ["__key", "__key_bl", "__dich_load", "__dich_backlog", "__cot_key", "__scan_ob", "__nv", "__den", "__xe_den", "__bd_do"]
                 if c in df.columns],
        inplace=True,
    )

    # Bỏ 2 cột không dùng
    df.drop(columns=[c for c in ["Loại quét", "HUB"] if c in df.columns], inplace=True)

    # ===== 8. ĐỊNH DẠNG CỘT =====
    # Mã vận đơn -> Number (Int64; giá trị không phải số sẽ thành ô trống)
    df[ma_unload] = pd.to_numeric(df[ma_unload], errors="coerce").astype("Int64")
    # DATE CUTOFF -> kiểu thời gian (datetime64)
    df["DATE CUTOFF"] = pd.to_datetime(df["DATE CUTOFF"], errors="coerce")

    return df


if __name__ == "__main__":
    result = build_kpi_base()

    print("=" * 70)
    print("KPI BASE")
    print(f"Rows: {len(result):,}")
    print("Các cột logic đã tạo: Ontime1AM, Bưu cục, COT, DATE CUTOFF, Scan OB, Ontime, Type OB, Thời gian xe đến kế hoạch / thực tế, Trạng thái")
    print("=" * 70)

    # ===== XUẤT 1 FILE EXCEL DUY NHẤT (đúng kiểu Number / Time) =====
    xlsx_output = BASE / "kpi_base_preview.xlsx"
    with pd.ExcelWriter(xlsx_output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="kpi_base_preview")
        ws = writer.sheets["kpi_base_preview"]
        headers = [c.value for c in ws[1]]

        # Mã vận đơn -> Number, không hiện dạng khoa học
        col_ma = headers.index("Mã vận đơn") + 1
        for row in ws.iter_rows(min_row=2, min_col=col_ma, max_col=col_ma):
            for cell in row:
                cell.number_format = "0"

        # Các cột thời gian -> Time (ngày + giờ)
        time_cols = ["DATE CUTOFF", "Scan OB",
                     "Thời gian xe đến kế hoạch", "Thời gian xe đến thực tế"]
        for name in time_cols:
            col = headers.index(name) + 1
            for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
                for cell in row:
                    cell.number_format = "yyyy-mm-dd hh:mm:ss"
            ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 22

        # Cho cột Mã vận đơn đủ rộng để không bị ####
        ws.column_dimensions[ws.cell(row=1, column=col_ma).column_letter].width = 18

    print(f"Excel: {xlsx_output}")
