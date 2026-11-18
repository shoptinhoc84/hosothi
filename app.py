import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Quản Lý Học Viên Thi Quá Hạn", page_icon="📋", layout="wide")

DATA_FILE = "danh_sach_qua_han.csv"

COLUMNS = [
    "Họ và tên",
    "Số CCCD",
    "Số điện thoại",
    "Hạng",
    "Ngày nhận hồ sơ",
    "Ngày thi cũ",
    "Nội dung rớt / vắng",
    "Lệ phí thi lại (VNĐ)",
    "Tình trạng phí",
    "Trạng thái thông báo",
    "Ngày dự kiến thi lại",
    "Ghi chú"
]

# Đã bổ sung đầy đủ hạng C và các hạng khác
HANG_GPLX = ["B", "B.01", "C", "C1", "D1", "D2", "D", "E", "CE", "A1", "A", "Khác"]
NOI_DUNG_ROT = ["Vắng thi", "Lý thuyết", "Mô phỏng", "Sa hình", "Đường trường", "Rớt nhiều phần"]
TINH_TRANG_PHI = ["Chưa đóng", "Đã đóng đủ", "Đóng một phần", "Miễn phí / Bảo lưu"]
TRANG_THAI_TB = ["Chưa thông báo", "Đã gọi / Nhắn Zalo", "Học viên đã xác nhận", "Đã xếp lịch thi"]


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_csv(DATA_FILE, dtype={"Số CCCD": str, "Số điện thoại": str})
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            return df[COLUMNS]
        except Exception:
            pass
    return pd.DataFrame(columns=COLUMNS)


def save_data(df):
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")


def to_printable_excel_bytes(df):
    output = BytesIO()
    export_df = df.copy()
    export_df.insert(0, "STT", range(1, len(export_df) + 1))

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="DS_Qua_Han", startrow=2)
        ws = writer.sheets["DS_Qua_Han"]

        # Tiêu đề bảng in
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(export_df.columns))
        title_cell = ws.cell(row=1, column=1)
        title_cell.value = f"DANH SÁCH HỌC VIÊN THI SÁT HẠCH QUÁ HẠN (Ngày xuất: {date.today().strftime('%d/%m/%Y')})"
        title_cell.font = Font(name="Times New Roman", size=14, bold=True)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        # Định dạng viền và font chữ chuẩn in ấn
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        # Định dạng dòng tiêu đề cột (row 3)
        ws.row_dimensions[3].height = 24
        for col_idx in range(1, len(export_df.columns) + 1):
            cell = ws.cell(row=3, column=col_idx)
            cell.font = Font(name="Times New Roman", size=11, bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Định dạng dữ liệu từng hàng
        for row_idx in range(4, 4 + len(export_df)):
            ws.row_dimensions[row_idx].height = 22
            for col_idx in range(1, len(export_df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Times New Roman", size=11)
                cell.border = thin_border
                col_name = export_df.columns[col_idx - 1]
                if col_name in ["STT", "Hạng", "Ngày nhận hồ sơ", "Ngày thi cũ", "Ngày dự kiến thi lại", "Số ngày quá hạn"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "Lệ phí thi lại (VNĐ)":
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        # Tự động căn độ rộng cột vừa trang A4 ngang
        col_widths = {
            "STT": 6,
            "Họ và tên": 24,
            "Số CCCD": 15,
            "Số điện thoại": 14,
            "Hạng": 8,
            "Ngày nhận hồ sơ": 15,
            "Ngày thi cũ": 14,
            "Nội dung rớt / vắng": 18,
            "Lệ phí thi lại (VNĐ)": 16,
            "Tình trạng phí": 16,
            "Trạng thái thông báo": 20,
            "Ngày dự kiến thi lại": 16,
            "Số ngày quá hạn": 14,
            "Ghi chú": 22
        }
        for col_idx, col_name in enumerate(export_df.columns, start=1):
            col_letter = ws.cell(row=3, column=col_idx).column_letter
            ws.column_dimensions[col_letter].width = col_widths.get(col_name, 15)

        # Thiết lập trang in A4 ngang vừa 1 chiều rộng trang
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    return output.getvalue()


def calc_overdue_days(row):
    ref_str = str(row.get("Ngày thi cũ", "")).strip()
    if not ref_str or ref_str == "nan":
        ref_str = str(row.get("Ngày nhận hồ sơ", "")).strip()
    try:
        ref_date = datetime.strptime(ref_str, "%d/%m/%Y").date()
        delta = (date.today() - ref_date).days
        return max(delta, 0)
    except Exception:
        return 0


if "df_students" not in st.session_state:
    st.session_state.df_students = load_data()

df = st.session_state.df_students.copy()

# ================= HEADER & THỐNG KÊ NHANH =================
st.title("📋 Quản Lý Danh Sách Thi Sát Hạch Quá Hạn")

if not df.empty:
    df["Số ngày quá hạn"] = df.apply(calc_overdue_days, axis=1)
else:
    df["Số ngày quá hạn"] = pd.Series(dtype="int")

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.metric("Tổng hồ sơ quá hạn", f"{len(df)} học viên")
with col_m2:
    chua_tb = len(df[df["Trạng thái thông báo"] == "Chưa thông báo"]) if not df.empty else 0
    st.metric("Chưa thông báo", f"{chua_tb} học viên")
with col_m3:
    chua_phi = len(df[df["Tình trạng phí"] == "Chưa đóng"]) if not df.empty else 0
    st.metric("Chưa đóng phí thi lại", f"{chua_phi} học viên")
with col_m4:
    tong_tien = pd.to_numeric(df["Lệ phí thi lại (VNĐ)"], errors="coerce").fillna(0).sum() if not df.empty else 0
    st.metric("Tổng phí dự kiến", f"{int(tong_tien):,} đ")

st.markdown("---")

# ================= SIDEBAR: KHÔI PHỤC DỮ LIỆU =================
with st.sidebar:
    st.header("📤 Khôi phục dữ liệu cũ")
    st.caption("Tải lên file Excel/CSV đã lưu trước đó để nạp lại danh sách vào hệ thống:")
    uploaded_file = st.file_uploader("Chọn file Excel (.xlsx) hoặc CSV", type=["xlsx", "csv"])
    if uploaded_file is not None:
        if st.button("Xác nhận nạp dữ liệu", use_container_width=True):
            try:
                if uploaded_file.name.endswith(".csv"):
                    new_df = pd.read_csv(uploaded_file, dtype={"Số CCCD": str, "Số điện thoại": str})
                else:
                    # Tự động nhận diện file Excel xuất từ hệ thống (có dòng tiêu đề ở row 2) hoặc file thường
                    preview_df = pd.read_excel(uploaded_file, nrows=3)
                    skip = 2 if "DANH SÁCH HỌC VIÊN" in str(preview_df.columns[0]) else 0
                    uploaded_file.seek(0)
                    new_df = pd.read_excel(uploaded_file, skiprows=skip, dtype={"Số CCCD": str, "Số điện thoại": str})
                for col in COLUMNS:
                    if col not in new_df.columns:
                        new_df[col] = ""
                new_df = new_df[COLUMNS].fillna("")
                st.session_state.df_students = new_df
                save_data(new_df)
                st.success("Đã khôi phục dữ liệu thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")

# ================= TABS CHỨC NĂNG =================
tab1, tab2 = st.tabs(["➕ Thêm học viên mới", "📊 Danh sách & Cập nhật thông báo"])

# --- TAB 1: THÊM MỚI ---
with tab1:
    with st.form("form_add_student", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            ho_ten = st.text_input("Họ và tên học viên *")
            cccd = st.text_input("Số CCCD")
            sdt = st.text_input("Số điện thoại / Zalo *")
            hang = st.selectbox("Hạng GPLX *", HANG_GPLX, index=2)  # Mặc định chọn hạng C
        with c2:
            ngay_nhan = st.date_input("Ngày nhận hồ sơ", value=date.today(), format="DD/MM/YYYY")
            ngay_thi_cu = st.date_input("Ngày thi cũ (bị rớt/vắng)", value=date.today(), format="DD/MM/YYYY")
            noi_dung_rot = st.selectbox("Nội dung rớt / vắng", NOI_DUNG_ROT)
            ngay_du_kien = st.text_input("Ngày dự kiến thi lại (VD: 25/11/2026)", value="")
        with c3:
            le_phi = st.number_input("Lệ phí thi lại (VNĐ)", min_value=0, step=50000, value=600000)
            tinh_trang_phi = st.selectbox("Tình trạng phí", TINH_TRANG_PHI)
            trang_thai_tb = st.selectbox("Trạng thái thông báo", TRANG_THAI_TB)
            ghi_chu = st.text_area("Ghi chú thêm", height=68)

        submitted = st.form_submit_button("💾 Lưu vào danh sách quá hạn", type="primary", use_container_width=True)
        if submitted:
            if not ho_ten.strip():
                st.warning("Vui lòng nhập Họ và tên học viên.")
            else:
                new_row = pd.DataFrame([{
                    "Họ và tên": ho_ten.strip().upper(),
                    "Số CCCD": cccd.strip(),
                    "Số điện thoại": sdt.strip(),
                    "Hạng": hang,
                    "Ngày nhận hồ sơ": ngay_nhan.strftime("%d/%m/%Y"),
                    "Ngày thi cũ": ngay_thi_cu.strftime("%d/%m/%Y"),
                    "Nội dung rớt / vắng": noi_dung_rot,
                    "Lệ phí thi lại (VNĐ)": int(le_phi),
                    "Tình trạng phí": tinh_trang_phi,
                    "Trạng thái thông báo": trang_thai_tb,
                    "Ngày dự kiến thi lại": ngay_du_kien.strip(),
                    "Ghi chú": ghi_chu.strip()
                }])
                updated_df = pd.concat([st.session_state.df_students, new_row], ignore_index=True)
                st.session_state.df_students = updated_df
                save_data(updated_df)
                st.success(f"Đã thêm học viên **{ho_ten.upper()}** (Hạng {hang}) vào danh sách!")
                st.rerun()

# --- TAB 2: DANH SÁCH, LỌC, SỬA & XUẤT EXCEL IN ---
with tab2:
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        search_txt = st.text_input("🔍 Tìm tên, CCCD hoặc SĐT")
    with f2:
        filter_hang = st.multiselect("Lọc theo Hạng", HANG_GPLX, default=[])
    with f3:
        filter_tb = st.multiselect("Lọc Trạng thái thông báo", TRANG_THAI_TB, default=[])
    with f4:
        filter_phi = st.multiselect("Lọc Tình trạng phí", TINH_TRANG_PHI, default=[])

    view_df = df.copy()

    if search_txt.strip():
        kw = search_txt.strip().lower()
        view_df = view_df[
            view_df["Họ và tên"].astype(str).str.lower().str.contains(kw) |
            view_df["Số CCCD"].astype(str).str.lower().str.contains(kw) |
            view_df["Số điện thoại"].astype(str).str.lower().str.contains(kw)
        ]
    if filter_hang:
        view_df = view_df[view_df["Hạng"].isin(filter_hang)]
    if filter_tb:
        view_df = view_df[view_df["Trạng thái thông báo"].isin(filter_tb)]
    if filter_phi:
        view_df = view_df[view_df["Tình trạng phí"].isin(filter_phi)]

    st.caption("💡 Nhấp đúp vào từng ô để sửa trực tiếp hoặc bấm dấu `+` ở cuối bảng để thêm nhanh dòng mới, sau đó bấm **Lưu thay đổi**.")

    edited_df = st.data_editor(
        view_df,
        use_container_width=True,
        num_rows="dynamic",
        disabled=["Số ngày quá hạn"],
        column_config={
            "Hạng": st.column_config.SelectboxColumn("Hạng", options=HANG_GPLX),
            "Nội dung rớt / vắng": st.column_config.SelectboxColumn("Nội dung rớt / vắng", options=NOI_DUNG_ROT),
            "Tình trạng phí": st.column_config.SelectboxColumn("Tình trạng phí", options=TINH_TRANG_PHI),
            "Trạng thái thông báo": st.column_config.SelectboxColumn("Trạng thái thông báo", options=TRANG_THAI_TB),
            "Lệ phí thi lại (VNĐ)": st.column_config.NumberColumn("Lệ phí (VNĐ)", format="%d"),
            "Số ngày quá hạn": st.column_config.NumberColumn("Quá hạn (ngày)", format="%d ngày")
        },
        key="editor_students"
    )

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if st.button("💾 Lưu tất cả thay đổi trên bảng", type="primary", use_container_width=True):
            clean_edited = edited_df.drop(columns=["Số ngày quá hạn"], errors="ignore")
            if not search_txt and not filter_hang and not filter_tb and not filter_phi:
                st.session_state.df_students = clean_edited.reset_index(drop=True)
            else:
                base = st.session_state.df_students.copy()
                base.update(clean_edited)
                st.session_state.df_students = base
            save_data(st.session_state.df_students)
            st.success("Đã lưu cập nhật danh sách!")
            st.rerun()

    with btn_col2:
        # Xuất đúng danh sách đang lọc/hiển thị để tiện in theo từng hạng hoặc từng đợt
        excel_bytes = to_printable_excel_bytes(view_df)
        st.download_button(
            label="🖨️ Tải file Excel (.xlsx) để in danh sách",
            data=excel_bytes,
            file_name=f"DS_Thi_Qua_Han_{date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
