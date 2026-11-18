import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from io import BytesIO

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

HANG_GPLX = ["B", "B.01", "C1", "D2", "E", "CE", "A1", "A", "Khác"]
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


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="QuaHanSatHach")
    return output.getvalue()


def calc_overdue_days(row):
    # Ưu tiên tính từ Ngày thi cũ, nếu trống thì tính từ Ngày nhận hồ sơ
    ref_str = str(row.get("Ngày thi cũ", "")).strip()
    if not ref_str or ref_str == "nan":
        ref_str = str(row.get("Ngày nhận hồ sơ", "")).strip()
    try:
        ref_date = datetime.strptime(ref_str, "%d/%m/%Y").date()
        delta = (date.today() - ref_date).days
        return max(delta, 0)
    except Exception:
        return 0


# Khởi tạo session state
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

# ================= SIDEBAR: SAO LƯU & KHÔI PHỤC DỮ LIỆU =================
with st.sidebar:
    st.header("💾 Sao lưu & Khôi phục")
    st.caption("Lưu ý: Nếu dùng Render gói Free, hãy tải file Excel về máy sau khi nhập xong để tránh mất dữ liệu khi server khởi động lại.")

    if not df.empty:
        excel_data = to_excel_bytes(df.drop(columns=["Số ngày quá hạn"], errors="ignore"))
        st.download_button(
            label="📥 Tải danh sách về Excel (.xlsx)",
            data=excel_data,
            file_name=f"DS_Qua_Han_Sat_Hach_{date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    uploaded_file = st.file_uploader("📤 Nạp lại dữ liệu từ file Excel/CSV", type=["xlsx", "csv"])
    if uploaded_file is not None:
        if st.button("Xác nhận nạp dữ liệu", use_container_width=True):
            try:
                if uploaded_file.name.endswith(".csv"):
                    new_df = pd.read_csv(uploaded_file, dtype={"Số CCCD": str, "Số điện thoại": str})
                else:
                    new_df = pd.read_excel(uploaded_file, dtype={"Số CCCD": str, "Số điện thoại": str})
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
            hang = st.selectbox("Hạng GPLX *", HANG_GPLX, index=4)  # Mặc định chọn hạng E
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

# --- TAB 2: DANH SÁCH, LỌC & SỬA TRỰC TIẾP ---
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

    st.caption("💡 Bạn có thể nhấp đúp trực tiếp vào từng ô trên bảng bên dưới để sửa trạng thái thông báo, tiền phí, ngày thi lại hoặc chọn dòng để xóa, sau đó bấm **Lưu thay đổi**.")

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

    if st.button("💾 Lưu tất cả thay đổi trên bảng", type="primary"):
        clean_edited = edited_df.drop(columns=["Số ngày quá hạn"], errors="ignore")
        # Nếu đang không lọc thì lưu đè toàn bộ; nếu đang lọc thì cập nhật theo index
        if not search_txt and not filter_hang and not filter_tb and not filter_phi:
            st.session_state.df_students = clean_edited.reset_index(drop=True)
        else:
            base = st.session_state.df_students.copy()
            base.update(clean_edited)
            st.session_state.df_students = base
        save_data(st.session_state.df_students)
        st.success("Đã lưu cập nhật danh sách!")
        st.rerun()