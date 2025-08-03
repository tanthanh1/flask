import pandas as pd

# Danh sách tên cột
ten_cot = [
    "STT", "IDChungTu/MaBill", "TenHangHoaDichVu", "DonViTinh/ChietKhau",
    "SoLuong", "DonGia", "ThanhTien", "ThueSuat", "TienThueGTGT", "NgayThangNamHD",
    "CCCD", "SoHoChieu", "HoTenNguoiMua", "TenDonVi", "MaNganSach", "MaSoThue",
    "DiaChi", "SoTaiKhoan", "HinhThucTT", "NhanBangEmail", "DSEmail",
    "NhanBangSMS", "DSSMS", "NhanBangBanIN", "HoTenNguoiNhan",
    "SoDienThoaiNguoiNhan", "SoNha", "Tinh/ThanhPho", "Huyen/Quan/ThiXa",
    "Xa/Phuong/ThiTran", "GhiChu"
]

# Tạo DataFrame trống với các cột trên
df_moi = pd.DataFrame(columns=ten_cot)


df_cu = pd.read_excel('input.xlsx')
# Ghi ra file Excel
df_moi.to_excel('file_mau.xlsx', index=False)