import pandas as pd

# Đọc file result.xlsx
df = pd.read_excel('input_0.xlsx')

df = df.drop(index=0)

# Reset lại chỉ số hàng (nếu muốn)
df = df.reset_index(drop=True)
df['SKU Subtotal After Discount'] = (
    df['SKU Subtotal After Discount']
    .astype(str)                  # chuyển thành chuỗi
    .str.replace(',', '')         # xóa dấu phẩy (nếu có, dùng cho kiểu 1,200.00)
    .str.strip()                  # loại bỏ khoảng trắng thừa
).astype(float)                   # chuyển thành số

# Tạo cột STT từ 1 đến n
df_moi = pd.DataFrame()
df_moi['STT'] = range(1, len(df) + 1)

# Gán các cột theo yêu cầu
df_moi['IDChungTu/MaBill'] = df['Order ID'].astype(str)
df_moi['TenHangHoaDichVu'] = df['Product Name']
# Tạo cột Đơn vị tính / Chiết khấu theo nội dung
df_moi['DonViTinh/ChietKhau'] = df['Product Name'].apply(
    lambda val: (
        'chiếc' if 'lược' in str(val).lower()
        else 'chai' if any(x in str(val).lower() for x in ['xịt tinh dầu', 'tinh dầu','nước hoa', 'dầu gội'])
        else 'dây' if 'thun buộc tóc' in str(val).lower()
        else 'hủ' if 'tế bào chết' in str(val).lower()
        else 'hủ'
    )
)

df_moi['SoLuong'] = df['Quantity'].astype(int)
#df_moi['SoLuong'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0).astype(int)

df_moi['ThanhTien'] = (df['SKU Subtotal After Discount'] / 1.08).astype(int)
df_moi['ThueSuat'] = 0.08
df_moi['TienThueGTGT'] = (df_moi['ThanhTien'] * 0.08).round(0).astype(int)
df_moi['DonGia'] = (df_moi['ThanhTien'] / df_moi['SoLuong']).astype(int)

cols = ['STT', 'IDChungTu/MaBill', 'TenHangHoaDichVu', 'DonViTinh/ChietKhau','SoLuong', 'DonGia', 'ThanhTien','ThueSuat','TienThueGTGT']
df_moi = df_moi[cols]


cot_bo_sung = [
    'NgayThangNamHD', 'CCCD', 'SoHoChieu', 'HoTenNguoiMua', 'TenDonVi',
    'MaNganSach', 'MaSoThue', 'DiaChi', 'SoTaiKhoan', 'HinhThucTT',
    'NhanBangEmail', 'DSEmail', 'NhanBangSMS', 'DSSMS', 'NhanBangBanIN',
    'HoTenNguoiNhan', 'SoDienThoaiNguoiNhan', 'SoNha', 'Tinh/ThanhPho',
    'Huyen/Quan/ThiXa', 'Xa/Phuong/ThiTran', 'GhiChu'
]

for cot in cot_bo_sung:
    df_moi[cot] = ""


# Ghi ra file ketqua.xlsx
df_moi.to_excel('ketqua.xlsx', index=False)