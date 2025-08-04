from flask import Flask, request, send_file, render_template_string
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

HTML_FORM = '''
<!doctype html>
<head>
    <meta charset="UTF-8">
    <title>Công ty bần nông tuyển kế toán bần nông</title>
    <style>
        body {
            background-color: #f8f9fa;
            font-family: Arial, sans-serif;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
            margin: 0;
        }
        .container {
            text-align: center;
            padding: 40px;
            background-color: white;
            border-radius: 12px;
            box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
        }
        h2,h1 {
            margin-bottom: 30px;
            color: #dc3545;
        }
        input[type=file] {
            margin-bottom: 20px;
        }
        input[type=submit] {
            padding: 10px 20px;
            font-size: 16px;
            background-color: #198754;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
        }
        input[type=submit]:hover {
            background-color: #157347;
        }
    </style>
</head>
<body class="container">
<div>

<h1>Chưa có ai xài chùa mà có kết cục tốt !!!</h1>
<h2>Người xài chùa, coi thường công sức người khác, nghèo 3 đời !!!</h2>
<h2>Bần nông từ sếp cho tới kế toán !!!</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value="Bấm vô đây">
</form>
</div>
</body>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_excel():
    if request.method == 'POST':
        f = request.files['file']
        if not f:
            return 'Không có file nào được tải lên.', 400

        # Đọc file Excel gốc
        df = pd.read_excel(f)
        

        #df = df_temp.drop(index=0)

        # Reset lại chỉ số hàng (nếu muốn)
        # df = df.reset_index(drop=True)

        print(df)
    
        df['SKU Subtotal After Discount'] = (
            df['SKU Subtotal After Discount']
            .astype(str).str.replace(',', '').str.strip()
        ).astype(float)

        df_moi = pd.DataFrame()
        df_moi['STT'] = range(1, len(df) + 1)
        # df_moi['IDChungTu/MaBill'] = df['Order ID'].astype(str)
        df_moi['IDChungTu/MaBill'] = df['ID ĐƠN HÀNG'].astype(str)
        df_moi['TenHangHoaDichVu'] = df['TÊN SẢN PHẨM']
        # df_moi['TenHangHoaDichVu'] = df['Product Name']
        # df_moi['SoLuong'] = df['Quantity'].astype(int)
        df_moi['SoLuong'] = df['SỐ LƯỢNG'].astype(int)
       


        # Tính đơn vị tính theo tên sản phẩm
        df_moi['DonViTinh/ChietKhau'] = df_moi['TenHangHoaDichVu'].apply(
            lambda val: (
                'chiếc' if 'lược' in str(val).lower()
                else 'chai' if any(x in str(val).lower() for x in ['xịt tinh dầu', 'tinh dầu','nước hoa', 'dầu gội'])
                else 'dây' if any(x in str(val).lower() for x in ['thun buộc tóc', 'dây'])
                else 'hủ' if 'tế bào chết' in str(val).lower()
                else 'hủ'
            )
        )

    #    df_moi['SoLuong'] = df['Quantity']
        df_moi['ThanhTien'] = (df['SKU Subtotal After Discount'] / 1.08).round(0).astype(int)
        df_moi['ThueSuat'] = 0.08
        df_moi['TienThueGTGT'] = (df_moi['ThanhTien'] * 0.08).round(0).astype(int)
        df_moi['DonGia'] = (df_moi['ThanhTien'] / df_moi['SoLuong']).round(0).astype(int)

        # Thêm cột phụ trống
        cot_bo_sung = [
            'NgayThangNamHD', 'CCCD', 'SoHoChieu', 'HoTenNguoiMua', 'TenDonVi',
            'MaNganSach', 'MaSoThue', 'DiaChi', 'SoTaiKhoan', 'HinhThucTT',
            'NhanBangEmail', 'DSEmail', 'NhanBangSMS', 'DSSMS', 'NhanBangBanIN',
            'HoTenNguoiNhan', 'SoDienThoaiNguoiNhan', 'SoNha', 'Tinh/ThanhPho',
            'Huyen/Quan/ThiXa', 'Xa/Phuong/ThiTran', 'GhiChu'
        ]
        for cot in cot_bo_sung:
            df_moi[cot] = ""

        # Sắp xếp lại cột
        cols = [
            'STT', 'IDChungTu/MaBill', 'TenHangHoaDichVu', 'DonViTinh/ChietKhau',
            'SoLuong', 'DonGia', 'ThanhTien', 'ThueSuat', 'TienThueGTGT'
        ] + cot_bo_sung
        df_moi = df_moi[cols]

        # Xuất kết quả vào bộ nhớ
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_moi.to_excel(writer, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="ketqua.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template_string(HTML_FORM)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=10000)
    # app.run(debug=True)
