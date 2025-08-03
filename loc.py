import pandas as pd

# Đọc file gốc
df = pd.read_excel('input.xlsx')

# Lấy các cột cần thiết
cot_can_lay = ['Order ID', 'Product Name', 'Quantity', 'SKU Subtotal After Discount']
df_moi = df[cot_can_lay]

# Ghi ra file mới
df_moi.to_excel('result.xlsx', index=False)