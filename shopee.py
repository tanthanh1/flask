import pandas as pd

# Đọc file Excel (thay 'ten_file.xlsx' bằng tên file thực tế của bạn)
df_cu = pd.read_excel('input.xlsx')
cot_lay = df_cu[['SKU ID', 'Product Name', 'Quantity',]]

df_moi.columns = ['Full Name', 'Email']

cot_lay.to_excel('file_moi.xlsx', index=False)

# Hiển thị 5 dòng đầu tiên
print(df_cu.head())
