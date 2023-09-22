import pandas as pd

# Đọc dữ liệu từ hai tệp tin Excel vào hai DataFrame
df_yyy = pd.read_excel('yyy.xlsx')
df_xxx = pd.read_excel('xxx.xlsx')

# Tạo một cột G với giá trị mặc định là False trong df_yyy
df_yyy['G'] = False

# Kiểm tra tên trong df_yyy có trong df_xxx không
for index, row in df_yyy.iterrows():
    if row['Email'] in df_xxx['Email'].values:
        df_yyy.at[index, 'G'] = True

# Lưu df_yyy sau khi đã kiểm tra vào tệp tin yyy.xlsx
df_yyy.to_excel('yyy.xlsx', index=False)
