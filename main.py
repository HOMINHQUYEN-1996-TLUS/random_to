
import openpyxl

# Mở tệp Excel chứa 50 tên và dữ liệu tương ứng
workbook = openpyxl.load_workbook('Book1.xlsx')
sheet_goc = workbook.active

# Tạo 5 sheet mới
for i in range(1, 6):
    new_sheet = workbook.create_sheet(title=f'To_{i}')

# Tạo danh sách các dòng không lặp lại và cột A và B tương ứng
danh_sach_dong = list(set((cell[0].value, cell[1].value) for cell in sheet_goc.iter_rows(min_row=1, max_row=50, min_col=1, max_col=2)))

# Phân chia danh sách dòng vào 5 phần bằng số dòng gần bằng nhau
so_dong_moi_sheet = len(danh_sach_dong) // 5

# Bắt đầu phân chia dữ liệu vào các sheet
for i, sheet in enumerate(workbook.sheetnames[1:], start=1):
    current_sheet = workbook[sheet]
    for dong in danh_sach_dong[:so_dong_moi_sheet]:
        current_sheet.append(dong)
    danh_sach_dong = danh_sach_dong[so_dong_moi_sheet:]

# Lưu tệp Excel sau khi đã chỉnh sửa
workbook.save('DANH_SACH_TO.xlsx')
