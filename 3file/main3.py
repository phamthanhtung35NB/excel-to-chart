import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from datetime import datetime
import matplotlib.dates as mdates
from matplotlib.gridspec import GridSpec

# Cấu hình font tiếng Việt
plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# Cấu hình style cho biểu đồ
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 8)

# 1. ĐỌC DỮ LIỆU TỪ FILE EXCEL
print("Đang đọc dữ liệu từ các file Excel...")

# Đọc file kết quả bài kiểm tra
df_ket_qua = pd.read_excel('Ket_qua_bai_kiem_tra_18_06_2025.xlsx', header=1)
df_ket_qua.columns = ['ho_ten', 'email', 'gioi_tinh', 'de_thi', 'ket_qua', 'thi_luc', 'thoi_gian_lam_bai', 'ten_khoa_hoc', 'ten_bai_hoc']

# Đọc file danh sách học viên
df_hoc_vien = pd.read_excel('Danh_sach_hoc_vien_tham_gia_18_06_2025.xlsx', header=1)
df_hoc_vien.columns = ['stt', 'ho_ten', 'email', 'trang_thai', 'ngay_dang_ky', 'so_dien_thoai', 'gioi_tinh', 'noi_cong_tac', 'nhom', 'combo_khoa_hoc']

# Đọc file tiến trình học tập
df_tien_trinh = pd.read_excel('tien_trinh_hoc_tap.xlsx', header=1)
df_tien_trinh.columns = ['stt', 'ho_ten', 'email', 'so_dien_thoai', 'nhom', 'khoa_hoc', 'giang_vien', 'thoi_gian_bat_dau', 'thoi_gian_ket_thuc', 'tien_trinh_hoc_tap', 'ket_qua_hoc_tap']

# 2. XỬ LÝ DỮ LIỆU
print("Đang xử lý dữ liệu...")

# Xử lý dữ liệu kết quả thi
df_ket_qua['diem_so'] = df_ket_qua['ket_qua'].str.extract(r'(\d+)/\d+').astype(float)
df_ket_qua['tong_diem'] = df_ket_qua['ket_qua'].str.extract(r'\d+/(\d+)').astype(float)
df_ket_qua['dat_hay_khong'] = df_ket_qua['ket_qua'].str.contains('Đạt')
df_ket_qua['thi_luc'] = pd.to_datetime(df_ket_qua['thi_luc'], format='%d/%m/%Y %H:%M:%S')

# Xử lý dữ liệu tiến trình
df_tien_trinh['tien_trinh_phan_tram'] = df_tien_trinh['tien_trinh_hoc_tap'].str.rstrip('%').astype(float)
df_tien_trinh = df_tien_trinh.dropna(subset=['khoa_hoc'])

# 3. TẠO BIỂU ĐỒ
print("Đang tạo biểu đồ...")

# Tạo figure với nhiều subplot
fig = plt.figure(figsize=(20, 24))
gs = GridSpec(6, 3, figure=fig, hspace=0.3, wspace=0.3)

# Biểu đồ 1: Phân bố điểm thi theo khóa học
ax1 = fig.add_subplot(gs[0, :2])
df_diem_khoa = df_ket_qua.groupby('ten_khoa_hoc')['diem_so'].mean().sort_values(ascending=True)
bars1 = ax1.barh(df_diem_khoa.index, df_diem_khoa.values, color='skyblue', edgecolor='navy')
ax1.set_xlabel('Điểm trung bình')
ax1.set_title('Điểm trung bình theo khóa học', fontsize=14, fontweight='bold')
for bar in bars1:
    width = bar.get_width()
    ax1.text(width, bar.get_y() + bar.get_height()/2, f'{width:.1f}', 
             ha='left', va='center', fontweight='bold')

# Biểu đồ 2: Tỷ lệ đạt/không đạt
ax2 = fig.add_subplot(gs[0, 2])
dat_counts = df_ket_qua['dat_hay_khong'].value_counts()
colors = ['#2ecc71', '#e74c3c']
wedges, texts, autotexts = ax2.pie(dat_counts.values, labels=['Đạt', 'Không đạt'], 
                                    colors=colors, autopct='%1.1f%%', startangle=90)
ax2.set_title('Tỷ lệ đạt/không đạt tổng thể', fontsize=14, fontweight='bold')

# Biểu đồ 3: Phân bố thời gian làm bài
ax3 = fig.add_subplot(gs[1, :])
df_ket_qua['thoi_gian_giay'] = pd.to_timedelta(df_ket_qua['thoi_gian_lam_bai']).dt.total_seconds()
ax3.hist(df_ket_qua['thoi_gian_giay'], bins=20, color='lightcoral', edgecolor='darkred', alpha=0.7)
ax3.set_xlabel('Thời gian làm bài (giây)')
ax3.set_ylabel('Số lượng bài thi')
ax3.set_title('Phân bố thời gian làm bài', fontsize=14, fontweight='bold')
ax3.axvline(df_ket_qua['thoi_gian_giay'].mean(), color='red', linestyle='--', 
            label=f'Trung bình: {df_ket_qua["thoi_gian_giay"].mean():.0f}s')
ax3.legend()

# Biểu đồ 4: Số lượng học viên theo khóa học (từ file tiến trình)
ax4 = fig.add_subplot(gs[2, :2])
khoa_hoc_counts = df_tien_trinh['khoa_hoc'].value_counts()
bars4 = ax4.bar(range(len(khoa_hoc_counts)), khoa_hoc_counts.values, color='lightgreen', edgecolor='darkgreen')
ax4.set_xticks(range(len(khoa_hoc_counts)))
ax4.set_xticklabels(khoa_hoc_counts.index, rotation=45, ha='right')
ax4.set_ylabel('Số lượng học viên')
ax4.set_title('Số lượng học viên đăng ký theo khóa học', fontsize=14, fontweight='bold')
for i, bar in enumerate(bars4):
    height = bar.get_height()
    ax4.text(bar.get_x() + bar.get_width()/2, height, f'{int(height)}',
             ha='center', va='bottom', fontweight='bold')

# Biểu đồ 5: Tiến trình học tập trung bình theo khóa học
ax5 = fig.add_subplot(gs[2, 2])
tien_trinh_tb = df_tien_trinh.groupby('khoa_hoc')['tien_trinh_phan_tram'].mean().sort_values()
ax5.barh(tien_trinh_tb.index, tien_trinh_tb.values, color='gold', edgecolor='orange')
ax5.set_xlabel('Tiến trình trung bình (%)')
ax5.set_title('Tiến trình học tập TB theo khóa', fontsize=14, fontweight='bold')
for i, v in enumerate(tien_trinh_tb.values):
    ax5.text(v + 1, i, f'{v:.1f}%', va='center', fontweight='bold')

# Biểu đồ 6: Phân bố giới tính học viên
ax6 = fig.add_subplot(gs[3, 0])
gioi_tinh_counts = df_hoc_vien['gioi_tinh'].value_counts()
if not gioi_tinh_counts.empty:
    ax6.pie(gioi_tinh_counts.values, labels=gioi_tinh_counts.index, autopct='%1.1f%%',
            colors=['lightblue', 'pink'], startangle=90)
    ax6.set_title('Phân bố giới tính học viên', fontsize=14, fontweight='bold')

# Biểu đồ 7: Số lượng bài thi theo ngày
ax7 = fig.add_subplot(gs[3, 1:])
df_ket_qua['ngay_thi'] = df_ket_qua['thi_luc'].dt.date
bai_thi_theo_ngay = df_ket_qua.groupby('ngay_thi').size()
ax7.plot(bai_thi_theo_ngay.index, bai_thi_theo_ngay.values, marker='o', linewidth=2, markersize=8)
ax7.set_xlabel('Ngày')
ax7.set_ylabel('Số lượng bài thi')
ax7.set_title('Số lượng bài thi theo ngày', fontsize=14, fontweight='bold')
ax7.grid(True, alpha=0.3)
for i, (date, count) in enumerate(bai_thi_theo_ngay.items()):
    ax7.text(date, count, str(count), ha='center', va='bottom', fontweight='bold')

# Biểu đồ 8: Heatmap điểm thi theo giờ trong ngày
ax8 = fig.add_subplot(gs[4, :])
df_ket_qua['gio_thi'] = df_ket_qua['thi_luc'].dt.hour
pivot_table = df_ket_qua.pivot_table(values='diem_so', index='gio_thi', columns='ten_khoa_hoc', aggfunc='mean')
if not pivot_table.empty:
    sns.heatmap(pivot_table.T, cmap='YlOrRd', annot=True, fmt='.1f', ax=ax8, cbar_kws={'label': 'Điểm TB'})
    ax8.set_xlabel('Giờ trong ngày')
    ax8.set_ylabel('Khóa học')
    ax8.set_title('Điểm trung bình theo giờ thi và khóa học', fontsize=14, fontweight='bold')

# Biểu đồ 9: Top học viên thi nhiều nhất
ax9 = fig.add_subplot(gs[5, :2])
top_hoc_vien = df_ket_qua['ho_ten'].value_counts().head(10)
bars9 = ax9.barh(top_hoc_vien.index, top_hoc_vien.values, color='mediumpurple', edgecolor='purple')
ax9.set_xlabel('Số lần thi')
ax9.set_title('Top 10 học viên thi nhiều nhất', fontsize=14, fontweight='bold')
for bar in bars9:
    width = bar.get_width()
    ax9.text(width, bar.get_y() + bar.get_height()/2, f'{int(width)}',
             ha='left', va='center', fontweight='bold')

# Biểu đồ 10: Tỷ lệ hoàn thành khóa học
ax10 = fig.add_subplot(gs[5, 2])
hoan_thanh_counts = df_tien_trinh['ket_qua_hoc_tap'].value_counts()
colors = ['#3498db', '#e74c3c', '#f39c12', '#95a5a6']
wedges, texts, autotexts = ax10.pie(hoan_thanh_counts.values, labels=hoan_thanh_counts.index,
                                     autopct='%1.1f%%', colors=colors[:len(hoan_thanh_counts)])
ax10.set_title('Tỷ lệ hoàn thành khóa học', fontsize=14, fontweight='bold')

plt.suptitle('BÁO CÁO PHÂN TÍCH DỮ LIỆU HỌC TẬP', fontsize=20, fontweight='bold', y=0.98)
plt.tight_layout()

# Lưu biểu đồ
plt.savefig('bao_cao_thong_ke_hoc_tap.png', dpi=300, bbox_inches='tight')
print("Đã lưu biểu đồ vào file: bao_cao_thong_ke_hoc_tap.png")

# 4. TẠO THÊM BIỂU ĐỒ CHI TIẾT
fig2, axes = plt.subplots(3, 2, figsize=(16, 18))
fig2.suptitle('PHÂN TÍCH CHI TIẾT THEO KHÓA HỌC', fontsize=18, fontweight='bold')

# Biểu đồ 11: Box plot điểm số theo khóa học
ax11 = axes[0, 0]
df_ket_qua.boxplot(column='diem_so', by='ten_khoa_hoc', ax=ax11)
ax11.set_title('Phân bố điểm chi tiết theo khóa học')
ax11.set_xlabel('Khóa học')
ax11.set_ylabel('Điểm số')
plt.setp(ax11.xaxis.get_majorticklabels(), rotation=45, ha='right')

# Biểu đồ 12: Scatter plot mối quan hệ thời gian làm bài và điểm
ax12 = axes[0, 1]
scatter = ax12.scatter(df_ket_qua['thoi_gian_giay'], df_ket_qua['diem_so'],
                       c=df_ket_qua['diem_so'], cmap='viridis', alpha=0.6, s=100)
ax12.set_xlabel('Thời gian làm bài (giây)')
ax12.set_ylabel('Điểm số')
ax12.set_title('Mối quan hệ giữa thời gian làm bài và điểm số')
plt.colorbar(scatter, ax=ax12, label='Điểm số')

# Biểu đồ 13: Stacked bar - Kết quả thi theo khóa học
ax13 = axes[1, 0]
result_pivot = pd.crosstab(df_ket_qua['ten_khoa_hoc'], df_ket_qua['dat_hay_khong'])
result_pivot.plot(kind='bar', stacked=True, ax=ax13, color=['#e74c3c', '#2ecc71'])
ax13.set_xlabel('Khóa học')
ax13.set_ylabel('Số lượng')
ax13.set_title('Kết quả thi theo khóa học (Đạt/Không đạt)')
ax13.legend(['Không đạt', 'Đạt'])
plt.setp(ax13.xaxis.get_majorticklabels(), rotation=45, ha='right')

# Biểu đồ 14: Histogram điểm số
ax14 = axes[1, 1]
ax14.hist(df_ket_qua['diem_so'], bins=11, range=(0, 10), color='teal', edgecolor='black', alpha=0.7)
ax14.set_xlabel('Điểm số')
ax14.set_ylabel('Số lượng bài thi')
ax14.set_title('Phân bố điểm số tổng thể')
ax14.axvline(df_ket_qua['diem_so'].mean(), color='red', linestyle='--',
             label=f'TB: {df_ket_qua["diem_so"].mean():.1f}')
ax14.axvline(df_ket_qua['diem_so'].median(), color='green', linestyle='--',
             label=f'Median: {df_ket_qua["diem_so"].median():.1f}')
ax14.legend()

# Biểu đồ 15: Tiến trình học tập theo giảng viên
ax15 = axes[2, 0]
giang_vien_progress = df_tien_trinh.groupby('giang_vien')['tien_trinh_phan_tram'].mean()
bars15 = ax15.bar(giang_vien_progress.index, giang_vien_progress.values, color='coral', edgecolor='darkred')
ax15.set_xlabel('Giảng viên')
ax15.set_ylabel('Tiến trình trung bình (%)')
ax15.set_title('Tiến trình học tập TB theo giảng viên')
for bar in bars15:
    height = bar.get_height()
    ax15.text(bar.get_x() + bar.get_width()/2, height, f'{height:.1f}%',
              ha='center', va='bottom', fontweight='bold')

# Biểu đồ 16: Radar chart - So sánh điểm TB các khóa học
ax16 = fig2.add_subplot(3, 2, 6, projection='polar')
khoa_hoc_diem = df_ket_qua.groupby('ten_khoa_hoc')['diem_so'].mean()
angles = np.linspace(0, 2 * np.pi, len(khoa_hoc_diem), endpoint=False).tolist()
values = khoa_hoc_diem.values.tolist()
values += values[:1]
angles += angles[:1]

ax16.plot(angles, values, 'o-', linewidth=2, color='red')
ax16.fill(angles, values, alpha=0.25, color='red')
ax16.set_xticks(angles[:-1])
ax16.set_xticklabels(khoa_hoc_diem.index, size=8)
ax16.set_ylim(0, 10)
ax16.set_title('So sánh điểm TB các khóa học', pad=20)
ax16.grid(True)

plt.tight_layout()
plt.savefig('phan_tich_chi_tiet_khoa_hoc.png', dpi=300, bbox_inches='tight')
print("Đã lưu biểu đồ chi tiết vào file: phan_tich_chi_tiet_khoa_hoc.png")

# 5. TẠO BÁO CÁO THỐNG KÊ
print("\n" + "="*50)
print("BÁO CÁO THỐNG KÊ TỔNG HỢP")
print("="*50)

print(f"\n1. THỐNG KÊ KẾT QUẢ THI:")
print(f"   - Tổng số bài thi: {len(df_ket_qua)}")
print(f"   - Điểm trung bình: {df_ket_qua['diem_so'].mean():.2f}")
print(f"   - Điểm cao nhất: {df_ket_qua['diem_so'].max()}")
print(f"   - Điểm thấp nhất: {df_ket_qua['diem_so'].min()}")
print(f"   - Tỷ lệ đạt: {(df_ket_qua['dat_hay_khong'].sum() / len(df_ket_qua) * 100):.1f}%")

print(f"\n2. THỐNG KÊ HỌC VIÊN:")
print(f"   - Tổng số học viên đăng ký: {len(df_hoc_vien)}")
print(f"   - Số học viên đã thi: {df_ket_qua['ho_ten'].nunique()}")

print(f"\n3. THỐNG KÊ TIẾN TRÌNH HỌC TẬP:")
print(f"   - Tiến trình trung bình: {df_tien_trinh['tien_trinh_phan_tram'].mean():.1f}%")
print(f"   - Số học viên hoàn thành: {(df_tien_trinh['ket_qua_hoc_tap'] == 'Hoàn thành').sum()}")

print("\n4. THỐNG KÊ THEO KHÓA HỌC:")
for khoa_hoc in df_ket_qua['ten_khoa_hoc'].unique():
    df_khoa = df_ket_qua[df_ket_qua['ten_khoa_hoc'] == khoa_hoc]
    print(f"\n   {khoa_hoc}:")
    print(f"   - Số bài thi: {len(df_khoa)}")
    print(f"   - Điểm TB: {df_khoa['diem_so'].mean():.2f}")
    print(f"   - Tỷ lệ đạt: {(df_khoa['dat_hay_khong'].sum() / len(df_khoa) * 100):.1f}%")

# Hiển thị biểu đồ
plt.show()

print("\nHoàn thành tạo biểu đồ! Đã lưu 2 file:")
print("1. bao_cao_thong_ke_hoc_tap.png - Báo cáo tổng quan")
print("2. phan_tich_chi_tiet_khoa_hoc.png - Phân tích chi tiết")