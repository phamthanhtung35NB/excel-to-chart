import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Thiết lập font tiếng Việt cho matplotlib
plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial Unicode MS', 'sans-serif']

class ExcelToChart:
    def __init__(self):
        """Khởi tạo class để xử lý Excel và tạo biểu đồ"""
        self.df_progress = None
        self.df_students = None
        
    def load_data(self, progress_file, students_file):
        """
        Load dữ liệu từ 2 file Excel
        
        Args:
            progress_file (str): Đường dẫn file tiến trình học tập
            students_file (str): Đường dẫn file danh sách học viên
        """
        # Đọc file tiến trình học tập
        self.df_progress = pd.read_excel(progress_file, skiprows=1)
        print(f"✅ Đã load file tiến trình học tập: {self.df_progress.shape[0]} dòng")
        
        # Đọc file danh sách học viên  
        self.df_students = pd.read_excel(students_file, skiprows=1)
        print(f"✅ Đã load file danh sách học viên: {self.df_students.shape[0]} dòng")
        
        # Làm sạch dữ liệu
        self._clean_data()
        
    def _clean_data(self):
        """Làm sạch và chuẩn hóa dữ liệu"""
        # Xử lý file tiến trình học tập
        if self.df_progress is not None:
            # Loại bỏ các dòng trống
            self.df_progress = self.df_progress.dropna(subset=['Họ tên'])
            
            # Chuyển đổi tiến trình học tập thành số
            self.df_progress['Tiến trình (%)'] = self.df_progress['Tiến trình học tập'].str.replace('%', '').astype(float)
            
        # Xử lý file danh sách học viên
        if self.df_students is not None:
            # Chuyển đổi ngày đăng ký
            self.df_students['Ngày đăng ký'] = pd.to_datetime(self.df_students['Ngày đăng ký'], format='%d/%m/%Y %H:%M:%S')
            
    def plot_learning_progress(self, figsize=(12, 8)):
        """
        Tạo biểu đồ thống kê tiến trình học tập
        
        Args:
            figsize (tuple): Kích thước biểu đồ
        """
        fig, axes = plt.subplots(2, 2, figsize=figsize)
        fig.suptitle('📊 THỐNG KÊ TIẾN TRÌNH HỌC TẬP', fontsize=16, fontweight='bold')
        
        # 1. Biểu đồ cột: Phân bố tiến trình học tập
        progress_ranges = ['0%', '1-25%', '26-50%', '51-75%', '76-99%', '100%']
        progress_counts = []
        
        for _, row in self.df_progress.iterrows():
            progress = row['Tiến trình (%)']
            if progress == 0:
                progress_counts.append('0%')
            elif 1 <= progress <= 25:
                progress_counts.append('1-25%')
            elif 26 <= progress <= 50:
                progress_counts.append('26-50%')
            elif 51 <= progress <= 75:
                progress_counts.append('51-75%')
            elif 76 <= progress <= 99:
                progress_counts.append('76-99%')
            elif progress == 100:
                progress_counts.append('100%')
        
        progress_df = pd.DataFrame({'Tiến trình': progress_counts})
        progress_summary = progress_df['Tiến trình'].value_counts().reindex(progress_ranges, fill_value=0)
        
        bars = axes[0,0].bar(progress_summary.index, progress_summary.values, 
                            color=['#ff6b6b', '#ffa726', '#ffca28', '#66bb6a', '#42a5f5', '#26c6da'])
        axes[0,0].set_title('📈 Phân bố Tiến trình Học tập')
        axes[0,0].set_ylabel('Số lượng học viên')
        axes[0,0].tick_params(axis='x', rotation=45)
        
        # Thêm số liệu lên cột
        for bar in bars:
            height = bar.get_height()
            if height > 0:
                axes[0,0].text(bar.get_x() + bar.get_width()/2., height + 0.1,
                              f'{int(height)}', ha='center', va='bottom')
        
        # 2. Biểu đồ tròn: Kết quả học tập
        result_counts = self.df_progress['Kết quả học tập'].value_counts()
        colors = ['#ff7979', '#74b9ff', '#00b894']
        wedges, texts, autotexts = axes[0,1].pie(result_counts.values, labels=result_counts.index, 
                                                autopct='%1.1f%%', colors=colors, startangle=90)
        axes[0,1].set_title('🎯 Kết quả Học tập')
        
        # 3. Biểu đồ ngang: Top khóa học phổ biến
        course_counts = self.df_progress['Khóa học'].value_counts().head(10)
        axes[1,0].barh(range(len(course_counts)), course_counts.values, color='#a29bfe')
        axes[1,0].set_yticks(range(len(course_counts)))
        axes[1,0].set_yticklabels([label[:30] + '...' if len(label) > 30 else label 
                                  for label in course_counts.index])
        axes[1,0].set_title('📚 Top 10 Khóa học Phổ biến')
        axes[1,0].set_xlabel('Số lượng học viên')
        
        # 4. Biểu đồ scatter: Tiến trình theo giảng viên
        teacher_progress = self.df_progress.groupby('Giảng viên')['Tiến trình (%)'].agg(['mean', 'count']).reset_index()
        teacher_progress = teacher_progress[teacher_progress['count'] >= 2]  # Chỉ hiện giảng viên có >= 2 học viên
        
        scatter = axes[1,1].scatter(teacher_progress['count'], teacher_progress['mean'], 
                                   s=teacher_progress['count']*20, alpha=0.6, color='#fd79a8')
        axes[1,1].set_xlabel('Số lượng học viên')
        axes[1,1].set_ylabel('Tiến trình trung bình (%)')
        axes[1,1].set_title('👨‍🏫 Hiệu quả Giảng viên')
        
        # Thêm tên giảng viên
        for i, row in teacher_progress.iterrows():
            axes[1,1].annotate(row['Giảng viên'][:15], 
                              (row['count'], row['mean']), 
                              xytext=(5, 5), textcoords='offset points', fontsize=8)
        
        plt.tight_layout()
        plt.show()
        
    def plot_student_registration(self, figsize=(12, 6)):
        """
        Tạo biểu đồ thống kê đăng ký học viên
        
        Args:
            figsize (tuple): Kích thước biểu đồ
        """
        fig, axes = plt.subplots(1, 2, figsize=figsize)
        fig.suptitle('📅 THỐNG KÊ ĐĂNG KÝ HỌC VIÊN', fontsize=16, fontweight='bold')
        
        # 1. Biểu đồ đường: Số lượng đăng ký theo ngày
        daily_registrations = self.df_students['Ngày đăng ký'].dt.date.value_counts().sort_index()
        
        axes[0].plot(daily_registrations.index, daily_registrations.values, 
                    marker='o', linewidth=2, markersize=8, color='#00b894')
        axes[0].set_title('📈 Lượng Đăng ký Theo Ngày')
        axes[0].set_ylabel('Số lượt đăng ký')
        axes[0].tick_params(axis='x', rotation=45)
        axes[0].grid(True, alpha=0.3)
        
        # Thêm số liệu lên điểm
        for x, y in zip(daily_registrations.index, daily_registrations.values):
            axes[0].annotate(f'{y}', (x, y), textcoords="offset points", 
                           xytext=(0,10), ha='center')
        
        # 2. Biểu đồ cột: Trạng thái học viên
        status_counts = self.df_students['Trạng thái'].value_counts()
        bars = axes[1].bar(status_counts.index, status_counts.values, 
                          color=['#00b894', '#e17055', '#fdcb6e'])
        axes[1].set_title('👥 Trạng thái Học viên')
        axes[1].set_ylabel('Số lượng')
        
        # Thêm số liệu lên cột
        for bar in bars:
            height = bar.get_height()
            axes[1].text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        plt.tight_layout()
        plt.show()
        
    def create_summary_report(self):
        """Tạo báo cáo tổng quan"""
        print("="*60)
        print("📋 BÁO CÁO TỔNG QUAN HỌC TẬP")
        print("="*60)
        
        # Thống kê từ file tiến trình học tập
        if self.df_progress is not None:
            total_enrollments = len(self.df_progress)
            avg_progress = self.df_progress['Tiến trình (%)'].mean()
            completed = len(self.df_progress[self.df_progress['Kết quả học tập'] == 'Hoàn thành'])
            completion_rate = (completed / total_enrollments) * 100
            
            print(f"📊 TIẾN TRÌNH HỌC TẬP:")
            print(f"   • Tổng số đăng ký khóa học: {total_enrollments:,}")
            print(f"   • Tiến trình trung bình: {avg_progress:.1f}%")
            print(f"   • Số học viên hoàn thành: {completed}")
            print(f"   • Tỷ lệ hoàn thành: {completion_rate:.1f}%")
            print(f"   • Số khóa học khác nhau: {self.df_progress['Khóa học'].nunique()}")
            print(f"   • Số giảng viên: {self.df_progress['Giảng viên'].nunique()}")
        
        # Thống kê từ file danh sách học viên
        if self.df_students is not None:
            total_students = len(self.df_students)
            active_students = len(self.df_students[self.df_students['Trạng thái'] == 'Hoạt động'])
            
            print(f"\n👥 HỌC VIÊN:")
            print(f"   • Tổng số học viên: {total_students}")
            print(f"   • Học viên đang hoạt động: {active_students}")
            print(f"   • Tỷ lệ hoạt động: {(active_students/total_students)*100:.1f}%")
            
            # Thống kê đăng ký theo thời gian
            latest_date = self.df_students['Ngày đăng ký'].max().strftime('%d/%m/%Y')
            earliest_date = self.df_students['Ngày đăng ký'].min().strftime('%d/%m/%Y')
            print(f"   • Ngày đăng ký gần nhất: {latest_date}")
            print(f"   • Ngày đăng ký sớm nhất: {earliest_date}")
        
        print("="*60)
        
    def export_charts(self, output_dir='charts'):
        """
        Xuất tất cả biểu đồ ra file hình ảnh
        
        Args:
            output_dir (str): Thư mục lưu biểu đồ
        """
        import os
        os.makedirs(output_dir, exist_ok=True)
        
        # Biểu đồ tiến trình học tập
        self.plot_learning_progress()
        plt.savefig(f'{output_dir}/learning_progress.png', dpi=300, bbox_inches='tight')
        print(f"✅ Đã lưu: {output_dir}/learning_progress.png")
        
        # Biểu đồ đăng ký học viên
        self.plot_student_registration()
        plt.savefig(f'{output_dir}/student_registration.png', dpi=300, bbox_inches='tight')
        print(f"✅ Đã lưu: {output_dir}/student_registration.png")

# Sử dụng
def main():
    """Hàm chính để chạy chương trình"""
    
    # Khởi tạo
    chart_maker = ExcelToChart()
    
    # Load dữ liệu (thay đổi đường dẫn file theo máy của bạn)
    progress_file = "tien_trinh_hoc_tap.xlsx"
    students_file = "Danh_sach_hoc_vien_tham_gia_18_06_2025.xlsx"
    
    try:
        # Load và xử lý dữ liệu
        chart_maker.load_data(progress_file, students_file)
        
        # Tạo báo cáo tổng quan
        chart_maker.create_summary_report()
        
        # Tạo các biểu đồ
        print("\n🎨 Đang tạo biểu đồ...")
        chart_maker.plot_learning_progress()
        chart_maker.plot_student_registration()
        
        # Xuất biểu đồ (tùy chọn)
        # chart_maker.export_charts()
        
        print("\n✅ Hoàn thành tất cả biểu đồ!")
        
    except FileNotFoundError as e:
        print(f"❌ Lỗi: Không tìm thấy file - {e}")
        print("💡 Hãy đảm bảo các file Excel đã được đặt trong cùng thư mục với script Python")
    except Exception as e:
        print(f"❌ Lỗi không mong muốn: {e}")

if __name__ == "__main__":
    main()
