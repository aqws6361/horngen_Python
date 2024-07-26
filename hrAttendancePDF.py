import os
import pyodbc
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
import datetime
from dateutil.relativedelta import relativedelta
from decimal import Decimal
import sys
import sql_connect

# 获取并打印当前工作目录
print("Current Working Directory:", os.getcwd())

# 获取脚本所在目录
def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder和 stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 获取当前时间
current_time = datetime.datetime.now()

# 获取前一个月的时间
previous_month_time = current_time - relativedelta(months=2)

# 提取前一个月的年、月
year = previous_month_time.strftime('%Y')
month = previous_month_time.strftime('%m')

# 格式化时间
formatted_time_YM = current_time.strftime('%Y/%m')
current_time_str = current_time.strftime('%Y-%m-%d %H:%M')
previous_time_YM = previous_month_time.strftime('%Y/%m')
YYYYMM = previous_month_time.strftime('%Y%m')

OP = 'OP'
HP = 'HP' + YYYYMM

# 创建连接
connection_string = sql_connect.mssql_MRPSDB

# 使用with语句管理连接和游标
with pyodbc.connect(connection_string) as conn:
    with conn.cursor() as cursor:
        # 执行查询
        query = f"SELECT TXT1, TXT2, TXT3, INT1, INT2, INT3 FROM SctlMast WHERE scrm_key = '{HP}'"
        cursor.execute(query)

        # 获取查询结果
        result = cursor.fetchone()

# 检查结果并存储到变量中
if result:
    horn63106320 = result[0]
    horn6330 = result[1]
    horn6340 = result[2]
    start63106320 = str(result[3])
    start6330 = str(result[4])
    start6340 = str(result[5])
else:
    print("没有符合条件的记录")

# 建立数据库连接
connection = pyodbc.connect(sql_connect.mssql_PASDB)

# 执行查询
query = f"""
    SELECT HR.EmployeeCode, HR.EmployeeCnName, HR.AttendanceRankName, FORMAT(HR.BeginTime, 'yyyy-MM-dd HH:mm') AS BeginTime,
           FORMAT(HR.EndTime, 'yyyy-MM-dd HH:mm') AS EndTime, HR.Hours, HR.Version, HR.attendanceType
    FROM HR_attendance HR
    WHERE (HR.EmployeeCode LIKE '1A%' OR HR.EmployeeCode LIKE '2A%')
      AND YEAR(HR.attendanceDate) = ?
      AND MONTH(HR.attendanceDate) = ?
      AND HR.attendanceType = '1'
      AND HR.Version = {horn63106320}
    ORDER BY HR.EmployeeCode, HR.BeginTime;
"""

with connection.cursor() as cursor:
    cursor.execute(query, (year, month))
    columns = [column[0] for column in cursor.description]
    data = cursor.fetchall()

# 将 pyodbc.Row 转换为 DataFrame
df = pd.DataFrame.from_records(data, columns=columns)

# 保存数据到 Excel 文件
excel_file = f'hrAttendanceReport_{YYYYMM}.xlsx'
df.to_excel(excel_file, index=False)
print(f"Excel file saved as {excel_file}")

# 创建PDF文件，指定文件保存路径
pdf_file = f'hrAttendancePDF_{YYYYMM}.pdf'
pdf = canvas.Canvas(pdf_file, pagesize=letter)
width, height = letter

# 加载中文字体文件
chinese_font_path = resource_path('C:/Users/10500/Desktop/字體/msjhl.ttc')  # 替换为你的中文字体文件路径
pdfmetrics.registerFont(TTFont('ChineseFont', chinese_font_path))

# 设置标题
pdf.setTitle("Attendance Report")

# 定义表头和数据字体
header_font = "ChineseFont"  # 使用加载的中文字体
data_font = "ChineseFont"    # 使用加载的中文字体

# 函数：绘制标题和基本信息
def draw_header(pdf, title, previous_time_YM, current_time_str, employee_code=None, employee_name=None, Version=None):
    pdf.setFont(header_font, 20)
    pdf.drawCentredString(width / 2.0, height - 30, title)
    pdf.setFont(header_font, 14)
    # 计算文本宽度以将其置中
    attendance_month_text = f"考勤月份:{previous_time_YM}"
    attendance_month_width = pdf.stringWidth(attendance_month_text, header_font, 14)
    pdf.drawString((width - attendance_month_width) / 2.0, height - 50, attendance_month_text)

    pdf.line(30, height - 75, width - 30, height - 75)
    if employee_code and employee_name:
        pdf.setFont(header_font, 12)
        pdf.drawString(400, height - 70, f"製表時間:{current_time_str}")
        pdf.drawString(30, height - 70, f"工號: {employee_code}")
        pdf.drawString(110, height - 70, f"姓名: {employee_name}")
        pdf.drawString(350, height - 70, f"版本: {Version}")

# 函数：绘制表头
def draw_table_header(pdf):
    pdf.setFont(header_font, 12)
    pdf.drawString(30, height - 90, "考勤班別")
    pdf.drawString(200, height - 90, "開始時間")  # 调整 x 坐标
    pdf.drawString(330, height - 90, "結束時間")  # 调整 x 坐标
    pdf.drawString(450, height - 90, "小時")      # 调整 x 坐标
    pdf.line(30, height - 95, width - 30, height - 95)

# 函数：绘制签名和日期栏位
def draw_signature_date(pdf):
    pdf.line(30, 70, width - 30, 70)
    pdf.drawString(400, 50, "簽名:__________________")
    pdf.drawString(400, 30, "日期:__________________")

# 初始设置标题和基本信息
draw_header(pdf, "考勤確認單", previous_time_YM, current_time_str)

# 设置数据
pdf.setFont(data_font, 12)
y = height - 130  # 调整初始 y 坐标
line_height = 20
prev_employee_code = None  # 记录上一行的工号
employee_list = []

for index, row in df.iterrows():
    if prev_employee_code is None or row['EmployeeCode'] != prev_employee_code:
        if prev_employee_code is not None:  # 新工号时分页
            pdf.showPage()
        draw_header(pdf, "考勤確認單", previous_time_YM, current_time_str, row['EmployeeCode'], row['EmployeeCnName'], row['Version'])
        draw_table_header(pdf)
        draw_signature_date(pdf)
        pdf.setFont(data_font, 12)
        y = height - 110  # 重新设置初始 y 坐标
        employee_list.append(f"{row['EmployeeCode']} - {row['EmployeeCnName']}")  # 添加到员工列表

    pdf.drawString(30, y, str(row['AttendanceRankName']))
    pdf.drawString(200, y, str(row['BeginTime']))  # 调整 x 坐标
    pdf.drawString(330, y, str(row['EndTime']))    # 调整 x 坐标
    pdf.drawString(450, y, str(row['Hours']))      # 调整 x 坐标
    y -= line_height

    if y < 50:  # 处理分页
        pdf.showPage()
        draw_header(pdf, "考勤確認單", previous_time_YM, current_time_str, row['EmployeeCode'], row['EmployeeCnName'])
        draw_table_header(pdf)
        draw_signature_date(pdf)
        draw_signature_date(pdf)
        pdf.setFont(data_font, 12)
        y = height - 110  # 重新设置初始 y 坐标

    prev_employee_code = row['EmployeeCode']

# 在最後一頁添加簽名和日期欄位
if y > 50:  # 當前頁還有空間時
    pdf.line(30, 70, width - 30, 70)
    pdf.drawString(400, 50, "簽名:__________________")
    pdf.drawString(400, 30, "日期:__________________")
else:  # 當前頁空間不足，添加新頁
    pdf.showPage()
    draw_header(pdf, "考勤確認單", previous_time_YM, current_time_str)
    draw_table_header(pdf)
    draw_signature_date(pdf)
    pdf.line(30, 70, width - 30, 70)
    pdf.drawString(400, 50, "簽名:__________________")
    pdf.drawString(400, 30, "日期:__________________")

# 添加一个新页面用于员工清单和确认栏位
pdf.showPage()
pdf.setFont(header_font, 16)
pdf.drawCentredString(width / 2.0, height - 30, "報表人員清單")

# 设置表格头部
pdf.setFont(header_font, 12)
pdf.drawString(30, height - 60, "工號")
pdf.drawString(200, height - 60, "姓名")
pdf.drawString(400, height - 60, "確認")

# 画表头的横线
pdf.line(30, height - 65, width - 62, height - 65)

# 绘制表格
pdf.setFont(data_font, 12)
y = height - 80
line_height = 20

for employee in employee_list:
    employee_code, employee_name = employee.split(' - ')
    pdf.drawString(30, y, employee_code)
    pdf.drawString(200, y, employee_name)

    # 画每行的横线
    pdf.line(30, y - 5, width - 62, y - 5)
    # 画每行的直线
    pdf.line(30, y + 15, 30, y - 5)
    pdf.line(180, y + 15, 180, y - 5)
    pdf.line(380, y + 15, 380, y - 5)
    pdf.line(550, y + 15, 550, y - 5)

    y -= line_height

    # 检查是否需要分页
    if y < 50:
        pdf.showPage()
        pdf.setFont(header_font, 16)
        pdf.drawCentredString(width / 2.0, height - 30, "報表人員清單")
        pdf.setFont(header_font, 12)
        pdf.drawString(30, height - 60, "工號")
        pdf.drawString(200, height - 60, "姓名")
        pdf.drawString(400, height - 60, "確認")
        pdf.line(30, height - 65, width - 62, height - 65)
        pdf.setFont(data_font, 12)
        y = height - 80

# 保存PDF文件
pdf.save()

print(f"PDF file saved as {pdf_file}")
