import os
import sys
import pandas as pd
import requests


# LINE Notify 權杖
token = 'dLF5rbxPQOTzjO2QaaleksNHRL5qzaywjFQZl6LbHVR'

# 获取当前脚本所在的目录
if getattr(sys, 'frozen', False):
    # 如果是以可执行文件形式运行
    script_dir = os.path.dirname(sys.executable)
else:
    # 如果是以脚本形式运行
    script_dir = os.path.dirname(os.path.abspath(__file__))

# 构建 Excel 文件的相對路径
excel_file = os.path.join(script_dir, '測試.xlsm')

try:
    data_frame = pd.read_excel(excel_file, sheet_name=0)

    # 获取列的名称
    columns = data_frame.columns.tolist()

    # 生成对齐的字符串
    message = '\n'
    for index, row in data_frame.iterrows():
        row_message = ''
        for col in columns:
            value = row[col]
            row_message += f"{col}:{str(value)}，"
        row_message = row_message[:-1]  # 移除最后一个逗号
        message += row_message + '\n'

    print(message)

    # HTTP 標頭參數與資料
    headers = {"Authorization": "Bearer " + token}
    data = {'message': message}

    # 以 requests 發送 POST 請求
    response = requests.post("https://notify-api.line.me/api/notify", headers=headers, data=data)
    response.raise_for_status()  # 检查请求是否成功

except Exception as e:
    print("发生错误:", str(e))

# 转换为可执行的EXE文件时添加 __main__ 保护
if __name__ == '__main__':
    pass