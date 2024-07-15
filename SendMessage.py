import requests
import pyodbc
import datetime

current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# LINE Notify 權杖
#token = 'dLF5rbxPQOTzjO2QaaleksNHRL5qzaywjFQZl6LbHVR'

# 建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=ODMGDB;UID=USER_MRPS;PWD=CTRLMRPS;')

# 创建游标
cursor = connection.cursor()

# 执行 SELECT 查询
select_query = "SELECT * FROM dbo.Mailfile"
cursor.execute(select_query)

# 获取查询结果
rows = cursor.fetchall()


# 遍历查询结果并打印每一行数据
for row in rows:
    flag = row.flag
    rcvyn = row.rcvyn
    if flag in ['N', 'W']:
        subject = row.subject
        body = row.body
        message = body
        attach = row.attach
        attenter_line = row.attenter_line
        # 拆分API token
        tokens = attenter_line.split(';')
        # 对每个API token 进行操作
        for token in tokens:
            print("Processing API token:", token)

            # HTTP 標頭參數與資料
            headers = {"Authorization": "Bearer " + token}
            data = {'message': message}

            if attach is not None:
                image = open(attach, 'rb')
                files = {'imageFile': image}
                # 以 requests 發送 POST 請求
                response = requests.post("https://notify-api.line.me/api/notify",
                                            headers=headers, data=data, files=files)
            else:
                response = requests.post("https://notify-api.line.me/api/notify",
                                            headers=headers, data=data)

            if response.status_code == 200:
                update_query = "UPDATE dbo.Mailfile SET flag = 'Y', rcvyn = 'Y', [update_date] = ? WHERE flag = 'N' or flag = 'W'"
                cursor.execute(update_query, current_time)
            else:
                update_query = "UPDATE dbo.Mailfile SET flag = 'W', rcvyn = 'Y', [update_date] = ? WHERE flag = 'N'"
                cursor.execute(update_query, current_time)

                # 提交事务
            connection.commit()

# 关闭连接
connection.close()