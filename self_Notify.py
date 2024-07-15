import test_secret
import requests
import pyodbc
import sql_connect

# 创建连接
connection_string = sql_connect.mssql_MRPSDB

HP = 'HP'
token = test_secret.line_Notify
message = "123"

# 使用with语句管理连接和游标
with pyodbc.connect(connection_string) as conn:
    with conn.cursor() as cursor:
        # 执行查询
        query = f"SELECT TXT1, TXT2, TXT3, INT1, INT2, INT3 FROM SctlMast WHERE scrm_key = '{HP}202405'"
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
    print(f"查询结果: {horn63106320}、{horn6330}、{horn6340}、{start63106320}、{start6330}、{start6340}")
    # HTTP 標頭參數與資料
    headers = {"Authorization": "Bearer " + token}
    data = {'message': message}

    response = requests.post("https://notify-api.line.me/api/notify",
                             headers=headers, data=data)
else:
    print("没有符合条件的记录")



