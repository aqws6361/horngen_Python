import pyodbc
import datetime
import sql_connect

# 建立数据库连接
connection = pyodbc.connect(sql_connect.mssql_MRPSDBself)

# 创建游标
cursor = connection.cursor()

# 获取当前时间
current_time = datetime.datetime.now()

# 计算时间间隔
delta = datetime.timedelta(hours=4)  # 假设你想查询过去1小时的数据
new_time = current_time - delta

# 执行 SELECT 查询
select_query = "SELECT T.site_id, T.Part_no, P.part_desc, " \
               "SUM(T.trx_qty) AS '生產量' " \
               "FROM dbo.Trnsinvd AS T " \
               "INNER JOIN dbo.Partmast AS P ON T.Part_no = P.part_no " \
               "WHERE T.date_created >= ? AND trx_ctrl='RCPSCH'" \
               "GROUP BY T.site_id, T.Part_no, P.part_desc;"

cursor.execute(select_query, new_time)

# 获取查询结果
rows = cursor.fetchall()

# 遍历查询结果并打印每一行数据
for row in rows:
    print(row)

# 关闭连接
connection.close()