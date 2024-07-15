import pyodbc
import datetime
import sql_connect

# 建立数据库连接
connection = pyodbc.connect(sql_connect.mssql_MRPSDBself)

# 创建游标
cursor = connection.cursor()

# 执行 SELECT 查询语句获取特定值
select_query_beforeTime = "SELECT TXT1, text1 FROM dbo.SctlMast WHERE scrm_key = 'NOTIFY'"
cursor.execute(select_query_beforeTime)
result = cursor.fetchone()
token = result[1] if result else None
tokens = token.split(";")
token1 = tokens[0]
if len(tokens) > 1:
  token2 = tokens[1]
else:
  token2 = None
beforeTime = result[0] if result else None

# 获取当前时间
current_time = datetime.datetime.now()
# 计算时间间隔
delta = datetime.timedelta(hours=int(beforeTime))
# 计算新的时间
new_time = current_time - delta

# 执行 SELECT 查询
select_query = "SELECT T.site_id, T.Part_no, P.part_desc, " \
               "SUM(T.trx_qty) AS '生產量' " \
               "FROM dbo.Trnsinvd AS T " \
               "INNER JOIN dbo.Partmast AS P ON T.Part_no = P.part_no " \
               "WHERE T.date_created >= ? and trx_ctrl='RCPSCH'" \
               "GROUP BY T.site_id, T.Part_no, P.part_desc;"

cursor.execute(select_query, new_time)
# 获取查询结果
rows = cursor.fetchall()

# 关闭连接和游标
cursor.close()
connection.close()

GrossWeight = 0
subjectFY = ''
subjectZB = ''
messageFY = ''
messageZB = ''
contentFY = ''
contentZB = ''
select_time = '\n===============' + '\n起始時間:' + new_time.strftime('%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
noOrder = '\n===============' + '\n此時段無工單'
# 遍历查询结果并插入到数据库表中
for row in rows:
    sum_qty = row.生產量
    FactoryID = row.site_id
    PartNo = row.Part_no
    Description = row.part_desc
    if FactoryID == '1':
        Factory = '1廠 宏恩芳苑廠 '
        subjectFY = '\n' + '標題:' + Factory + '工單報工'
        bodyFY = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip('0').rstrip('.') + 'kg'
        contentFY = contentFY + bodyFY

    if FactoryID == '2':
        Factory = '2廠 宏恩彰濱廠 '
        subjectZB = '\n' + '標題:' + Factory + '工單報工'
        bodyZB = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip('0').rstrip('.') + 'kg'
        contentZB = contentZB + bodyZB

    messageFY = subjectFY + contentFY + select_time
    messageZB = subjectZB + contentZB + select_time

if contentFY == '':
    messageFY = noOrder + select_time

if contentZB == '':
    messageZB = noOrder + select_time

if messageFY != '' and token2 != None:
    # 建立新的数据库连接和游标
    conn = pyodbc.connect(sql_connect.mssql_MRPSDBself)
    cursor = conn.cursor()
    # 执行 INSERT 语句
    insert_query = "INSERT INTO dbo.Mailfile (sender, attenter_line, subject, body, flag, rcvyn, create_date) " \
                   "VALUES (?, ?, ?, ?, ?, ?, ?)"
    cursor.execute(insert_query, 'adam', token2, '原料', messageFY, 'N', 'N', current_time)
    # 提交事务
    conn.commit()
    # 关闭连接和游标
    cursor.close()
    conn.close()

if messageZB != '':
    # 建立新的数据库连接和游标
    conn = pyodbc.connect(sql_connect.mssql_MRPSDBself)
    cursor = conn.cursor()
    # 执行 INSERT 语句
    insert_query = "INSERT INTO dbo.Mailfile (sender, attenter_line, subject, body, flag, rcvyn, create_date) " \
                   "VALUES (?, ?, ?, ?, ?, ?, ?)"
    cursor.execute(insert_query, 'adam', token1, '原料', messageZB, 'N', 'N', current_time)
    # 提交事务
    conn.commit()
    # 关闭连接和游标
    cursor.close()
    conn.close()