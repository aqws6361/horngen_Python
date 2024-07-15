import pyodbc
import datetime

# 建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=MRPSDB;UID=USER_MRPS;PWD=CTRLMRPS;')

# 创建游标
cursor = connection.cursor()

# 执行 SELECT 查询语句获取特定值
select_query_beforeTime = "SELECT TXT3, text1 FROM dbo.SctlMast WHERE scrm_key = 'NOTIFY'"
cursor.execute(select_query_beforeTime)
result = cursor.fetchone()
token = result[1] if result else None
beforeTime = result[0] if result else None

# 获取当前时间
current_time = datetime.datetime.now()
hour = int(current_time.strftime('%H'))

if 6 < hour <= 10:
    timeId = '01'
elif 10 < hour <= 14:
    timeId = '02'
elif 14 < hour <= 18:
    timeId = '03'
else:
    timeId = '04'

# 格式化时间
YYYYMMDD = current_time.strftime('%Y%m%d')
YYYYWW = current_time.strftime('%Y%W')

# 计算时间间隔
delta = datetime.timedelta(hours=int(beforeTime))

# 计算新的时间
new_time = current_time - delta

# 执行 SELECT 查询
select_query = "SELECT " \
               "T.site_id, " \
               "T.Part_no, " \
               "P.part_desc, " \
               "OM.line_id, " \
               "SUM(T.trx_qty) AS '生產量', " \
               "OD.po_um " \
               "FROM dbo.Trnsinvd AS T " \
               "INNER JOIN dbo.Partmast AS P ON T.Part_no = P.part_no " \
               "INNER JOIN ORDRMAST OM ON left(T.order_no,7) = OM.order_no " \
               "INNER JOIN OrdrDetl OD ON OM.order_no = OD.order_no " \
               "WHERE T.date_created >= ? and trx_ctrl='RCPSCH' " \
               "GROUP BY T.site_id, OM.line_id, T.Part_no, P.part_desc, OD.po_um; " \

cursor.execute(select_query, new_time)
# 获取查询结果
rows = cursor.fetchall()

# 关闭连接和游标
cursor.close()
connection.close()

GrossWeight = 0

hnFYmessage = ''
hnFYcontent = ''
hnFYsubject = ''
hnFYbody = ''
hnFYselect_time = ''

hnABmessage = ''
hnABcontent = ''
hnABsubject = ''
hnABbody = ''
hnABselect_time = ''

hjFYmessage = ''
hjFYcontent = ''
hjFYsubject = ''
hjFYbody = ''
hjFYselect_time = ''

hsFYmessage = ''
hsFYcontent = ''
hsFYsubject = ''
hsFYbody = ''
hsFYselect_time = ''

# 建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=HiiPDB;UID=USER_HIIP;PWD=CTRLHIIP;')

# 创建游标
cursor = connection.cursor()

if rows == []:
    # 插入数据
    cursor.execute("""
                INSERT INTO dbo.DalyProdInfo (SiteId, LineId, WeekId, DateId, TimeId, MtrNo, Qty, Uom, CreatDt)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (0, 0, YYYYWW, YYYYMMDD, timeId, 0, 0, 0, current_time))
    # 提交更改
    connection.commit()

# 遍历查询结果并插入到数据库表中
for row in rows:
    sum_qty = row.生產量
    FactoryID = row.site_id
    PartNo = row.Part_no
    Description = row.part_desc
    craftsmanshipCode = row.line_id[1]
    productionLine = row.line_id[-2:]
    Uom = row.po_um

    # 插入数据
    cursor.execute("""
            INSERT INTO dbo.DalyProdInfo (SiteId, LineId, WeekId, DateId, TimeId, MtrNo, Qty, Uom, CreatDt)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (FactoryID, craftsmanshipCode, YYYYWW, YYYYMMDD, timeId, PartNo, sum_qty, Uom, current_time))

    if FactoryID == '1':
        Factory = '1廠 宏恩芳苑廠'
        if craftsmanshipCode == '1':
            craftsmanshipName = '粉碎'
        elif craftsmanshipCode == '2':
            craftsmanshipName = '洗料'
        elif craftsmanshipCode == '3':
            craftsmanshipName = '製粒'
        elif craftsmanshipCode == '4':
            craftsmanshipName = '後拌'
        hnFYsubject = '\n' + '標題:' + Factory + ' 工單報工\n' + '製程:' + craftsmanshipName + '\t產線:第' + productionLine + '線'
        hnFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime(
            '%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hnFYbody = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip(
            '0').rstrip('.') + 'kg'
    elif FactoryID == '2':
        Factory = '2廠 宏恩彰濱廠'
        if craftsmanshipCode == '1':
            craftsmanshipName = '粉碎'
        elif craftsmanshipCode == '2':
            craftsmanshipName = '洗料'
        elif craftsmanshipCode == '3':
            craftsmanshipName = '製粒'
        elif craftsmanshipCode == '4':
            craftsmanshipName = '後拌'
        hnABsubject = '\n' + '標題:' + Factory + ' 工單報工\n' + '製程:' + craftsmanshipName + '\t產線:第' + productionLine + '線'
        hnABselect_time = '\n===============' + '\n起始時間:' + new_time.strftime(
            '%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hnABbody = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip(
            '0').rstrip('.') + 'kg'
    elif FactoryID == '3':
        Factory = '3廠 宏聚芳苑廠'
        if craftsmanshipCode == '1':
            craftsmanshipName = '粉碎'
        elif craftsmanshipCode == '2':
            craftsmanshipName = '洗料'
        elif craftsmanshipCode == '3':
            craftsmanshipName = '製粒'
        elif craftsmanshipCode == '4':
            craftsmanshipName = '後拌'
        hjFYsubject = '\n' + '標題:' + Factory + ' 工單報工\n' + '製程:' + craftsmanshipName + '\t產線:第' + productionLine + '線'
        hjFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime(
            '%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hjFYbody = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip(
            '0').rstrip('.') + 'kg'
    elif FactoryID == '4':
        Factory = '4廠 宏盛芳苑廠'
        if craftsmanshipCode == '1':
            craftsmanshipName = '粉碎'
        elif craftsmanshipCode == '2':
            craftsmanshipName = '洗料'
        elif craftsmanshipCode == '3':
            craftsmanshipName = '製粒'
        elif craftsmanshipCode == '4':
            craftsmanshipName = '後拌'
        hsFYsubject = '\n' + '標題:' + Factory + ' 工單報工\n' + '製程:' + craftsmanshipName + '\t產線:第' + productionLine + '線'
        hsFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime(
            '%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        hsFYbody = '\n===============' + '\n料號:' + PartNo + '\n說明:' + Description + '\n總數:' + str(sum_qty).rstrip(
            '0').rstrip('.') + 'kg'

    hnFYcontent = hnFYcontent + hnFYbody
    hnFYmessage = hnFYsubject + hnFYcontent + hnFYselect_time
    hnABcontent = hnABcontent + hnABbody
    hnABmessage = hnABsubject + hnABcontent + hnABselect_time
    hjFYcontent = hjFYcontent + hjFYbody
    hjFYmessage = hjFYsubject + hjFYcontent + hjFYselect_time
    hsFYcontent = hsFYcontent + hsFYbody
    hsFYmessage = hsFYsubject + hsFYcontent + hsFYselect_time

if hnFYmessage == '':
    hnFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime('%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    hnFYmessage = '\n===============\n此時段無工單' + hnFYselect_time
if hnABmessage == '':
    hnABselect_time = '\n===============' + '\n起始時間:' + new_time.strftime('%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    hnABmessage = '\n===============\n此時段無工單' + hnABselect_time
if hjFYmessage == '':
    hjFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime('%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    hjFYmessage = '\n===============\n此時段無工單' + hjFYselect_time
if hsFYmessage == '':
    hsFYselect_time = '\n===============' + '\n起始時間:' + new_time.strftime('%Y-%m-%d %H:%M:%S') + '\n終止時間:' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    hsFYmessage = '\n===============\n此時段無工單' + hsFYselect_time

# 提交更改
connection.commit()

# 关闭连接和游标
cursor.close()
connection.close()

# 建立新的数据库连接和游标
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.168.53.53;'
                      'Database=ODMGDB;'
                      'UID=USER_MRPS;'
                      'PWD=CTRLMRPS;')
cursor = conn.cursor()

# 执行 INSERT 语句
insert_query = "INSERT INTO dbo.Mailfile (sender, attenter_line, subject, body, flag, rcvyn, create_date) " \
               "VALUES (?, ?, ?, ?, ?, ?, ?)"
cursor.execute(insert_query, 'adam', token, '原料', hnABmessage, 'N', 'N', current_time)

# 提交事务
conn.commit()

# 关闭连接和游标
cursor.close()
conn.close()

# # 測試用
# insert_query = """
# INSERT INTO dbo.Mailfile (sender, attenter_line, subject, body, flag, rcvyn, create_date)
# VALUES (?, ?, ?, ?, ?, ?, ?),
#        (?, ?, ?, ?, ?, ?, ?),
# """
#
# values = [
#     ('adam', token, '原料', message, 'N', 'N', current_time),
#     ('john', token, '成品', message, 'N', 'N', current_time),
#     ...
# ]
#
# cursor.execute(insert_query, *values)