import pyodbc
import datetime

# 建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=MRPSDB;UID=IT_Adam;PWD=0eopaf.rk;')

# 创建游标
cursor = connection.cursor()

# 执行 SELECT 查询语句获取特定值
select_query_beforeTime = "SELECT text1 FROM dbo.SctlMast WHERE scrm_key = 'NotifyHR'"
cursor.execute(select_query_beforeTime)
result = cursor.fetchone()
token = result[0] if result else None

# 关闭连接和游标
cursor.close()
connection.close()

# 获取当前时间
current_time = datetime.datetime.now()
formatted_current_time = datetime.datetime.now().strftime("%Y%m%d")

# 重新建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=HRMDB;UID=IT_Adam;PWD=0eopaf.rk;')

# 重新创建游标
cursor = connection.cursor()

# 执行 SELECT 查询
select_query = "SELECT [AttendanceCollectLog].[FileName]" \
               "FROM [AttendanceCollectLog] AS [AttendanceCollectLog]  " \
               "LEFT JOIN [MachineType] AS [AttendanceCollectLog_MachineType_MachineTypeId] " \
               "ON [AttendanceCollectLog].[MachineTypeId]=[AttendanceCollectLog_MachineType_MachineTypeId].[MachineTypeId]  " \
               "LEFT JOIN [Employee] AS [AttendanceCollectLog_Employee_CollectEmployeeId] " \
               "ON [AttendanceCollectLog].[CollectEmployeeId]=[AttendanceCollectLog_Employee_CollectEmployeeId].[EmployeeId]  " \
               "LEFT JOIN [User] AS [AttendanceCollectLog_User_CreateBy] " \
               "ON [AttendanceCollectLog].[CreateBy]=[AttendanceCollectLog_User_CreateBy].[UserId]  " \
               "LEFT JOIN [User] AS [AttendanceCollectLog_User_LastModifiedBy] " \
               "ON [AttendanceCollectLog].[LastModifiedBy]=[AttendanceCollectLog_User_LastModifiedBy].[UserId]  " \
               "WHERE [AttendanceCollectLog].[FileSize] <> 0 " \
               "AND [AttendanceCollectLog].[FileName] = '" + formatted_current_time + ".txt'" \
               "ORDER BY [AttendanceCollectLog].[AttendanceCollectLogId]"

# 执行查询
cursor.execute(select_query)
results = cursor.fetchall()

# 檢查是否有查詢到值
if results:
    # 檢查每一個檔案是否存在
    found_file = False  # 設定一個標誌來表示是否找到檔案

    for row in results:
        filename = row[0]
        expected_filename = formatted_current_time + ".txt"
        if filename == expected_filename:
            found_file = True
            break  # 找到一個就中斷迴圈

    # 根據找到檔案與否印出結果
    if found_file:
        message = '\n' + '宏恩宏聚刷卡' + '\n===============' + '\n今日刷卡數據已匯入HR' + '\n查詢日期: ' + formatted_current_time
    else:
        message = '\n' + '宏恩宏聚刷卡' + '\n===============' + '\n今日刷卡數據未匯入HR' + '\n查詢日期: ' + formatted_current_time
else:
    message = '\n' + '宏恩宏聚刷卡' + '\n===============' + '\n今日刷卡數據未匯入HR' + '\n查詢日期: ' + formatted_current_time

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
cursor.execute(insert_query, 'adam', token, '刷卡', message, 'N', 'N', datetime.datetime.now())

# 提交事务
conn.commit()

# 关闭连接和游标
cursor.close()
connection.close()
