import pyodbc
import sql_connect

# 建立与 SQL Server 的连接
conn = pyodbc.connect(sql_connect.mssql_MRPSDBself)

# 创建游标
cursor = conn.cursor()

# 执行 DELETE 语句
sql = "DELETE FROM dbo.Mailfile WHERE [flag] = 'N'"
cursor.execute(sql)

# 提交事务
conn.commit()

# 关闭连接
conn.close()