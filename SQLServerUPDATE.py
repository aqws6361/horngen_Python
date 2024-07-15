import pyodbc

# 建立与SQL Server的连接
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.168.53.53;'
                      'Database=MRPSDB;'
                      'UID=USER_MRPS;'
                      'PWD=CTRLMRPS;')

# 创建游标
cursor = conn.cursor()

# 执行UPDATE语句
sql = r"UPDATE dbo.SctlMast SET [text1] = 'v0qE9gzGGyBTNrI9ROXRDIzvwvgDG4WRv1viHCg4bPp' WHERE [scrm_key] = 'NOTIFY'"
#sql = r"UPDATE dbo.Mailfile SET attach = 'C:\Users\10500\Desktop\宏恩集團.png' WHERE [iden_key] = '1'"
cursor.execute(sql)

# 提交事务
conn.commit()

# 关闭连接
conn.close()