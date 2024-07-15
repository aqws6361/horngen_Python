import pyodbc

# 建立数据库连接
connection = pyodbc.connect('DRIVER={SQL Server};'
                            'SERVER=192.168.53.53;'
                            'DATABASE=MRPSDB;'
                            'UID=USER_MRPS;'
                            'PWD=CTRLMRPS;')

# 创建游标
cursor = connection.cursor()

# 构建插入语句
insert_query = f"INSERT INTO dbo.SctlMast (scrm_key, date_lst_update, TXT1, text1) " \
               f"VALUES ('Factory', GETDATE(), '12')"

# 执行插入语句
cursor.execute(insert_query)

# 提交事务
connection.commit()

# 关闭连接
connection.close()