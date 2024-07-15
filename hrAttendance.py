import pyodbc
import datetime
import sql_connect
from dateutil.relativedelta import relativedelta

# 获取当前时间
current_time = datetime.datetime.now()

# 获取前一个月的时间
previous_month_time = current_time - relativedelta(months=1)

# 提取年、月
year = previous_month_time.strftime('%Y')
month = previous_month_time.strftime('%m')


# 格式化时间
previous_time_YMD = current_time.strftime('%Y-%m-%d')
formatted_time_YM = previous_month_time.strftime('%Y-%m')
current_time_str = current_time.strftime('%Y-%m-%d %H:%M')
EFGP_time_YMD = current_time.strftime('%Y/%m/%d')
YYYYMM = previous_month_time.strftime('%Y%m')

OP = 'OP'
HP = 'HP' + YYYYMM

# 创建连接
connection_string = sql_connect.mssql_MRPSDB

# 使用with语句管理连接和游标
with pyodbc.connect(connection_string) as conn:
    with conn.cursor() as cursor:
        # 执行查询
        query = f"SELECT TXT1, TXT2, TXT3, INT1, INT2, INT3 FROM SctlMast WHERE scrm_key = '{HP}'"
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
else:
    print("没有符合条件的记录")

# 建立数据库连接
connection1 = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=PASDB;UID=IT_Adam;PWD=0eopaf.rk;')
connection2 = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=EFGPTEST;UID=IT_Adam;PWD=0eopaf.rk;')
connection3 = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.53.53;DATABASE=EFGPTEST;UID=IT_Adam;PWD=0eopaf.rk;')

try:
    # 创建游标
    cursor1 = connection1.cursor()
    cursor2 = connection2.cursor()
    cursor3 = connection3.cursor()

    if start63106320 == '1':
        # 插入到第一个目标表 BPM_ctrl
        insert_query_table1 = f"""
            INSERT INTO PASDB.dbo.BPM_ctrl (EmployeeCode, attendanceType, crt_dt, period_id, cnf_ver, status, cnf_time, INVOKERUNITID) 
            SELECT DISTINCT HR.EmployeeCode, HR.attendanceType, '{previous_time_YMD}', '{formatted_time_YM}', HR.Version, '{OP}', '{current_time_str}', HR.INVOKERUNITID 
            FROM PASDB.dbo.HR_attendance HR 
            WHERE (HR.EmployeeCode LIKE '1%' OR HR.EmployeeCode LIKE '2%') 
              AND HR.EmployeeCode NOT LIKE '_A%'
              AND YEAR(HR.attendanceDate) = ? 
              AND MONTH(HR.attendanceDate) = ? 
              AND HR.Version = '{horn63106320}' 
            ORDER BY HR.EmployeeCode;
        """

        # 执行插入操作
        cursor1.execute(insert_query_table1, (year, month))
        connection1.commit()

        # 插入到第二个目标表 BPM0008_python
        insert_query_table2 = f"""
                INSERT INTO BPM0008_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                FROM PASDB.dbo.HR_attendance HR
                LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                ON HR.EmployeeCode = BPM.EmployeeCode
                WHERE (HR.EmployeeCode LIKE '1%' OR HR.EmployeeCode LIKE '2%')
                  AND HR.EmployeeCode NOT LIKE '_A%'
                  AND YEAR(HR.attendanceDate) = ?
                  AND MONTH(HR.attendanceDate) = ?
                  AND HR.attendanceType = '1'
                  AND BPM.status = 'OP'
                  AND HR.Version = '{horn63106320}'
                ORDER BY HR.EmployeeCode;
            """

        # 执行插入操作
        cursor2.execute(insert_query_table2, (year, month))
        connection2.commit()

        # 插入到第三个目标表 BPM0009_python
        insert_query_table3 = f"""
                INSERT INTO BPM0009_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                FROM PASDB.dbo.HR_attendance HR
                LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                ON HR.EmployeeCode = BPM.EmployeeCode
                WHERE (HR.EmployeeCode LIKE '1%' OR HR.EmployeeCode LIKE '2%')
                  AND HR.EmployeeCode NOT LIKE '_A%'
                  AND YEAR(HR.attendanceDate) = ?
                  AND MONTH(HR.attendanceDate) = ?
                  AND HR.attendanceType = '2'
                  AND BPM.status = 'OP'
                  AND HR.Version = '{horn63106320}'
                ORDER BY HR.EmployeeCode;
                """

        # 执行插入操作
        cursor3.execute(insert_query_table3, (year, month))
        connection3.commit()

        update_query_table1 = f"""UPDATE MRPSDB.dbo.SctlMast
            SET INT1 = 0
            WHERE scrm_key = '{HP}'
        """
        # 執行更新操作
        cursor1.execute(update_query_table1)
        connection1.commit()

    if start6330 == '1':
        # 插入到第一个目标表 BPM_ctrl
        insert_query_table1 = f"""
                    INSERT INTO BPM_ctrl (EmployeeCode, attendanceType, crt_dt, period_id, cnf_ver, status, cnf_time, INVOKERUNITID)
                    SELECT DISTINCT HR.EmployeeCode, HR.attendanceType, '{previous_time_YMD}', '{formatted_time_YM}', HR.Version, '{OP}', '{current_time_str}', HR.INVOKERUNITID
                    FROM PASDB.dbo.HR_attendance HR
                    WHERE HR.EmployeeCode LIKE '3%'
                      AND HR.EmployeeCode NOT LIKE '_A%'
                      AND YEAR(HR.attendanceDate) = ?
                      AND MONTH(HR.attendanceDate) = ?
                      AND HR.Version = '{horn6330}'
                    ORDER BY HR.EmployeeCode;
                """

        # 执行插入操作
        cursor1.execute(insert_query_table1, (year, month))
        connection1.commit()

        # 插入到第二个目标表 BPM0008_python
        insert_query_table2 = f"""
                        INSERT INTO BPM0008_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                        SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                        FROM PASDB.dbo.HR_attendance HR
                        LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                        ON HR.EmployeeCode = BPM.EmployeeCode
                        WHERE HR.EmployeeCode LIKE '3%'
                          AND HR.EmployeeCode NOT LIKE '_A%'
                          AND YEAR(HR.attendanceDate) = ?
                          AND MONTH(HR.attendanceDate) = ?
                          AND HR.attendanceType = '1'
                          AND BPM.status = 'OP'
                          AND HR.Version = '{horn6330}'
                        ORDER BY HR.EmployeeCode;
                    """

        # 执行插入操作
        cursor2.execute(insert_query_table2, (year, month))
        connection2.commit()

        # 插入到第三个目标表 BPM0009_python
        insert_query_table3 = f"""
                        INSERT INTO BPM0009_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                        SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                        FROM PASDB.dbo.HR_attendance HR
                        LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                        ON HR.EmployeeCode = BPM.EmployeeCode
                        WHERE HR.EmployeeCode LIKE '3%'
                          AND HR.EmployeeCode NOT LIKE '_A%'
                          AND YEAR(HR.attendanceDate) = ?
                          AND MONTH(HR.attendanceDate) = ?
                          AND HR.attendanceType = '2'
                          AND BPM.status = 'OP'
                          AND HR.Version = '{horn6330}'
                        ORDER BY HR.EmployeeCode;
                        """

        # 执行插入操作
        cursor3.execute(insert_query_table3, (year, month))
        connection3.commit()

        update_query_table1 = f"""UPDATE MRPSDB.dbo.SctlMast
                    SET INT2 = 0
                    WHERE scrm_key = '{HP}'
                """
        # 執行更新操作
        cursor1.execute(update_query_table1)
        connection1.commit()

    if start6340 == '1':
        # 插入到第一个目标表 BPM_ctrl
        insert_query_table1 = f"""
                            INSERT INTO BPM_ctrl (EmployeeCode, attendanceType, crt_dt, period_id, cnf_ver, status, cnf_time, INVOKERUNITID)
                            SELECT DISTINCT HR.EmployeeCode, HR.attendanceType, '{previous_time_YMD}', '{formatted_time_YM}', HR.Version, '{OP}', '{current_time_str}', HR.INVOKERUNITID
                            FROM PASDB.dbo.HR_attendance HR
                            WHERE HR.EmployeeCode LIKE '4%'
                              AND HR.EmployeeCode NOT LIKE '_A%'
                              AND YEAR(HR.attendanceDate) = ?
                              AND MONTH(HR.attendanceDate) = ?
                              AND HR.Version = '{horn6340}'
                            ORDER BY HR.EmployeeCode;
                        """

        # 执行插入操作
        cursor1.execute(insert_query_table1, (year, month))
        connection1.commit()

        # 插入到第二个目标表 BPM0008_python
        insert_query_table2 = f"""
                                INSERT INTO BPM0008_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                                SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                                FROM PASDB.dbo.HR_attendance HR
                                LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                                ON HR.EmployeeCode = BPM.EmployeeCode
                                WHERE HR.EmployeeCode LIKE '4%'
                                  AND HR.EmployeeCode NOT LIKE '_A%'
                                  AND YEAR(HR.attendanceDate) = ?
                                  AND MONTH(HR.attendanceDate) = ?
                                  AND HR.attendanceType = '1'
                                  AND BPM.status = 'OP'
                                  AND HR.Version = '{horn6340}'
                                ORDER BY HR.EmployeeCode;
                            """

        # 执行插入操作
        cursor2.execute(insert_query_table2, (year, month))
        connection2.commit()

        # 插入到第三个目标表 BPM0009_python
        insert_query_table3 = f"""
                                INSERT INTO BPM0009_python (INVOKERID, INVOKERUNITID, txtAppDate, txtPersonID, txtPersonName, txtIssuerDeptId)
                                SELECT DISTINCT HR.EmployeeCode, HR.INVOKERUNITID, '{EFGP_time_YMD}', HR.EmployeeCode, HR.EmployeeCnName, HR.INVOKERUNITID
                                FROM PASDB.dbo.HR_attendance HR
                                LEFT JOIN PASDB.dbo.BPM_ctrl BPM
                                ON HR.EmployeeCode = BPM.EmployeeCode
                                WHERE HR.EmployeeCode LIKE '4%'
                                  AND HR.EmployeeCode NOT LIKE '_A%'
                                  AND YEAR(HR.attendanceDate) = ?
                                  AND MONTH(HR.attendanceDate) = ?
                                  AND HR.attendanceType = '2'
                                  AND BPM.status = 'OP'
                                  AND HR.Version = '{horn6340}'
                                ORDER BY HR.EmployeeCode;
                                """

        # 执行插入操作
        cursor3.execute(insert_query_table3, (year, month))
        connection3.commit()

        update_query_table1 = f"""UPDATE MRPSDB.dbo.SctlMast
                            SET INT3 = 0
                            WHERE scrm_key = '{HP}'
                        """
        # 執行更新操作
        cursor1.execute(update_query_table1)
        connection1.commit()

    print("Transactions committed.")

except Exception as e:
    print(f"An error occurred: {e}")
    connection1.rollback()
    connection2.rollback()
    connection3.rollback()
finally:
    # 关闭游标和连接
    if cursor1:
        cursor1.close()
    if connection1:
        connection1.close()
    if cursor2:
        cursor2.close()
    if connection2:
        connection2.close()
    if cursor3:
        cursor3.close()
    if connection3:
        connection3.close()