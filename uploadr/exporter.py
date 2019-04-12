# !/usr/bin/env python3
# @Description : python读取Oracle数据库数据到Excel

import os
import cx_Oracle
import xlsxwriter
import datetime
# from ora_settings import ORA_CONNECTION, ORA_USERNAME, ORA_PASSWORD, SQL

ORA_CONNECTION = '192.168.201.13:1521/PDBORCL'
ORA_USERNAME = 'UT_BHH_HSJB'
ORA_PASSWORD = 'UT_BHH_HSJB'

SQL = "SMM_RSRC^^^SELECT A.* FROM SMM_RSRC A INNER JOIN SMM_RSRC_PIC_DIV B ON A.RSRC_SEQ = B.RSRC_SEQ INNER JOIN (SELECT MAX(AA.RSRC_APLY_PRD_ST_DT) AS RSRC_APLY_PRD_ST_DT FROM SMM_RSRC AA WHERE AA.RSRC_CD = '0000123' AND AA.DEL_FLG = '0') C ON A.RSRC_APLY_PRD_ST_DT = C.RSRC_APLY_PRD_ST_DT WHERE A.RSRC_CD = '0000234' AND A.DEL_FLG = '0' AND B.DEL_FLG = '0' ORDER BY A.RSRC_CD ASC  $$$SMM_RSRC_PIC_DIV^^^SELECT B.* FROM SMM_RSRC A INNER JOIN SMM_RSRC_PIC_DIV B ON A.RSRC_SEQ = B.RSRC_SEQ INNER JOIN (SELECT MAX(AA.RSRC_APLY_PRD_ST_DT) AS RSRC_APLY_PRD_ST_DT FROM SMM_RSRC AA WHERE AA.RSRC_CD = '0000123' AND AA.DEL_FLG = '0') C ON A.RSRC_APLY_PRD_ST_DT = C.RSRC_APLY_PRD_ST_DT WHERE A.RSRC_CD = '0000234' AND A.DEL_FLG = '0' AND B.DEL_FLG = '0' ORDER BY A.RSRC_CD ASC  "

# 防止读取结果出现中文乱码（问号）情况
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'


def get_oracle(connection, username, password):
    """
    connection: 'IP:PORT/SID'
        type: str
    username:
        type: str
    password:
        type: str
    rtype: cx_Oracle.connection object
    """
    try:
        print("username: {username}\npassword: {password}\nconnection: {connection}")
        conn = cx_Oracle.connect(username, password, connection)
        return conn
    except cx_Oracle.Error as e:
        exit(e)


def write_excel(myworkbook, fields, contents, sheet_name):
    """
    filename: '20190101.xlsx'
        type: str
    fields:
        type: list
    contents:
        type: list
    """
    # 标题栏格式定义
    format_title = myworkbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter'
    })
    # 内容区格式定义
    format_content = myworkbook.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })

    # sheet1名称定义
    sheet = myworkbook.add_worksheet(sheet_name)

    # 标题栏冻结
    sheet.freeze_panes(1, 0)

    # 标题栏
    for field in range(len(fields)):
        sheet.write(0, field, fields[field][0], format_title)

    # 内容区
    for row in range(len(contents)):
        for col in range(len(fields)):
            ceil = contents[row][col]
            if ceil is not None:
                sheet.write(row + 1, col, str(ceil), format_content)
    print("已保存至{filename}")

def main():
    print('-----------------start------------------')
    # 设置文件名称
    now = datetime.datetime.now()
    m_file = now.strftime('%Y%m%d%H%M%S') + '.xlsx'

    # 获取数据库连接
    conn = get_oracle(ORA_CONNECTION, ORA_USERNAME, ORA_PASSWORD)

    myworkbook = xlsxwriter.Workbook(m_file)
    sql_list = SQL.split('$$$')

    for sql_one in sql_list:

        sheet_name = sql_one.split('^^^')[0]
        strsql = sql_one.split('^^^')[1]

        # 查询数据库
        with conn.cursor() as cursor:
            try:
                result = cursor.execute(strsql)
                print('---------------config-end---------------')
                print(result.statement)
                print('-----------------SQL-end----------------')
                xlsx_fields = cursor.description
                xlsx_contents = cursor.fetchall()
            except cx_Oracle.DatabaseError as e:
                conn.close()
                exit(e.args[0])

            # 往Excel写数据
            write_excel(myworkbook, xlsx_fields, xlsx_contents, sheet_name)

            print('共有{str(len(xlsx_contents))}行，{str(len(result.description))}列')
            print('-------------------end------------------')

    myworkbook.close()
    conn.close()

if __name__ == '__main__':
    main()

# ora_settings.py
# ----------------------------------------
# ORA_CONNECTION = '127.0.0.1:1521/orcl'
# ORA_USERNAME = 'user'
# ORA_PASSWORD = 'password'
#
# SQL = 'select * from user_tables'

