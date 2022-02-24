import datetime
from datetime import date, timedelta
import pandas as pd
import pymssql


class SQL_Func:
    def __init__(self):
        self.conn = pymssql.connect(server='xxxx',
                                    user='xxxx',
                                    password='xxx',
                                    database='xxxx',
                                    # as_dict = True,
                                    port='xxxx')
        self.cursor = self.conn.cursor()

    def get_table(self, sql_query):
        # 打開數劇庫連接
        df = pd.read_sql(sql_query, self.conn)
        return df

    def query(self, sql_query):
        try:
            # sql 查詢
            self.cursor.execute(sql_query)
            # 獲取數據
            latest_day_from_db = list(self.cursor.fetchmany(1)[0])[0]  # use fetch 獲得數據
        except:
            print('Unable to fetch the data')

        cursor.close()  # 關閉指針對象
        conn.close()  # 關閉數劇庫連接
        return latest_day_from_db

    def insert(self, df_):
        for index, row in df_.iterrows():
            try:
                U1 = (row['Date'])
                U2 = row['Currency']
                U3 = (row['Buying_Cash_Rate'])
                U4 = (row['Selling_Cash_Rate'])
                U5 = (row['Buying_Spot_Rate'])
                U6 = (row['Selling_Spot_Rate'])
                print(U1, U2, U3, U4, U5, U6)

                # sql 插入數據
                self.cursor.execute("INSERT INTO dbo.Foreign_Exchange_Rate VALUES (%s, %s, %s, %s, %s, %s)",
                                    (U1, U2, U3, U4, U5, U6))
                self.conn.commit()  # 提交修改到db
                self.cursor.close()  # 關閉指針對象
                self.conn.close()  # 關閉db
            except:
                print('Unable to insert the data to db')

    def delete(self, sql_query):
        # SQL 語句 Can be specified by developer
        try:
            # sql 執行
            self.cursor.execute(sql_query)
            # 提交修改到db
            self.conn.commit()
        except:
            print('Unable to delete the db table data')

        self.cursor.close()  # 關閉指針對象
        self.conn.close()  # 關閉數劇庫連接

    def update(self):
        sql = ''  # SQL 更新db語句
        try:
            # sql 執行
            self.cursor.execute(sql)
            # 提交修改到db
            self.conn.commit()
        except:
            print('Unable to update the db table data')

        self.cursor.close()  # 關閉指針對象
        self.conn.close()  # 關閉數劇庫連接
