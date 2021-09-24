import os
import sys
import sqlite3
import pandas as pd
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from utility.static import now
from utility.setting import db_tick


con = sqlite3.connect(db_tick)
df_name = pd.read_sql("SELECT name FROM sqlite_master WHERE TYPE = 'table'", con)
table_list = list(df_name['name'].values)
if 'moneytop' in table_list:
    table_list.remove('moneytop')
if 'codename' in table_list:
    table_list.remove('codename')

count = len(table_list)
for i, code in enumerate(table_list):
    df = pd.read_sql(f"SELECT * FROM '{code}'", con)
    df = df.set_index('index')
    if '매도2호가' in df.columns:
        df.rename(columns={
            '매도2호가': '매도호가2', '매도1호가': '매도호가1', '매수1호가': '매수호가1', '매수2호가': '매수호가2',
            '매도2잔량': '매도잔량2', '매도1잔량': '매도잔량1', '매수1잔량': '매수잔량1', '매수2잔량': '매수잔량2'
        }, inplace=True)
        df.to_sql(code, con, if_exists='replace', chunksize=1000)
    print(f'[{now()}] 틱데이터 칼럼명 확인 및 변경 중 ... {i+1}/{count}')
con.close()
