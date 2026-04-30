# -*- coding: utf-8 -*-
import pyodbc
import pandas as pd
import warnings

# 屏蔽 Pandas 的 SQLAlchemy 警告
warnings.filterwarnings('ignore', category=UserWarning)

# 数据库配置
DB_PARAMS = {
    "DRIVER": "{SQL Server}",
    "SERVER": "cw02",
    "DATABASE": "UFDATA_001_2026",
    "UID": "Sa",
    "PWD": "ByWenJianGuo2005"
}


def get_1012_full_details(ccode_prefix='1012'):
    conn_str = f"DRIVER={DB_PARAMS['DRIVER']};SERVER={DB_PARAMS['SERVER']};DATABASE={DB_PARAMS['DATABASE']};UID={DB_PARAMS['UID']};PWD={DB_PARAMS['PWD']};"
    conn = pyodbc.connect(conn_str)

    # 1. 获取 2026 年年初余额 (1月期初)
    # 逻辑：取一级科目 1012 的年初数作为起点
    opening_sql = "SELECT SUM(mb) FROM gl_accsum WHERE ccode = ? AND iperiod = 1"
    cursor = conn.cursor()
    cursor.execute(opening_sql, (ccode_prefix,))
    opening_balance = float(cursor.fetchone()[0] or 0)

    # 2. 获取 1-4 月所有级次的凭证分录
    # 通过 LIKE '1012%' 穿透 1-6 级科目
    detail_sql = """
    SELECT 
        v.iperiod as 期间,
        v.dbill_date as 日期,
        v.ino_id as 凭证号,
        v.ccode as 科目编码,
        c.ccode_name as 科目名称,
        v.cdigest as 摘要,
        v.md as 借方,
        v.mc as 贷方
    FROM gl_accvouch v
    LEFT JOIN code c ON v.ccode = c.ccode
    WHERE v.ccode LIKE ? 
      AND v.iperiod BETWEEN 1 AND 4 
      AND (v.iflag IS NULL OR v.iflag = 0)
    ORDER BY v.dbill_date, v.ino_id, v.ccode
    """
    df = pd.read_sql(detail_sql, conn, params=(ccode_prefix + '%',))
    conn.close()

    # 3. 计算动态余额
    current_bal = opening_balance
    balances = []
    for _, row in df.iterrows():
        current_bal += (row['借方'] - row['贷方'])
        balances.append(current_bal)
    df['余额'] = balances

    return opening_balance, df


# 执行查询
opening_bal, full_df = get_1012_full_details()

print(f"\n{'=' * 100}")
print(f"       1012 及其下级科目 (1-6级) 1-4月全量明细账")
print(f"{'=' * 100}")
print(f"2026年年初余额: {opening_bal:,.2f}")
print(f"{'-' * 100}")

if not full_df.empty:
    # 格式化日期
    full_df['日期'] = full_df['日期'].dt.strftime('%m-%d')
    # 调整列顺序方便查看
    display_cols = ['期间', '日期', '凭证号', '科目编码', '科目名称', '摘要', '借方', '贷方', '余额']
    print(full_df[display_cols].to_string(index=False))
else:
    print("1-4 月份未发现相关凭证分录。")

print(f"{'-' * 100}")
final_bal = full_df['余额'].iloc[-1] if not full_df.empty else opening_bal
print(f"截止 4 月末累计余额: {final_bal:,.2f}")
print(f"{'=' * 100}")

# 如果需要导出到 Excel 进一步分析，取消下面行的注释
# full_df.to_excel("1012全级次明细.xlsx", index=False)