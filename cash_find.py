# -*- coding: utf-8 -*-
import pyodbc
from datetime import datetime

# 数据库连接参数
DB_PARAMS = {
    "DRIVER": "{SQL Server}",
    "SERVER": "cw02",
    "DATABASE": "UFDATA_001_2026",  # 自动锁定 2026 年度库
    "UID": "Sa",
    "PWD": "ByWenJianGuo2005"
}


def get_cash_balance_logic(target_period, ccode='1001'):
    """
    基于“真值逻辑”查询指定月份的期末余额
    :param target_period: 目标月份 (如 4)
    :param ccode: 科目编码 (现金默认 1001)
    """
    conn_str = (
        f"DRIVER={DB_PARAMS['DRIVER']};"
        f"SERVER={DB_PARAMS['SERVER']};"
        f"DATABASE={DB_PARAMS['DATABASE']};"
        f"UID={DB_PARAMS['UID']};"
        f"PWD={DB_PARAMS['PWD']};"
    )

    # 核心 SQL：年初余额 + 1月至目标月的累计变动
    sql = """
    SELECT 
        (ISNULL(A.年初金额, 0) + ISNULL(B.期间变动金额, 0)) as 最终期末余额
    FROM 
    (
        -- 取 1001 年初数
        SELECT mb as 年初金额 
        FROM gl_accsum 
        WHERE ccode = ? AND iperiod = 1
    ) A
    FULL JOIN 
    (
        -- 取 1 月至目标月(含)的凭证变动汇总
        SELECT 
            SUM(md - mc) as 期间变动金额
        FROM gl_accvouch
        WHERE ccode = ? AND iperiod <= ? 
          AND (iflag IS NULL OR iflag = 0)
    ) B ON 1=1
    """

    try:
        with pyodbc.connect(conn_str, timeout=5) as conn:
            cursor = conn.cursor()
            # 注意：此处 target_period 传入 4，查询的是 4 月底的余额
            cursor.execute(sql, (ccode, ccode, target_period))
            row = cursor.fetchone()
            return float(row[0]) if row and row[0] is not None else 0
    except Exception as e:
        print(f"查询失败: {e}")
        return None


if __name__ == "__main__":
    # 执行查询
    cash_code = '1001'  # 如果你的现金科目是 100101 等，请在此修改
    balance = get_cash_balance_logic(4, cash_code)

    print(f"\n{'=' * 40}")
    print(f"   北农（海利）2026年4月现金余额校验")
    print(f"{'=' * 40}")
    print(f"查询科目: {cash_code}")
    print(f"截止日期: 2026-04-30")
    if balance is not None:
        print(f"当前账面余额: {balance:,.2f} 元")
    print(f"{'=' * 40}")