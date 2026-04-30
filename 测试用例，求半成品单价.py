# -*- coding: utf-8 -*-
import pyodbc
from datetime import datetime

# 1. 数据库配置 (后期建议放入 config.py)
DB_CONFIG = {
    "SERVER": "cw02",
    "UID": "Sa",
    "PWD": "ByWenJianGuo2005",
    "DRIVER": "{SQL Server}"
}


def get_true_opening_balance(acc_set, voucher_date_str, ccode):
    """
    测试案例：基于“年初数 + 凭证变动”计算真正的期初结存
    :param acc_set: 账套号，如 '001'
    :param voucher_date_str: 凭证日期，如 '20260430'
    :param ccode: 科目编码
    """
    # 自动解析日期与数据库名
    try:
        dt = datetime.strptime(voucher_date_str.replace("-", ""), "%Y%m%d")
        target_year = dt.year
        target_period = dt.month  # 目标月份（如4月）
        db_name = f"UFDATA_{acc_set}_{target_year}"
    except Exception as e:
        return f"日期解析失败: {e}"

    conn_str = (
        f"DRIVER={DB_CONFIG['DRIVER']};"
        f"SERVER={DB_CONFIG['SERVER']};"
        f"DATABASE={db_name};"
        f"UID={DB_CONFIG['UID']};"
        f"PWD={DB_CONFIG['PWD']};"
    )

    # 你提供的“年初+变动”终极 SQL 逻辑
    sql = """
    SELECT 
        (ISNULL(A.年初金额, 0) + ISNULL(B.期间变动金额, 0)) as 最终期初金额,
        (ISNULL(A.年初数量, 0) + ISNULL(B.期间变动数量, 0)) as 最终期初数量
    FROM 
    (
        -- 取年初数 (1月期初即为年初)
        SELECT mb as 年初金额, nb_s as 年初数量 
        FROM gl_accsum 
        WHERE ccode = ? AND iperiod = 1
    ) A
    FULL JOIN 
    (
        -- 取 1 月至目标月之前的凭证变动汇总
        SELECT 
            SUM(md - mc) as 期间变动金额,
            SUM(nd_s - nc_s) as 期间变动数量
        FROM gl_accvouch
        WHERE ccode = ? AND iperiod < ? 
          AND (iflag IS NULL OR iflag = 0)
    ) B ON 1=1
    """

    try:
        with pyodbc.connect(conn_str, timeout=5) as conn:
            cursor = conn.cursor()
            cursor.execute(sql, (ccode, ccode, target_period))
            row = cursor.fetchone()

            if row:
                amt = float(row[0] or 0)
                qty = float(row[1] or 0)
                price = amt / qty if qty != 0 else 0

                return {
                    "账套": acc_set,
                    "年度": target_year,
                    "目标月份": target_period,
                    "科目": ccode,
                    "期初总金额": round(amt, 2),
                    "期初总数量": round(qty, 4),
                    "推算单价": round(price, 5),
                    "状态": "校验完成"
                }
            return "未找到相关数据"
    except Exception as e:
        return f"数据库连接或查询出错: {e}"


# --- 执行验证 ---
if __name__ == "__main__":
    # 模拟场景：校验 2026年4月 的期初数据
    # 结果应与你的截图完全一致：621,686.91 / 11,660.000
    test_result = get_true_opening_balance(
        acc_set="001",
        voucher_date_str="20260430",
        ccode="140302010105"
    )

    print(f"\n{'=' * 40}")
    print(f"       U8 期初真值取数算法校验")
    print(f"{'=' * 40}")

    if isinstance(test_result, dict):
        for k, v in test_result.items():
            print(f"{k:<10}: {v}")
    else:
        print(test_result)
    print(f"{'=' * 40}")