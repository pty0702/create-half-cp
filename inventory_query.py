# -*- coding: utf-8 -*-
#针对 140302（原材料/自制半成品） 及其下级末级科目进行“数量金额式”汇总

import pyodbc
import pandas as pd
import warnings
from datetime import datetime

# 屏蔽 Pandas 警告
warnings.filterwarnings('ignore', category=UserWarning)

# 1. 数据库配置
DB_CONFIG = {
    "DRIVER": "{SQL Server}",
    "SERVER": "cw02",
    "DATABASE": "UFDATA_001_2026",
    "UID": "Sa",
    "PWD": "ByWenJianGuo2005"
}


def get_inventory_data(target_period=4, ccode_prefix='140302'):
    """
    执行数量金额式查询逻辑
    """
    conn_str = (
        f"DRIVER={DB_CONFIG['DRIVER']};SERVER={DB_CONFIG['SERVER']};"
        f"DATABASE={DB_CONFIG['DATABASE']};UID={DB_CONFIG['UID']};PWD={DB_CONFIG['PWD']};"
    )
    conn = pyodbc.connect(conn_str)

    # 核心 SQL：年初余额 + 期间累计变动
    main_sql = f"""
    SELECT 
        Base.ccode as 科目编码,
        Base.ccode_name as 科目名称,
        (ISNULL(OpenBal.年初金额, 0) + ISNULL(PrevVouch.变动金额, 0)) as 期初金额,
        (ISNULL(OpenBal.年初数量, 0) + ISNULL(PrevVouch.变动数量, 0)) as 期初数量,
        ISNULL(CurrVouch.md, 0) as 本期借方金额,
        ISNULL(CurrVouch.nd_s, 0) as 本期借方数量,
        ISNULL(CurrVouch.mc, 0) as 本期贷方金额,
        ISNULL(CurrVouch.nc_s, 0) as 本期贷方数量
    FROM 
        (SELECT ccode, ccode_name FROM code WHERE ccode LIKE '{ccode_prefix}%' AND bproperty = 1) Base
    LEFT JOIN 
        (SELECT ccode, mb as 年初金额, nb_s as 年初数量 FROM gl_accsum WHERE iperiod = 1) OpenBal 
        ON Base.ccode = OpenBal.ccode
    LEFT JOIN 
        (SELECT ccode, SUM(md - mc) as 变动金额, SUM(nd_s - nc_s) as 变动数量 
         FROM gl_accvouch WHERE iperiod < {target_period} AND (iflag = 0 OR iflag IS NULL)
         GROUP BY ccode) PrevVouch ON Base.ccode = PrevVouch.ccode
    LEFT JOIN 
        (SELECT ccode, SUM(md) as md, SUM(nd_s) as nd_s, SUM(mc) as mc, SUM(nc_s) as nc_s 
         FROM gl_accvouch WHERE iperiod = {target_period} AND (iflag = 0 OR iflag IS NULL)
         GROUP BY ccode) CurrVouch ON Base.ccode = CurrVouch.ccode
    """

    df = pd.read_sql(main_sql, conn)
    conn.close()

    # 计算衍生指标
    df['期末金额'] = df['期初金额'] + df['本期借方金额'] - df['本期贷方金额']
    df['期末数量'] = df['期初数量'] + df['本期借方数量'] - df['本期贷方数量']

    # --- 核心修改点：单价保留 3 位小数 ---
    df['期末单价'] = df.apply(
        lambda x: round(x['期末金额'] / x['期末数量'], 3) if x['期末数量'] != 0 else 0,
        axis=1
    )

    # 过滤掉无数据的行
    df = df[(df['期初数量'] != 0) | (df['本期借方数量'] != 0) | (df['本期贷方数量'] != 0)].copy()

    return df


def export_to_excel(df, period):
    """
    导出为格式化的 Excel 文件
    """
    timestamp = datetime.now().strftime("%H%M%S")
    file_name = f"140302数量金额明细_{period}月_{timestamp}.xlsx"

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='明细账')

        # 获取 worksheet 对象进行样式微调
        worksheet = writer.sheets['明细账']

        # 设置列宽
        worksheet.column_dimensions['A'].width = 15
        worksheet.column_dimensions['B'].width = 30
        for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
            worksheet.column_dimensions[col].width = 15

    print(f"✅ Excel 文件导出成功：{file_name}")


if __name__ == "__main__":
    current_period = 4
    print(f"正在调取 {current_period} 月存货数据并计算 3 位精度单价...")

    data = get_inventory_data(target_period=current_period)

    if not data.empty:
        # 1. 控制台预览（设置显示精度为 3 位）
        print(data.head(10).to_string(index=False, formatters={
            '期末单价': '{:,.3f}'.format,
            '期初金额': '{:,.2f}'.format,
            '期末金额': '{:,.2f}'.format
        }))
        # 2. 导出文件
        export_to_excel(data, current_period)
    else:
        print("未发现匹配的数据。")