import pandas as pd
import re
import os
from utils import rename_col, save_excel, generate_lf_frame

def calculate_lf_table(xls):
    df = generate_lf_frame(xls, '肺功能检查(lb_lf)', '一秒用力呼气容积（FEV1）')
    df = df.reset_index(drop=False)

    # 创建一个新的DataFrame来存储结果
    new_df = pd.DataFrame()
    for _, row in df.iterrows(): # 遍历所有rows
        # 将当前行添加到新的DataFrame中
        new_df = pd.concat([new_df, row.to_frame().T], ignore_index=True)

        # 初始化新row
        new_row = {'受试者编号': '差值'}
        base_val = None # 记录D1列的值，初始化为None
        for col, val in row.items(): # 遍历该行的所有列值
            if re.match(r'^D1', col):
                # 当前为"D1"列，记录当前值
                if re.match(r'^-?\d+(\.\d+)?$', val):
                    base_val = pd.to_numeric(val)
                continue
            
            if (
                re.match(r'^V\d+', col) and # 当前列名为非首次访视
                base_val != None and # "D1"列值非None
                re.match(r'^-?\d+(\.\d+)?$', val) # 当前列值为数字
            ):
                # 计算
                delta = str(round(pd.to_numeric(val) - base_val, 3))
                new_row[col] = delta
            else:
                new_row[col] = ''
            prev_val = val
        # print(new_row)
        new_row['受试者编号'] = '差值'
        new_df = pd.concat([new_df, pd.Series(new_row).to_frame().T], ignore_index=True)
        
    # 相关列统一转数值
    target_cols = [col for col in new_df.columns if re.match(r'^[DV]\d+', col)]
    new_df[target_cols] = new_df[target_cols].apply(pd.to_numeric)

    # print(new_df)
    new_df = new_df.reset_index(drop=True)

    # 保存结果
    save_excel(new_df, sheet_name = 'FEV1 V(x) 与 D1的差值')