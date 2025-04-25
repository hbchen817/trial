import pandas as pd
import re
import os
from utils import rename_col, save_excel, generate_lf_frame, generate_ivm_frame

def calculate_lf_and_ivm(xls):
    df1 = generate_lf_frame(xls, "肺功能检查(lb_lf)", col_name='FEV1实测值/预测值%')
    df2 = generate_lf_frame(xls, "肺功能检查(lb_lf)", col_name='一秒用力呼气容积（FEV1）')
    df = pd.concat([df1, df2], axis=1)

    ivm_df = generate_ivm_frame(xls, sheet_name='吸气流速测定(lb_ivm)', col_name='测定均值（系统计）(L/min)')

    # 往FEV Frame中合入流速测定数据Frame
    df = pd.concat([df, ivm_df], axis=1)
    df = df.reset_index(drop=False)
    # 去除Frame中'V'开头的列
    df = df.loc[:, ~df.columns.str.lower().str.startswith('v')]
    series = df['受试者编号']
    # print(df)
    # exit(0)

    # "人口统计学(dm)" datasheet预处理
    subject_df = pd.read_excel(xls, "人口统计学(dm)")
    subject_df = subject_df.drop(index = 0)
    subject_df.set_index('受试者编号', inplace=True)

    # 增加"性别"列
    id_to_gender = dict(zip(subject_df.index, subject_df['性别']))
    gender_series = series.map(id_to_gender)
    df.insert(loc=1, column='性别', value=gender_series)

    # 增加"年龄"列
    id_to_age = dict(zip(subject_df.index, subject_df['年龄']))
    age_series = series.map(id_to_age)
    df.insert(loc=2, column='年龄', value=age_series)

    # 列名包含如下关键字的列统一转数字
    substrs = ['一秒用力呼气容积（FEV1）', 
            'FEV1实测值/预测值%',
            '流速测定均值(L/min)',
            '年龄']
    target_cols = [col for col in df.columns if any(sub in col for sub in substrs)]
    df[target_cols] = df[target_cols].apply(pd.to_numeric)

    # 过滤掉"age < 50"的受试者
    df = df[df['年龄'] >= 50]

    # 保存结果
    save_excel(df, sheet_name = 'FEV1实测值-预测值 VS 流速 for D1')
