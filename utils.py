import os
import pandas as pd

# 列名映射函数：将访视编号(2,3,...)修改为(D1,V3...)
def rename_col(col, suffix):
    if (col == '2'):
        return 'D1' + '\n(' + suffix + ')'
    if (col.isdigit()):
        return 'V' + col + '\n(' + suffix + ')'
    return col

def save_excel(df: pd.DataFrame, sheet_name, file_path='stat.xlsx'):
    if os.path.exists(file_path):
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
def generate_lf_frame(xls, sheet_name, col_name):
    # 预处理
    df = pd.read_excel(xls, sheet_name)
    # 去除首行（英文标题）
    df = df.drop(index = 0)
    # 行过滤
    df = df[(df['访视编号'] != '1') & (df['检查项目'] == col_name)]
    # 列选择
    df = df[['受试者编号', '访视编号', '检查结果']]
    df = df.drop_duplicates()
    df = df.reset_index(drop=True) # 使用默认索引（0,1,2...)
    
    # wide frame（增加‘访视编号值‘对应列）
    df = df.pivot(index='受试者编号', columns='访视编号', values='检查结果')
    df = df.fillna("")

    # 生成新列名dict
    col_map = {col: rename_col(col, col_name) for col in df.columns}
    # 对df应用"列名映射"
    df.rename(columns=col_map, inplace=True)
    # print(df)
    
    return df

def generate_ivm_frame(xls, sheet_name, col_name):
    # 建立'受试者编号'为索引的吸气流速测定Frame
    df = pd.read_excel(xls, sheet_name)
    df = df.drop(index = 0)
    lbres_series = df.groupby(['受试者编号', '访视编号'])[col_name].min()
    df = lbres_series.to_frame() # 从Series创建Frame
    df = df.reset_index(drop=False)
    df = df.pivot(index='受试者编号', columns='访视编号', values=col_name)
    # 生成新列名dict
    col_map = {col: rename_col(col, col_name) for col in df.columns}
    # 对lbivm_df应用"列名映射"
    df.rename(columns=col_map, inplace=True)
    
    return df