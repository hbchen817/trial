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
            df.to_excel(writer, sheet_name=sheet_name, index=False, float_format='%.3f')
    else:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False, float_format='%.3f')
            
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

def generate_pef_frame(xls, sheet_name):
    # 预处理
    df = pd.read_excel(xls, sheet_name, header=1)
    # 去除首行
    #df = df.drop(index = 0)
    # 选取：（1）筛选期的数据 且 （2）CSN是合法值的数据
    # print(df['VISITOID'].dtype) # Int64
    # print(df['CSN'].dtype)  # float64
    # print(df['BEFSTDAT1'].dtype)
    # print(df['BEFRES1'].dtype)
    # print(df['BEFSTDAT2'].dtype)
    # print(df['BEFRES2'].dtype)
    df['CSN'] = df['CSN'].astype('Int64')
    
    # 获取实验期数
    visitoid_set = set(df['VISITOID'].drop_duplicates().reset_index(drop=True))
    # print(visitoid_set)
    visitoid_set.difference_update({1,2})
                    
    temp_df = df[(df['VISITOID'] == 2) & df['CSN'].notna() & df['BEFRES1'].notna()] # 筛选出D1且'CSN'合法的行
    
    grouped = temp_df.groupby(['SUBJID', 'VISITOID'])['CSN'].agg(set).reset_index()
    target_csn = { 1,2,3,4,5,6,7 }
    grouped = grouped[grouped['CSN'].apply(lambda x: target_csn.issubset(x))]
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)
    # print(grouped)
    # return
    
    # 根据D1阶段测量的天数判断有资格参加统计的病患列表
    subid_set = set(grouped['SUBJID'].drop_duplicates().reset_index(drop=True))
    
    # 获取D1阶段的最终有效数据
    temp_df = df[df['SUBJID'].isin(subid_set) & (df['VISITOID'] == 2) & df['BEFRES1'].notna()]
    grouped = temp_df.groupby(['SUBJID', 'VISITOID']).apply(lambda x: x.sort_values('CSN')[-7:]).reset_index(drop=True)
    # temp_df = grouped[['SUBJID', 'VISITOID', 'CSN', 'BEFRES1']] # 治疗期D1的数据
    temp_df = grouped[['SUBJID', 'VISITOID', 'BEFRES1']] # 治疗期D1的数据
    df_final = temp_df.groupby(['SUBJID', 'VISITOID'])[['BEFRES1']].mean()
    # print(df_final)
    # return
    
    for s in subid_set:  # 合格受试者遍历
        for v in visitoid_set: # 合格实验期数遍历
            temp_df = df[(df['SUBJID'] == s) & (df['VISITOID'] == v) & df['BEFRES1'].notna()]
            target_csn = {1,2,3,4,5,6,7,8,9,10,11}
            grouped = temp_df.groupby(['SUBJID', 'VISITOID'])['CSN'].agg(set).reset_index()
            grouped = grouped[grouped['CSN'].apply(lambda x: target_csn.issubset(x))]
              
            if not grouped.empty:
                temp_df = temp_df[(temp_df['SUBJID'] == s) & (temp_df['VISITOID'] == v)]
                temp_df = temp_df[['SUBJID', 'VISITOID', 'BEFRES1']]
                temp_df = temp_df.groupby(['SUBJID', 'VISITOID'])[['BEFRES1']].mean()
                # print(temp_df)
                df_final = pd.concat([df_final, temp_df], ignore_index=False)
   
    df_final = df_final.reset_index(drop=False)
    # print(df_final.index)
    df_final = df_final.sort_values(by=['SUBJID', 'VISITOID'])
    df_final = df_final.reset_index(drop=True)
    # print(df_final)
    
    df_final = df_final.pivot(index='SUBJID', columns='VISITOID', values='BEFRES1')
    # df_final = df_final.fillna('')
    # print(df_final)
    return df_final
    
    
    
    
    df = df[(df['VISITOID'] != 1) & (df['CSN'].notna())]
    
    # 选取:（1）第一期治疗的后7天的数据 或（2）其他治疗期的所有数据
   
    df = df[((df['VISITOID'] == 2) & (df['CSN'].between(8,14))) | (df['VISITOID'] != 2)]
    
    # 选取：
    df = df[(df['VISITOID'] == 2) & (df['CSN'].between(8,14))]
    # print(df)
    
    # 列选择
    df = df[['SUBJID', 'VISITOID', 'CSN', 'BEFSTDAT1', 'BEFRES1']]
    df = df.drop_duplicates()
    df = df.reset_index(drop=True) # 使用默认索引
    print(df)
    
    # wide frame（增加'测量日期'列）
    # df = df.pivot(index='SUBJID', columns='BEFSTDAT1', values='BEFRES1')
    
    return df
    

def generate_ivm_frame(xls, sheet_name, col_name):
    # 预处理
    df = pd.read_excel(xls, sheet_name)
    # 去除首行
    df = df.drop(index = 0)
    # 过滤重复数据（因为三次测定的均值相同，所以取任一行）
    lbres_series = df.groupby(['受试者编号', '访视编号'])[col_name].min()
    df = lbres_series.to_frame() # 从Series创建Frame
    # print(df)
    df = df.reset_index(drop=False)
    df = df.pivot(index='受试者编号', columns='访视编号', values=col_name)
    # 生成新列名dict
    col_map = {col: rename_col(col, col_name) for col in df.columns}
    # 对lbivm_df应用"列名映射"
    df.rename(columns=col_map, inplace=True)
    
    return df


    