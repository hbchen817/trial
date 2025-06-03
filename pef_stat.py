from utils import generate_pef_frame, save_excel
import pandas as pd



def calculate_delta(row: pd.Series):
    result = {}
    for col_idx in range(1, len(row)):
        if not pd.isna(row.iloc[col_idx]):
            result[f'{row.index[col_idx]}_delta'] = row.iloc[col_idx] - row.iloc[0]
    return pd.Series(result)

def calculate_pef(xls):
    df = generate_pef_frame(xls, 'PEF测量(lb_pef)')
    
    delta_df = df.apply(calculate_delta, axis = 1)
    df = pd.concat([df, delta_df], axis=1)
    df = df.fillna('')
    df = df.reset_index()
    # print(df)
    # 保存结果
    save_excel(df, sheet_name = 'pef试验delta值')