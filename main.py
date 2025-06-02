import pandas as pd
import openpyxl
from lf_stat import calculate_lf_table
from lf_ivm_stat import calculate_lf_and_ivm
from pef_stat import calculate_pef

pd.set_option("display.max_rows", None)

with pd.ExcelFile("./data.xlsx", engine="openpyxl") as xls:
    calculate_lf_table(xls) # 统计lf（肺检查）表
    calculate_lf_and_ivm(xls) # 统计lf（肺检查）和ivm（流速）表
    calculate_pef(xls)





