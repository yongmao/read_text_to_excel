# 分析报告格式


from openpyxl import load_workbook
import pandas as pd
import re

from pandas import isnull

excel_file = "用车明细.xlsx"
sheet_name='Sheet1'

delimiters = [':', '：']
regex = '|'.join(map(re.escape, delimiters))

key_set = set()
# 分析出行报告，找出标题
def parse_v_data():
    print("分析微信出车报告格式")
    print("====================================")

    df = pd.read_excel(excel_file,sheet_name=sheet_name)

    for index,row in df.iterrows():
        wx_cxbg_str = row['备注1']
        print(wx_cxbg_str)

        if isnull(wx_cxbg_str):
            continue
        print("====================================")
        continue

        cell_lines = wx_cxbg_str.split('\n')

        for cell_line in cell_lines:
            print(cell_line)
            # cell_line_std = cell_line.strip().replace(" ", "")
            cell_array = re.split(regex, cell_line)

            key = ""
            value = ""

            print(cell_array)

            if len(cell_array) == 1:
                # 不包含：
                continue
            elif len(cell_array) == 2:
                key = cell_array[0]
                value = cell_array[1]
            elif len(cell_array) > 2:
                key = cell_array[0]
                value = ':'.join(cell_array[1:])
            else:
                continue
            print(key,value)
            key_final = key.strip().replace(" ", "")
            value_final = value.strip().replace(" ", "")
            print(key_final,value_final)
            key_set.add(key_final)
    # for index,row in df.iterrows():
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # print_hi('PyCharm')
    parse_v_data()
    print(sorted(key_set))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
