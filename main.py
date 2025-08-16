# 读text文件，再导入到excel
from datetime import datetime

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import math
import re

from pandas import isnull

celldict = {
    '出车报告': '',
    '日期': '日期',
    '出车日期': '日期',
    '司机': '司机',
    '姓名': '司机',
    '电话': '电话',
    '车型': '车型',
    '车牌': '车牌',
    '车号': '车牌',
    '时间': '时间',
    '用车时长': '用车时长',
    '超出小时': '超出小时',
    '出车行程': '出车行程',
    '行程': '出车行程',
    '行驶里程': '行驶里程',
    '行驶': '行驶里程',
    '超公里数': '超公里数',
    '停车费用': '停车费用',
    '停车': '停车费用',
    '高速费用': '高速费用',
    '高速': '高速费用',
    '客户': '客户'
}


def kilo_string_to_float(str):
    if isnull(str):
        return 0
    else:
        array = split_string_with_multiple_delimiters(str, "公里")
        if len(array) > 0:
            return array[0].strip()
        else:
            return str.strip()

def time_string_to_minutes(time_str):
    """
    将形如 "x小时y分钟" 的字符串转换为总分钟数。

    Args:
        time_str:  小时分钟字符串，例如 "2小时30分钟"。

    Returns:
        总分钟数，例如对于 "2小时30分钟" 返回 150。
    """
    if time_str.find("小时") != -1 and time_str.find("分") != -1:
        match = re.match(r'(\d+)小时(\d+)分', time_str)
        if match:
            hours = int(match.group(1))
            minutes = int(match.group(2))
            total_minutes = hours * 60 + minutes
            return total_minutes
        else:
            return None  # 或者抛出异常，表示格式不正确
    elif time_str.find("小时") != -1 and time_str.find("分") == -1:
        match = re.match(r'(\d+)小时', time_str)
        if match:
            hours = int(match.group(1))
            total_minutes = hours * 60
            return total_minutes
        else:
            return None  # 或者抛出异常，表示格式不正确
    elif time_str.find("分") != -1:
        match = re.match(r'(\d+)分', time_str)
        if match:
            minutes = int(match.group(1))
            total_minutes =  minutes
            return total_minutes
        else:
            return None  # 或者抛出异常，表示格式不正确
# 1.5元 => 1.5
def fee_str_float(str):
    if isnull(str):
        return 0
    else:
        array = split_string_with_multiple_delimiters(str, "元")
        if len(array) > 0 :
            return array[0].strip()
        else:
            return str.strip()

def split_string_with_multiple_delimiters(text, delimiters):
  """
  使用多个分隔符分割字符串。

  参数:
    text: 要分割的字符串。
    delimiters: 分隔符列表或元组。

  返回:
    包含分割后子字符串的列表。
  """
  regex = '|'.join(map(re.escape, delimiters))
  return re.split(regex, text)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

def std_datetime(datestr):
    if datestr.find("年") != -1:
        db_datetime1 = datetime.strptime(datestr, '%Y年%m月%d日')
        return db_datetime1.strftime('%Y/%m/%d')
    elif datestr.find(".") != -1:
        db_datetime2 = datetime.strptime(datestr, '%Y.%m.%d')
        return db_datetime2.strftime('%Y/%m/%d')
    elif datestr.find("-") != -1:
        db_datetime3 = datetime.strptime(datestr, '%Y-%m-%d')
        return db_datetime3.strftime('%Y/%m/%d')
    elif datestr.find("/") != -1:
        return datestr
    else:
        return datestr

def import_data():
    # cell_set = set()
    # 读取文件
    delimiters = [':', '：']
    df = pd.read_excel('./wxbg.xlsx',sheet_name='微信出车报告')

    # 写文件
    existing_file = './ccjl.xlsx'
    df_existing = pd.read_excel(existing_file, sheet_name='出车报告')

    for index,row in df.iterrows():
        # 来自微信的出行报告, row是一个出行报告
        if  math.isnan(row['导入']):


            new_data = {}   # 待添加的数据
            wx_cxbg_str = row['来自微信的出车报告']
            # print(wx_cxbg_str)

            # 格式分析
            cell_lines = wx_cxbg_str.split('\n')
            # print(cell_lines)


            # 循环读取每一份报告
            for cell_line in cell_lines:

                cells = split_string_with_multiple_delimiters(cell_line, delimiters)
                # print(result)
                cell_name = cells[0].strip()
                if cell_name.isspace():
                    continue

                cell_acron = celldict[cell_name]

                if cell_acron.isspace():
                    # print(cell_acron)
                    continue

                # get the value
                value = ''
                if len(cells) == 2:
                    value = cells[1]
                else:
                    if len(cells) > 2:
                        value = cells[1]
                        for index,item in enumerate(cells):
                            # print(index,item)
                            if index < 2:
                                    continue
                            value += ":" + item

                match cell_acron:
                    case "日期":
                        # print("日期.")
                        # 日期 2025年8月12日  2025.8.11
                        # print("日期.",std_datetime(value))
                        new_data["日期"] = [std_datetime(value)]

                    case "司机":
                        # print("司机")
                        # 司机 赵江
                        new_data["司机"] = [value]

                    case "电话":
                        # print("电话")
                        # 电话 13911134510
                        new_data["电话"] = [value]

                    case "车型":
                        # print("车型")
                        # 车型 别克653T
                        new_data["车型"] = [value]

                    case "车牌":
                        # print("车牌")
                        # 车牌 京N8YY99
                        new_data["车牌"] = [value]

                    case "时间":
                        # print("时间")
                        # 时间 11:30-15:35
                        new_data["时间"] = [value]

                    case "用车时长":
                        # print("用车时长")
                        # 用车时长 1小时
                        # 用车时长（分）
                        new_data["用车时长"] = [value]
                        new_data["用车时长（分）"] = [time_string_to_minutes(value)]

                    case "超出小时":
                        # print("超出小时")
                        new_data["超出小时"] = [value]

                    case "出车行程":
                        # print("出车行程")
                        # 出车行程 四元桥宜家~首都机场T3~国贸商城~柏悦酒店
                        new_data["出车行程"] = [value]

                    case "行驶里程":
                        # print("行驶里程")
                        # 行驶里程 55公里
                        # new_data["行驶里程"] = [value]
                        kilo = kilo_string_to_float(value)
                        new_data["行驶里程"] = [kilo_string_to_float(value)]

                    case "超公里数":
                        # print("超公里数")
                        new_data["超公里数"] = [value]

                    case "停车费用":
                        # print("停车费用")
                        # 停车费用 28
                        # new_data["停车费用"] = [value]
                        new_data["停车费用"] = [fee_str_float(value)]

                    case "高速费用":
                        # print("高速费用")
                        # 高速费用 20
                        # new_data["高速费用"] = [value]
                        new_data["高速费用"] = [fee_str_float(value)]

                    case "客户":
                        # print("客户")
                        # 客户 苏小姐
                        new_data["客户"] = [value]
                    case _:
                        print("其他." + cell_acron)

                print(cell_acron,value)
            # if

            print(new_data)                   # 新的字典

            # New data to append
            df_new = pd.DataFrame(new_data)   # 创建新数据行

            # Append new data
            df_existing = pd.concat([df_existing, df_new], ignore_index=False)

        # if  math.isnan(row['导入']):
    # for index,row

    # write to excel
    df_existing.to_excel(existing_file,sheet_name='出车报告',index=False)

    # print(sorted(cell_set))





# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # print_hi('PyCharm')
    import_data()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
