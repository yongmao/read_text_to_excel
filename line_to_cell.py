
from datetime import datetime
def std_date(datestr):
    datestr = datestr.replace("（", "(").split("(")[0]
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

# 以1为起点
def parse(line:str):
    cell_init = ()
    cell = ()

    line_std = line.replace("：",":").strip()
    line_tuple = line_std.partition(":")
    if line_tuple[1] == ":":
        key = line_tuple[0].strip()
        value = line_tuple[2].strip()
        match key:
            case "ID":
                cell = cell_init + (1,)

            case "客户" | "客人" | "乘客姓名" | "联系人"|"客人姓名"|"接机牌":
                cell = cell_init + (2,)
                cell = cell + (value,)

            case "用车时间" | "日期" | "出车日期":
                cell = cell_init + (3,)
                cell = cell + (std_date(value),)
                # cell = cell + (value,)

            case "司机" | "姓名" | "师傅" | "服务司机"|"司机姓名"|"出车司机":
                cell = cell_init + (4,)
                cell = cell + (value,)

            case "电话" | "手机"|"司机电话":
                cell = cell_init + (5,)
                cell = cell + (value,)

            case "车型" | "车子"|"车辆型号":
                cell = cell_init + (6,)
                cell = cell + (value,)

            case "车号" | "车牌"|"车牌号码":
                cell = cell_init + (7,)
                cell = cell + (value,)

            case "用车地点" | "城市" | "用车地":
                cell = cell_init + (8,)
                cell = cell + (value,)

            case "行程" | "出车行程" | "路线"|"用车行程":
                cell = cell_init + (9,)
                cell = cell + (value,)

            case "Time" | "时间"|"起止时间":
                cell = cell_init + (10,)
                cell = cell + (value,)

            case "用车时长":
                cell = cell_init + (11,)
                cell = cell + (value,)

            case "超出小时":
                cell = cell_init + (12,)
                cell = cell + (value,)

            case "行车里程" | "行驶" | "行驶里程" | "公里" | "里程"|"行驶公里":
                cell = cell_init + (13,)
                # 公里
                cell = cell + (value.split("公里")[0],)

            case "超公里数":
                cell = cell_init + (14,)
                cell = cell + (value,)

            case "OK":
                cell = cell_init + (15,)
                cell = cell + (value.split("元")[0],)

            case "OT超时":
                cell = cell_init + (16,)
                cell = cell + (value.split("元")[0],)

            case "停车费" | "停车"|"停车费用":
                cell = cell_init + (17,)
                cell = cell + (value.split("元")[0],)

            case "高速费" | "高速" | "高速费用":
                cell = cell_init + (18,)
                cell = cell + (value.split("元")[0],)

            case "餐费" | "歺补":
                cell = cell_init + (19,)
                cell = cell + (value.split("元")[0],)

            # 人数 人数：2位
            # 航班信息：CZ346, 13:40
            # 举牌子：Avakova Karine （给卡丽娜）
            # 接机牌：Mr.Mads.GyIdenberg
            # 起点：首都机场T3
            # 终点：世贸工三国际公寓
            # 联系人：待定
            # 餐厅地址：黄浦区圆明园路88号益外滩源2楼西侧201
            # 过路费：11元
            case _:
                print("===" + key)
    return cell



if __name__ == '__main__':
    # print_hi('PyCharm')
    mycell = parse("司机：肖永茂")
    print(mycell)

    print(parse("出车报告"))

