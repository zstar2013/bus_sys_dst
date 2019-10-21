from xlwt import Workbook, easyxf
from ui.style import xlsStyle
import ui.const.const_value as const
from collections import Counter

default_width = 256 * 30
width_short = 256 * 5
width_longer = 256 * 10


###=====================================统计表部分==========================================
#初始化 统计表表头
def init_feebback_sum_table(ws, date, team="",offsetY=0):
    cells = list.copy(const.mergeCells_base_sum)
    str_date = date.strftime('%Y{y}%m{m}').format(y='年', m='月')
    cells[1]["value"] = str_date + "         第1页"
    if team is "":
        cells[3] = {"location": [2, 4, 1, 2], "value": "线路", "style": "table_normal"}
        del (cells[4])
    for i in [0, 1, 5, 10, 11, 12]:
        ws.col(i).width = width_short
    ws.col(3).width = width_longer
    for cell in cells:
        write_cell(ws, cell,offsetY)
		
# 生成单车燃料消耗统计表
def write_car_oil_cost_sum(ws, date, data, team="", route=""):
    init_feebback_sum_table(ws=ws, date=date, team=team)
    current_index = 5
    style_ = easyxf(xlsStyle.styles["table"]["table_normal"])
    for i in range(len(data)):
        item=data.iloc[i]
        ws.write(current_index, 2, item['sub_car_id'], style_)
        target_value=float(item["target_oil_cost"])
        ws.write(current_index, 5, target_value, style_)
        write_feedback_sum_row(ws, current_index, item, date, style_)
        current_index += 1
    sum_=data.sum()
    ws.write(current_index, 2, "总计:", style_)
    most_target=Counter(data["target_oil_cost"]).most_common()[0][0]
    ws.write(current_index, 5, int(most_target), style_)
    write_feedback_sum_row(ws, current_index, sum_, date, style_,case="sum")
	
    s_team = "一\n公\n司\n" + str(team) + "\n车\n队"
    ws.write_merge(5, current_index, 0, 0, s_team, style_)
    ws.write_merge(5, current_index, 1, 1, route, style_)
    ws.write_merge(current_index + 2, current_index + 2, 9, 11, "制表人：", )
    
#描统计表单行 
def write_feedback_sum_row(ws, current_index, result, date, style_,case="oil"):
    total_target = result["total_oil_target"]
    if case=="oil":
        if result['oil_cost'] is None:
            return
    elif case=="elc":
        if result["electric_cost"] is None:
            return

    total_oilwear = result['oil_cost']
    ws.write(current_index, 3, round(result['mileage'], 2), style_)
    ws.write(current_index, 4, round(total_target, 2), style_)
    ws.write(current_index, 6, round(result['oil_cost'], 2), style_)
    ws.write(current_index, 7,
             round(result['oil_cost'] * 100 / result['mileage'] if result['mileage'] != 0 else 0,
                   2), style_)
    if total_oilwear - total_target >= 0:
        ws.write(current_index, 8, "", style_)
        ws.write(current_index, 9, round(total_oilwear - total_target, 2), style_)
    else:
        ws.write(current_index, 8, round(total_target - total_oilwear, 2), style_)
        ws.write(current_index, 9, "", style_)
    ws.write(current_index, 10, round(result['maintain'], 2), style_)
    ws.write(current_index, 11, round(result['follow'], 2), style_)
    ws.write(current_index, 12, round(result['inspection'], 2), style_)
	
###=====================================统计表部分结束==========================================

###=====================================汇总表部分==============================================

#初始化 汇总表表头
def init_feedback_detail_table(ws, date, team="", route=""):
    cells = list.copy(const.mergeCells_base_detail)
    str_date = date.strftime('%Y{y}%m{m}').format(y='年', m='月')
    cells[2]["value"] = str_date
    if team is not "" and route is not "":
        cells[1]["value"] = "福州市公交第一公司" + str(team) + "车队   " + str(route) + "路"
    for i in [6, 7]:
        ws.col(i).width = width_longer
    for cell in cells:
        write_cell(ws, cell)

# 生成单车运行情况汇总表
def write_car_oil_cost_detail(ws, date,data, team="", route=""):
    style_ = easyxf(xlsStyle.styles["table"]["table_normal"])
    init_feedback_detail_table(ws=ws, date=date, team=team, route=route)
    fileds_ = list.copy(const.detail_query_fields)
    current_index = 5
    for i in range(len(data)):
        item=data.iloc[i]
        ws.write(current_index, 2, item['sub_car_id'], style_)
        target_value=float(item["target_oil_cost"])
        ws.write(current_index, 5, target_value, style_)
        write_detail_row(ws=ws, current_index=current_index, result=item, style_=style_, fileds=fileds_)
        current_index += 1
    for i in range(1, 19):
        ws.write(current_index, i, "", style_)

    result = MonthlyFeedback.objects.filter(date=date, route=route).aggregate(*sum_, target_total=Sum(F('mileage') * F(
        "target_in_compute") / 100,
                                                                                                      output_field=fields.FloatField()))
    current_index += 1
    ws.write(current_index, 0, "总计", style_)
    write_detail_row(ws=ws, current_index=current_index, result=result, style_=style_, fileds=fileds_sum)

    current_index += 2
    rmdlist = RouteMonthlyDetail.objects.filter(date=date, route=route).values("num_fir_maintain", "num_sec_maintain")
    if len(rmdlist) is 0:
        return
    rmd = rmdlist[0]
    style_ = easyxf(xlsStyle.styles["table"]["table_tail"])
    ws.write_merge(current_index, current_index, 0, 5, "一保：" + str(rmd["num_fir_maintain"]) + "部", style_)
    ws.write_merge(current_index, current_index, 6, 10, "二保：" + str(rmd["num_sec_maintain"]) + "部", style_)
def write_detail_row(ws, current_index, result, style_, fileds):
    if result[fileds[1]] is None:
        return
    total_days = result[fileds[1]] + result[fileds[2]] + result[fileds[3]]
    ws.write(current_index, 1, round(total_days, 2), style_)
    ws.write(current_index, 2, round(result[fileds[2]], 2), style_)
    total_well_days = result[fileds[1]] + result[fileds[3]]
    ws.write(current_index, 3,
             round(total_well_days, 2), style_)
    ws.write(current_index, 4,
             round(result[fileds[3]], 2), style_)
    ws.write(current_index, 5,
             round(result[fileds[1]], 2), style_)
    ws.write(current_index, 6, round(result[fileds[0]], 2), style_)

    ws.write(current_index, 7, round(
        result[fileds[0]] - result[fileds[6]] - result[fileds[5]] - result[
            fileds[4]], 2), style_)
    ws.write(current_index, 8, round(result[fileds[6]], 2), style_)
    ws.write(current_index, 9, round(result[fileds[5]], 2), style_)
    ws.write(current_index, 10, round(result[fileds[4]], 2), style_)
    ws.write(current_index, 11, round(total_well_days / total_days * 100 if total_days != 0 else 0, 2), style_)
    ws.write(current_index, 12, round(result[fileds[1]] / total_days * 100 if total_days != 0 else 0, 2), style_)
    ws.write(current_index, 13, round(
        result[fileds[8]] * 60 / result[fileds[0]] * 100 if result[fileds[0]] != 0 else 0, 2),
             style_)
    ws.write(current_index, 14, round(result[fileds[7]], 2), style_)
    ws.write(current_index, 15, round(result[fileds[8]], 2), style_)
    ws.write(current_index, 16, round(result["target_total"], 2), style_)
    ws.write(current_index, 17, round(result[fileds[9]], 2), style_)
    ws.write(current_index, 18, "", style_)

###=====================================汇总表部分结束==========================================

#填充cell	
def write_cell(ws, mc, offset=0):
    style_ = easyxf(xlsStyle.styles["table"][mc.get("style")])
    local_ = mc.get("location")
    if (len(local_) == 2):
        r, c = local_
        ws.write(offset + r, c, mc.get("value"), style_)
    if (len(local_) == 4):
        r1, r2, c1, c2 = local_
        ws.write_merge(offset + r1, offset + r2, c1, c2, mc.get("value"), style_)