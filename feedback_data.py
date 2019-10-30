from xlwt import Workbook, easyxf
from ui.style import xlsStyle
import ui.const.const_value as const
from collections import Counter
from bus_info.models import list_team_routes    
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
    if case=="oil":    
        if result['oil_cost'] is None:    
            return
        total_target = result["total_oil_target"]
        total_cost = result['oil_cost'] 
    elif case=="elec":    
        if result["electric_cost"] is None:    
            return
        total_target = result["total_elc_target"]
        total_cost = result['elec_cost']
    
       
    ws.write(current_index, 3, round(result['mileage'], 2), style_)    
    ws.write(current_index, 4, round(total_target, 2), style_)    
    ws.write(current_index, 6, round(total_cost, 2), style_)    
    ws.write(current_index, 7,    
             round(total_cost * 100 / result['mileage'] if result['mileage'] != 0 else 0,    
                   2), style_)    
    if total_cost - total_target >= 0:    
        ws.write(current_index, 8, "", style_)    
        ws.write(current_index, 9, round(total_cost - total_target, 2), style_)    
    else:    
        ws.write(current_index, 8, round(total_target - total_cost, 2), style_)    
        ws.write(current_index, 9, "", style_)    
    ws.write(current_index, 10, round(float(result['maintain']), 2), style_)    
    ws.write(current_index, 11, round(result['follow'], 2), style_)    
    ws.write(current_index, 12, round(result['inspection'], 2), style_)


# 生成单车电能消耗统计表
def write_car_electric_cost(ws, route, date, data, team=""):
    const.mergeCells_base_sum[0]["value"] = "车公里单车电能消耗统计表"
    const.mergeCells_base_sum[6]["value"] = "电能消耗(度)"
    init_feebback_sum_table(ws=ws, date=date, team=team)
    style_ = easyxf(xlsStyle.styles["table"]["table_normal"])
    current_index = 5
    data_elec=data[data[power_type]=="电"]
    #TODO 还需完成混动车充电情况
    for item in data_elec:
        target_value = float(item["target_elc_cost"])
        total_target=float(item["total_elc_target"])
        ws.write(current_index, 5, target_value, style_)
        write_feedback_sum_row(ws, current_index, item, date, style_,case="elec")
        current_index += 1

    result = MonthlyFeedback.objects.filter(date=date, route=route,carInfo__cartype__power_type="电").aggregate(
        mileage__sum=Sum('mileage'), electric_cost=Sum('electric_cost'),
        maintain__sum=Sum('maintain'), follow__sum=Sum('follow'),
        inspection=Sum('inspection'),
        target_total=Sum(F('mileage') * F('target_in_compute_2') / 100,
                         output_field=fields.FloatField()))
    #result["target_total"] = target_value * (result['mileage__sum'] if result['mileage__sum'] is not None else 0) / 100
    write_feedback_elect_row_sum(ws, current_index, result, date, style_)
    ws.write_merge(5, current_index, 1, 1, route, style_)
    ws.write_merge(5, current_index, 0, 0, "公\n交\n一\n公\n司", style_)
    ws.write(current_index, 2, "总计：", style_)    
        
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
    current_index = 5    
    for i in range(len(data)):    
        item=data.iloc[i]
        ws.write(current_index, 0, item['sub_car_id'], style_)    
        target_value=float(item["target_oil_cost"])      
        write_detail_row(ws=ws, current_index=current_index, result=item, style_=style_)    
        current_index += 1    
    for i in range(1, 19):    
        ws.write(current_index, i, "", style_)    
    sum_=data.sum()   
    current_index += 1    
    ws.write(current_index, 0, "总计：", style_)    
    write_detail_row(ws=ws, current_index=current_index, result=sum_, style_=style_)        
    current_index += 2
    #TODO 一保二保数目待实现
    # rmdlist = RouteMonthlyDetail.objects.filter(date=date, route=route).values("num_fir_maintain", "num_sec_maintain")    
    # if len(rmdlist) is 0:    
        # return    
    # rmd = rmdlist[0]    
    # style_ = easyxf(xlsStyle.styles["table"]["table_tail"])    
    # ws.write_merge(current_index, current_index, 0, 5, "一保：" + str(rmd["num_fir_maintain"]) + "部", style_)    
    # ws.write_merge(current_index, current_index, 6, 10, "二保：" + str(rmd["num_sec_maintain"]) + "部", style_)    
def write_detail_row(ws, current_index, result, style_):    
    if result['work_days'] is None:    
        return    
    total_days = float(result['work_days'] + result['fix_days'] + result['stop_days'])    
    fix_days=float(result['fix_days'])    
    total_well_days = float(result['work_days'] + result['stop_days'])    
    stop_days=float(result['stop_days'])    
    work_days=float(result['work_days'])    
    ws.write(current_index, 1, round(total_days, 2), style_)    
    ws.write(current_index, 2, round(fix_days, 2), style_)    
    ws.write(current_index, 3,    
             round(total_well_days, 2), style_)    
    ws.write(current_index, 4,    
             round(stop_days, 2), style_)    
    ws.write(current_index, 5,    
             round(work_days, 2), style_)    
    ws.write(current_index, 6, round(result['mileage'], 2), style_)    
    
    ws.write(current_index, 7, round(    
        result['mileage'] - result['engage_mileage'] - result['public_mileage'] - result[    
            'shunt_mileage'], 2), style_)    
    ws.write(current_index, 8, round(float(result['engage_mileage']), 2), style_)    
    ws.write(current_index, 9, round(float(result['public_mileage']), 2), style_)    
    ws.write(current_index, 10, round(float(result['shunt_mileage']), 2), style_)    
    ws.write(current_index, 11, round(total_well_days / total_days * 100 if total_days != 0 else 0, 2), style_)    
    ws.write(current_index, 12, round(result['work_days'] / total_days * 100 if total_days != 0 else 0, 2), style_)    
    ws.write(current_index, 13, round(    
        result['fault_minutes'] * 60 / result['mileage'] * 100 if result['mileage'] != 0 else 0, 2),    
             style_)    
    ws.write(current_index, 14, round(float(result["fault_times"]), 2), style_)    
    ws.write(current_index, 15, round(float(result["fault_minutes"]), 2), style_)    
    ws.write(current_index, 16, round(float(result["total_oil_target"]), 2), style_)    
    ws.write(current_index, 17, round(result["oil_cost"], 2), style_)    
    ws.write(current_index, 18, "", style_)    
    
###=====================================汇总表部分结束==========================================    

###======================================总表部分开始===========================================


# 生成反馈统计月报总表
def write_monthy_feedback_sum_table(ws, date, data):
    style_ = easyxf(xlsStyle.styles["table"]["table_normal"])
    init_feebback_sum_table(ws, date)
    current_index = 5
    for item in list_team_routes:
        for route in item.get("routes"):
            data_route=data[data['route']==route]
            data_sum=data_route.sum()
            ws.write_merge(current_index, current_index, 1, 2, route, style_)
            list_=Counter(data_route["target_oil_cost"]).most_common()
            if len(list_)!=0:
                most_target=list_[0][0]    
            ws.write(current_index, 5, most_target,style_)
            write_feedback_sum_row(ws, current_index, data_sum, date, style_)
            current_index += 1

    ws.write_merge(current_index, current_index, 1, 2, "汇总", style_)
    sum_=data.sum()
    write_feedback_sum_row(ws, current_index, sum_, date, style_)
    current_index += 1
    ws.write_merge(current_index, current_index, 1, 2, "国有统计", style_)
    write_feedback_sum_row(ws, current_index, sum_, date, style_)
    current_index += 1
    ws.write_merge(current_index, current_index + 2, 10, 12,
                   "二保：" + str(round(sum_['maintain'] if sum_['maintain'] is not None else 0, 2)) + "\n跟车：" + str(
                       round(sum_['follow'] if sum_['follow'] is not None else 0, 2)),
                   style_)

    ws.write_merge(current_index, current_index + 2, 1, 9, "", style_)

    current_index += 3
    for cell in const.mergeCells_base_sum_tail:
        write_cell(ws, cell, current_index)

    current_index += 1

    for num in const.backup_cars_num:
        ws.write_merge(current_index, current_index, 1, 2, num, style_)
        ws.write_merge(current_index, current_index, 3, 4, "", style_)
        ws.write_merge(current_index, current_index, 5, 6, "", style_)
        ws.write_merge(current_index, current_index, 7, 8, "", style_)
        for i in range(9, 13):
            ws.write(current_index, i, "", style_)
        current_index += 1

    ws.write_merge(current_index, current_index, 1, 2, "小计：", style_)
    ws.write_merge(current_index, current_index, 3, 4, "", style_)
    ws.write_merge(current_index, current_index, 5, 6, "小计：", style_)
    ws.write_merge(current_index, current_index, 7, 8, "", style_)
    for i in range(9, 13):
        ws.write(current_index, i, "", style_)

    ws.write_merge(5, current_index, 0, 0, "公\n交\n一\n公\n司", style_)

    ws.write_merge(current_index + 2, current_index + 2, 9, 11, "制表人：", )
    
    
#生成反馈汇总月报总表
def write_monthy_feedback_detail_table(ws, date,data):
    """
    生成月反馈汇总表
    :param ws:
    :param date:
    :return:
    """
    style_ = easyxf(xlsStyle.styles["table"]["table_normal"])
    init_feedback_detail_table(ws, date)
    current_index = 5
    for item in list_team_routes:
        for route in item.get("routes"):
            data_route=data[data['route']==route]
            data_sum=data_route.sum()
            ws.write(current_index, 0, route, style_)
            write_detail_row(ws=ws, current_index=current_index, result=data_sum, style_=style_)
            current_index += 1
    for i in range(1, 19):
        ws.write(current_index, i, "", style_)
    current_index += 1
    ws.write(current_index, 0, "总计", style_)
    data_sum=data.sum()
    write_detail_row(ws=ws, current_index=current_index, result=data_sum, style_=style_)

#生成公里总表    
def write_monthy_mileage_table(ws,date,data):
    data_=data[["sub_car_id","route","mileage"]]
    for current_index,item in enumerate(data_.values):
        for index,value in enumerate(item):
            ws.write(current_index,index,value)
    
###==================================总表部分结束=====================================================
    
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