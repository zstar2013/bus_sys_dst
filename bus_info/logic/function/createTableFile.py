from xlwt import Workbook, easyxf
from ui.style import xlsStyle
from bus_info.logic.oilsta_data.oilstationCompute import OliFileDispatcher
from tools.filetool import searchforFileWithCallback
import traceback
import ui.const.const_value as const
import datetime
default_width = 256 * 30
width_short = 256 * 5
width_longer = 256 * 10
avg_target_set=dict()

# 创建油料中心文件
def create_oilstation_feedback_file(newWb, str_date, path):
    ws = newWb.add_sheet("油料中心数据表")
    styleTitle = easyxf(xlsStyle.styles["table"]["table_title"])
    ws.write_merge(0, 0, 0, 2, u'各线路油耗油料中心数据', styleTitle)

    styleSubTitle = easyxf(xlsStyle.styles["table"]["table_subtitle"])
    ws.write_merge(1, 1, 0, 2, str_date, styleSubTitle)

    stylenormal = easyxf(xlsStyle.styles["table"]["table_normal"])
    ws.write(2, 0, u'车队', stylenormal)
    ws.write(2, 1, u'线路', stylenormal)
    ws.write(2, 2, u'油耗', stylenormal)

    ws.col(0).width = default_width
    ws.col(1).width = default_width
    ws.col(2).width = default_width

    path_ = path
    dper_ = OliFileDispatcher()
    searchforFileWithCallback(path=path_, dper=dper_)
    team_results = OilChargeDetail.objects.values('team').distinct()
    currentIndex = 2
    total_charge = 0
    second_charge = 0
    back_up_charge = 0
    for team_result in team_results:
        team_charge = 0
        start_index = currentIndex + 1
        route_results = OilChargeDetail.objects.filter(team=team_result['team']).values('route').distinct()

        for route_result in route_results:
            currentIndex += 1
            charge_result = OilChargeDetail.objects.filter(route=route_result['route']).aggregate(Sum('charge'))
            ws.write(currentIndex, 1, route_result['route'], stylenormal)
            print(route_result['route'])
            ws.write(currentIndex, 2, charge_result['charge__sum'], stylenormal)
            print(charge_result['charge__sum'])
            team_charge += charge_result['charge__sum']
        currentIndex += 1
        ws.write_merge(start_index, currentIndex, 0, 0, team_result['team'], stylenormal)
        ws.write(currentIndex, 1, "汇总", stylenormal)
        ws.write(currentIndex, 2, team_charge, stylenormal)
        total_charge += team_charge
        if team_result['team'] == "营达":
            second_charge = team_charge
        if team_result['team'] == "一公司公备汇总":
            back_up_charge = team_charge

    ws.write_merge(currentIndex + 1, currentIndex + 4, 0, 0, "汇总", stylenormal)
    ws.write(currentIndex + 1, 1, "一公司营运汇总", stylenormal)
    ws.write(currentIndex + 1, 2, total_charge - back_up_charge - second_charge, stylenormal)

    currentIndex += 1

    ws.write(currentIndex + 1, 1, "一公司汇总", stylenormal)
    ws.write(currentIndex + 1, 2, total_charge - second_charge, stylenormal)

    currentIndex += 1
    ws.write(currentIndex + 1, 1, "营达汇总", stylenormal)
    ws.write(currentIndex + 1, 2, second_charge, stylenormal)

    currentIndex += 1
    ws.write(currentIndex + 1, 1, "全汇总", stylenormal)
    ws.write(currentIndex + 1, 2, total_charge, stylenormal)

    ws_detail = newWb.add_sheet("油料中心明细")
    styleTitle = easyxf(xlsStyle.styles["table"]["table_title"])
    ws_detail.write_merge(0, 0, 0, 6, u'油料中心明细', styleTitle)
    fields = ['车队', '线路', '车号', '交易时间', '印刷号', '加油站位置', '加油量']
    stylenormal = easyxf(xlsStyle.styles["table"]["table_normal"])
    for i in range(0, 7):
        ws_detail.col(i).width = default_width
        ws_detail.write(1, i, fields[i], stylenormal)
    currentIndex = 2
    list_ = OilChargeDetail.objects.filter(paytime__startswith='2018-07')
    for detail in list_:
        currentIndex += 1
        ws_detail.write(currentIndex, 0, detail.team)
        ws_detail.write(currentIndex, 1, detail.route)
        ws_detail.write(currentIndex, 2, detail.car_id)
        ws_detail.write(currentIndex, 3, detail.paytime)
        ws_detail.write(currentIndex, 4, detail.printnum)
        ws_detail.write(currentIndex, 5, detail.station)
        ws_detail.write(currentIndex, 6, detail.charge)
