import xlrd
from bus_info.models import BusInfo, MonthlyFeedback,RouteMonthlyDetail
from tools import strtool as st
import re
import traceback
import os
from tools.filetool import FileDispatcher
from tools.strtool import isContain,isContainOr

xOffset=3

class xlsDispatcher(FileDispatcher):
    current_date=None
    def __init__(self,date):
        self.current_date=date


    def file_dispatch(self,filepath):
        filter_key = ["油耗和汇总表","2019.","每月油耗 (version 1)","油表","油耗汇总表"]
        filter_table_keys = ["统计"]
        filter_table_keys_detail=["汇总"]
        filter_type = ".xls"
        items=[]
        if os.path.splitext(filepath)[1] == filter_type and isContainOr(keys=filter_key, target=filepath):
            data = xlrd.open_workbook(filepath)
            for sheet in data.sheets():
                if isContainOr(filter_table_keys,sheet.name):
                    print(sheet.name)
                    table = data.sheet_by_name(sheet.name)
                    #print("row and cols",table.nrows,table.ncols)
                    loacl_of_startItem = getStartIndex(table)
                    #print(loacl_of_startItem)
                    if loacl_of_startItem is None:
                        continue
                    else:
                        #items.extend(load_feedback_sum_data(table, loacl_of_startItem))
                        load_feedback_sum_data(table, loacl_of_startItem,self.current_date)
            for sheet in data.sheets():
                if isContainOr(filter_table_keys_detail,sheet.name):
                    print(sheet.name)
                    table = data.sheet_by_name(sheet.name)
                    loacl_of_startItem = getStartIndex(table)
                    if loacl_of_startItem is None:
                        continue
                    else:
                        #items.extend(load_feedb_detail_data(table, loacl_of_startItem))
                        load_feedb_detail_data(table, loacl_of_startItem,self.current_date)

        return items

avg_target_set=dict()
from django.db.models import Sum,Avg
from django.db.models import F, fields
def get_one_car_target(route,date,car_type):
    """
    返回单车指标
    :param feedbackItem:
    :return:
    """
    if car_type.target_value1!=-1:
        if date.month > 5 and date.month < 10:
            return car_type.target_value2
        else:
            return car_type.target_value1
    key=str(route)+str(date)+str(car_type)
    #先从池中查找
    if avg_target_set.__contains__(key):
        return avg_target_set[key]
    #查找
    else:
        result = MonthlyFeedback.objects.filter(route=route, date=date,
                                                carInfo__cartype=car_type).aggregate(
            avg_target=Avg(F('oilwear') / F('mileage') * 100 if F('mileage') != 0 else 0))

        avg_target = 0
        if result["avg_target"] is not None:
            avg_target=round(result["avg_target"],0)
            print("_______________________________avg_target:", result["avg_target"])
        avg_target_set[key]=avg_target
        return avg_target



#读取统计数据文件
def load_feedback_sum_data(table, item,date):
    nrows = table.nrows
    indexR = item[0]
    carNoC = item[1]
    route = getRoute(table.name)
    for i in range(indexR, nrows):
        if table.cell(i, carNoC).value is "" or 0:
            continue
        if st.contains(str(table.cell(i, carNoC).value), "计"):
            continue
        car_id=getCar_id(table.cell(i, carNoC).value,table.name)
        bi_=bil.search_for_sub_car_id(car_id, route)
        if bi_ is None:
            print(" no result !!!! car_id:",car_id,"route:",route,"date:",date)
        mileage = getFloatValue(table.cell(i, carNoC + 1).value)  # 车公里
        team_target = getFloatValue(table.cell(i, carNoC + 3).value)  # 车队上报指标
        elec,oilwear=0,0
        if bi_.cartype.power_type=="电":
            elec=getFloatValue(table.cell(i, carNoC + 4).value)
        else:
            oilwear = getFloatValue(table.cell(i, carNoC + 4).value)  # 车辆油耗
        maintain = getFloatValue(table.cell(i, carNoC + 8).value)  # 二保
        follow = getFloatValue(table.cell(i, carNoC + 9).value)  # 跟车
        inspection=0
        if table.ncols>carNoC + 10 :
            try:
                inspection = getFloatValue(table.cell(i, carNoC + 10).value)
            except:
                traceback.print_exc()
                inspection = 0
        mf = MonthlyFeedback(fb_car_id=car_id,date=date, route=route, carInfo=bi_ if bi_ != None else None, mileage=mileage,
                             oilwear=oilwear, maintain=maintain, follow=follow,team_target=team_target,inspection=inspection,electric_cost=elec)
        mf.save()



#获取明细数据
def load_feedb_detail_data(table, item,date):

    nrows = table.nrows
    indexR = item[0]
    carNoC = item[1]
    route = getRoute(table.name)
    saveMainteData(table,route,date)
    for i in range(indexR, nrows):
        if table.cell(i, carNoC).value is "":
            continue
        if table.cell(i, carNoC).value is "0":
            continue
        if st.contains(str(table.cell(i, carNoC).value), "计"):
            continue
        if st.contains(str(table.cell(i, carNoC).value), "一保"):
            continue
        if st.contains(str(table.cell(i, carNoC).value), "备注"):
            continue
        car_id = getCar_id(table.cell(i, carNoC).value, table.name)
        try:
            mf=MonthlyFeedback.objects.get(fb_car_id=car_id,route=route,date=date)
        except:
            traceback.print_exc()
            print("--------------------------car_id:", car_id,"route:",route,date)
            continue

        #
        if mf is not None:
            try:
                mf.work_days= getFloatValue(table.cell(i, carNoC + 5).value)       # 工作天数
                mf.fix_days=getFloatValue(table.cell(i, carNoC + 2).value)         # 修理天数
                mf.stop_days=getFloatValue(table.cell(i, carNoC + 4).value)        # 停驶天数
                mf.shunt_mileage=getFloatValue(table.cell(i, carNoC + 10).value)   # 调车公里
                mf.engage_mileage=getFloatValue(table.cell(i, carNoC + 8).value)   # 包车公里
                mf.public_mileage=getFloatValue(table.cell(i, carNoC + 9).value)   # 公用公里
                mf.fault_times=getFloatValue(table.cell(i, carNoC + 14).value)     # 故障次数
                mf.fault_minutes=getFloatValue(table.cell(i, carNoC + 15).value)   # 故障分钟
                mf.target_in_compute=get_one_car_target(route,date,mf.carInfo.cartype)
                if date.month > 5 and date.month < 10:
                    mf.target_in_compute_2 = mf.carInfo.cartype.target_value4
                else:
                    mf.target_in_compute_2 = mf.carInfo.cartype.target_value3

                mf.save()
            except:
                traceback.print_exc()
                print("--------------------------car_id",car_id,"route:",route)
                continue

def getFloatValue(s):
    result=str(s).replace('/t','')
    if result =="":
        return 0.0

    else:
        try:
            return float(result)
        except:
            return 0.0

def getIntValue(s):
    result = str(s).replace('/t', '')
    if result =="":
        return 0
    else:
        return int(result)

#保存一保二保数目
def saveMainteData(table,route,date):
    # 获取每张表格一保二保数目
    frist = 0
    second = 0
    for i in range(0, table.nrows):
        for j in range(0, table.ncols):

            string = str(table.cell(i, j).value)
            if st.contains(string, "一保"):
                result = re.findall(r"\d+\.?\d*", string)
                frist=result[0] if result !=[] else 0
            if st.contains(string, "二保"):
                result=re.findall(r"\d+\.?\d*", string)
                second = result[0] if result !=[] else 0
    try:
        rmc=RouteMonthlyDetail.objects.get(route=route,date=date)
        rmc.num_fir_maintain += getIntValue(frist)
        rmc.num_sec_maintain += getIntValue(second)
        rmc.save()
    except RouteMonthlyDetail.DoesNotExist:
        nrmc=RouteMonthlyDetail(route=route,date=date,num_fir_maintain=frist,num_sec_maintain=second)
        nrmc.save()
        return


# #读取统计数据文件
# def load_sum_data(table, item,date):
#     nrows = table.nrows
#     list = []
#     indexR = item[0]
#     carNoC = item[1]
#     for i in range(indexR, nrows):
#         if table.cell(i, carNoC).value is "":
#             continue
#         if st.contains(str(table.cell(i, carNoC).value), "小计"):
#             continue
#         try:
#
#             car_id=getCar_id(table.cell(i, carNoC).value,table.name)
#
#             if not check_ID_validty(car_id):
#                 print("route:",getRoute(table.name),"car_id:",car_id)
#             bi = BusInfo.objects.get(car_id=car_id)
#         except BusInfo.DoesNotExist:
#             bi = None
#         #if bi is not None:
#         route = getRoute(table.name)
#         mileage = getFloatValue(table.cell(i, carNoC + 1).value)  # 车公里
#         team_target = getFloatValue(table.cell(i, carNoC + 3).value)  # 车队上报指标
#         oilwear = getFloatValue(table.cell(i, carNoC + 4).value)  # 车辆油耗
#         maintain = getFloatValue(table.cell(i, carNoC + 8).value)  # 二保
#         follow = getFloatValue(table.cell(i, carNoC + 9).value)  # 跟车
#         mf = MonthlyFeedback(date=datetime.datetime(2018, 6, 1), route=route, carInfo=bi, mileage=mileage,
#                              oilwear=oilwear, maintain=maintain, follow=follow,team_target=team_target,fb_car_id=car_id)
#         #mf.save()
#         list.append(mf)
#
#     return list

#对车牌号进行判断，如果不符合标准则标记为暂定无效数据
def check_ID_validty(id):
    s = str(id)
    #若车牌号小于4位，则判定为失效
    if len(s) < 4:
        return False
    return True

#获取扫描起始位置
def getStartIndex(table,target="车号"):
    rowindex = None
    colindex = None
    nrows = table.nrows
    ncols = table.ncols
    for i in range(0, nrows):
        for j in range(0, ncols):
            if table.cell(i, j).value ==target:
                rowindex = i
                colindex = j
                break
    if rowindex is None:
        return None
    item = [rowindex + xOffset, colindex]
    return item

#获取对应的路线名称
def getRoute(tablename):
    if st.contains(tablename, "专线"):
        return "海峡专线"
    if st.contains(tablename, "夜班"):
        return "夜班一号线"
    if st.contains(tablename.lower(), "k2"):
        return "k2"
    if st.contains(tablename.lower(), "21支"):
        return "142"
    if st.contains(tablename.lower(), "30支"):
        return "149"
    route = re.findall(r"\d+\.?\d*", tablename)
    print(route)
    return route[0]

from bus_info.logic import bus_info_logic as bil
# 对车辆id进行处理
def getCar_id(id, tablename):
    s = str(id)
    if (st.contains(s,"Ｆ")):
        s=s.replace("Ｆ","F")
    if (st.contains(s, "路")):
        s = s.split("路")[1].strip()
    if (st.contains(s, "线")):
        s = s.split("线")[1].strip()
    if(st.contains(s,"/")):
        s=s.split("/")[1].strip()
    if st.contains(s, ".") or len(s) == 3:
        s = s.split(".")[0]
        if len(s)==4:
            return s
        route = getRoute(tablename)
        if route == "1" or route == "17" or route == "28" or route == "161" or route=="k2":
            s = "A" + s
        elif route == "501"or route=="200":
            s = "B" + s
    return s


# 获取当前文件最后一行
def getlastrowindex(filename,sheetname):
    data = xlrd.open_workbook(filename)
    table = data.sheet_by_name(sheetname)
    nrows = table.nrows
    return nrows

def checkRouteExist(result,name):
    for i in range(0,len(result)):
        if result[i][0]==name:
            return i
    return 0

