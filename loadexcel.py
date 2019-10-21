import numpy as np
import pandas as pd
from pandas import Series, DataFrame
import xlrd
from tools.filetool import FileDispatcher
import datetime
import os
from tools.strtool import isContainOr,contains
import re
class xlsDispatcher(FileDispatcher):
    current_date=None
    def __init__(self,date,path):
        self.current_date=date
        self.p_index=len(os.path.split(path)[0].split("//"))+1


    def file_dispatch(self,filepath):
        filter_key = ["油耗和汇总表","2019.","每月油耗 (version 1)","油表","油耗汇总表"]
        filter_table_keys = ["统计"]
        filter_table_keys_detail=["汇总"]
        filter_type = ".xls"
        paths_sum=[]
        paths_detail=[]
        items=[]
        items_detail=[]
        if os.path.splitext(filepath)[1] == filter_type and isContainOr(keys=filter_key, target=filepath):
            data = xlrd.open_workbook(filepath)
            for sheet in data.sheets():
                if isContainOr(filter_table_keys,sheet.name):
                    paths_sum.append({"filepath":filepath,"sheetname":sheet.name})
                    #items.append(excel2Dataframe_Sum(filepath,self.p_index,sheet.name,"车号"))
                if isContainOr(filter_table_keys_detail,sheet.name):
                    paths_detail.append({"filepath":filepath,"sheetname":sheet.name})
                    #details=excel2Dataframe_Detail(filepath,self.p_index,sheet.name,"车号")
                    #if details is not None:
                        #items_detail.append(details)
                        
        return (paths_sum,paths_detail)

"""
in        data 原数据dataframe
        str_target查找的目标字符 
        bool_contain 是否包含，默认为false
#out     定位列表[row,col]
"""
def local_value(data,str_target,bool_contain=False):
    target_index=None
    for indexs in data.index:
        for  i in range(len(data.loc[indexs].values)):
            if str_target in str(data.loc[indexs].values[i]):
                target_index=[indexs,i]
                break
    return target_index
        
"""
读取统计表
"""
def excel2Dataframe_Sum(path,p_index,sheetname,target_by):
    xls=pd.ExcelFile(path)
    parsed=pd.read_excel(xls,sheetname)
    target_index=local_value(parsed,target_by,True)
    team=os.path.split(path)[0].split("//")[p_index][0]
    return tim_data(parsed,target_index,getRoute(sheetname),team)


"""
读取汇总表
"""
def excel2Dataframe_Detail(path,p_index,sheetname,target_by):
    xls=pd.ExcelFile(path)
    parsed=pd.read_excel(xls,sheetname)
    target_index=local_value(parsed,target_by,True)
    if target_index is None:
        return None
    team=os.path.split(path)[0].split("//")[p_index][0]
    return tim_detail_data(parsed,target_index,getRoute(sheetname),team)    

                
def excel2Dataframe(path,target_by,date_):
    dper=xlsDispatcher(date_,path)
    return     searchforFileWithCallback(path,dper)
    
    

#规整统计数据
def tim_data(data,target_index,route,team):
    #删掉纵轴无关数据
    new_obj=data.drop(data.columns[:target_index[1]],axis=1)
    new_obj=new_obj.drop(data.columns[target_index[1]+11:],axis=1)
    
    #删掉横轴无关数据
    _index1=new_obj.index[:target_index[0]+3]
    _index2=new_obj[new_obj[new_obj.columns[0]].isnull()|new_obj[new_obj.columns[0]].str.contains('总计')|new_obj[new_obj.columns[0]].str.contains('小计：')].index
    new_obj=new_obj.drop(_index1|_index2)
    # 如果列数小于11，增加inspection列并赋值为0，否则直接赋值为第11列
    if len(new_obj.columns)<11:
        new_obj["inspection"]=0.0
    new_obj.columns=list(range(len(new_obj.columns)))
    new_obj=new_obj.rename(columns={0:'car_id',1:'mileage',4:'oil_cost',8:'maintain',9:'follow',10:'inspection'})
    new_obj['route']=route
    new_obj['team']=team
    new_obj=new_obj[['car_id','mileage','oil_cost','maintain','follow','inspection','route','team']].fillna(0)
    cleaned=new_obj.replace("二保",0)
    return cleaned
    
#规整汇总数据
def tim_detail_data(data,target_index,route,team):
    #删掉纵轴无关数据
    new_obj=data.drop(data.columns[:target_index[1]],axis=1)
    new_obj=new_obj.drop(data.columns[target_index[1]+18:],axis=1)
    
    #删掉横轴无关数据
    _index1=new_obj.index[:target_index[0]+3]
    rule1=new_obj[new_obj.columns[0]].isnull()
    rule2=new_obj[new_obj.columns[0]].str.contains('计')
    rule3=new_obj[new_obj.columns[0]].str.contains('一保')
    rule4=new_obj[new_obj.columns[0]].str.contains('备注')
    rule5=new_obj[new_obj.columns[0]]==0
    _index2=new_obj[rule1|rule2|rule3|rule4|rule5].index
    new_obj=new_obj.drop(_index1|_index2)
    new_obj.columns=list(range(len(new_obj.columns)))
    new_obj=new_obj.rename(columns={0:'car_id',2:'fix_days',4:'stop_days',5:'work_days',8:'engage_mileage',9:'public_mileage',10:'shunt_mileage',
                                        14:'fault_times',15:'fault_minutes'
    })
    new_obj['route']=route
    new_obj['team']=team
    new_obj=new_obj[['car_id','fix_days','stop_days','work_days','engage_mileage','public_mileage','shunt_mileage',
                                        'fault_times','fault_minutes','route','team']].fillna(0)
    return new_obj
    
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
    
#规整数据
def tim_data_test(data,target_index,route):
    dict_={'车号':'car_id','车公里':'mileage','实绩':'oil_cost','二保':'maintain','跟车':'follow'}
    #获取每个所需列所在的索引的列表
    list_loc_x=[local_value(data,key)[1] for key in  dict_.keys()]
    #筛选列表数据
    new_obj=data.iloc[:,list_loc_x]
    #重新设置列名
    new_obj.columns=dict_.values()
    #删除车号行
    new_obj=new_obj[new_obj["car_id"]!="车号"]
    #删除第一列为nan的行，其余nan值赋值为零
    new_obj=new_obj[pd.notnull(new_obj["car_id"])].fillna(0) 
    #替换无用值
    cleaned=new_obj.replace("二保",0)
    #设置线路
    cleaned["route"]=route
    return cleaned
    
    
#检索符合条件的文件并处理，将结果返回列表items
def searchforFileWithCallback(path,dper):
    items = []
    if os.path.isdir(path):
        dirs = os.listdir(path)
        for file in dirs:
            _filepath = path + "//" + str(file)
            if os.path.isdir(_filepath):
                item=searchforFileWithCallback(path=_filepath,dper=dper)
                if item is not None:
                    items += item
            if os.path.isfile(_filepath):
                item=dper.file_dispatch(filepath=_filepath)
                if item is not None and len(item)>0:
                    items.append(item)
    elif os.path.isfile(path):
        item = dper.file_dispatch(filepath=path)
        if len(item) > 0:
            items.append(item)
    else:
        return None
    return items
    
    
#获取对应的路线名称
def getRoute(tablename):
    if contains(tablename, "专线"):
        return "海峡专线"
    if contains(tablename, "夜间"):
        return "夜班一号线"
    if contains(tablename, "夜班"):
        return "夜班一号线"
    if contains(tablename.lower(), "k2"):
        return "k2"
    if contains(tablename.lower(), "21支"):
        return "142"
    if contains(tablename.lower(), "30支"):
        return "149"
    if contains(tablename.lower(), "57路区间"):
        return "57区间"
    route = re.findall(r"\d+\.?\d*", tablename)
    #print(route)
    return route[0]
    
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
    return s

def update_car_info_from_xml(path,sheetname,targetstr,date):
    xls=pd.ExcelFile(path)
    parsed=pd.read_excel(xls,sheetname)
    target_index=local_value(parsed,target_by,True)