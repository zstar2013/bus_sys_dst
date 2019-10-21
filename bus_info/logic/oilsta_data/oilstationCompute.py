import xlrd
import os
import tools.xlstool as xt
from tools.filetool import FileDispatcher
from tools.strtool import contains,isContain
class OliFileDispatcher(FileDispatcher):
    def file_dispatch(self,filepath):
        filter_key = "明细"
        filter_type = ".xls"
        if os.path.splitext(filepath)[1] == filter_type and isContain(key=filter_key, target=filepath):
            data = xlrd.open_workbook(filepath)
            table = data.sheet_by_name("1")
            return load_xml_data(table, filepath)

def load_xml_data(table,filepath):
    # 获取当前表格的行数
    nrows = table.nrows
    # 获取当前表格的列数
    ncols = table.ncols
    item = {}
    item["team"] =str.split(filepath,'\\')[-2]
    item["route"] = os.path.split(filepath)[1][0:-6].replace("路","")
    totalCharge=0
    index = xt.getStartIndex(table,"交易时间")
    for i in range(index, nrows):

        if table.cell(i, 0).value is "":
            continue
        if contains(table.cell(i, 0).value, "合计"):
            continue
        charge_=table.cell(i, 6).value
        ocd=OilChargeDetail(car_id=str(table.cell(i, 3).value),
                            route=item["route"],
                            paytime=str(table.cell(i,0).value),
                            charge = charge_,
                            printnum=str(table.cell(i, 2).value),
                            station=str(table.cell(i, 5).value),
                            team=item["team"])
        ocd.save()
        totalCharge+=float(charge_)
    item["charge"]=totalCharge
    return item

