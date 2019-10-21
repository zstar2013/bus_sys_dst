import numpy as np
import pandas as pd
from pandas import Series, DataFrame
import datetime
import os
from tools.filetool import FileDispatcher,searchforFileWithCallback
from tools.strtool import isContainOr,contains,isContain
import tools.xlstool as xt
import xlsxwriter
from xlwt import Workbook, easyxf
from ui.style import xlsStyle

OIL_TABLE_NAME="1"
width_normal=16
width_longer = 32

#从xml文件中读取油耗信息，并导出为pandas数据
def load_xml_data(sheetname,filepath):
	xls=pd.ExcelFile(filepath)
	parsed=pd.read_excel(xls,sheetname)
	parsed.columns=parsed.iloc[2].values
	parsed=parsed.drop(parsed.index[:3])
	parsed=parsed.drop(parsed.index[-2:])
	path=os.path.dirname(filepath)
	team=path.split("\\")[-1]
	route=filepath.split("\\")[-1].replace("明细.xls","")
	parsed["team"]=team
	parsed["route"]=route
	return parsed
    
	
class OliFileDispatcher(FileDispatcher):
	def file_dispatch(self,filepath):
		filter_key = ["明细"]
		filter_type = ".xls"
		if os.path.splitext(filepath)[1] == filter_type and isContainOr(keys=filter_key, target=filepath):
			return load_xml_data(OIL_TABLE_NAME, filepath)
			
def load_oil_data(path):
	dper=OliFileDispatcher()
	item=searchforFileWithCallback(path,dper)
	table=pd.concat(item)
	table.index=range(len(table))
	table["加油量（升）"]=pd.to_numeric(table["加油量（升）"])
	s_group=table.groupby(["team","route"])["加油量（升）"]
	result_df=s_group.sum()
	return result_df

def colnum_to_name(colnum):
    if type(colnum) != int:
        return colnum
    if colnum > 25:
        ch1 = chr(colnum % 26 + 65)
        ch2 = chr(colnum / 26 + 64)
#         print ch2+ch1
        return ch2 + ch1
    else:
#         print chr(colnum % 26 + 65)
        return chr(colnum % 26 + 65)
	
#创建油料中心文件
def create_oilstation_feedback_file(path, str_date,data):
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("油料中心数据表")
    styleTitle = wb.add_format(xlsStyle.styles_xlsxwriter["table"]['table_title'])
    ws.merge_range(0, 2, 0, 0, u'各线路油耗油料中心数据', styleTitle)
    styleSubTitle = wb.add_format(xlsStyle.styles_xlsxwriter["table"]['table_subtitle'])
    ws.merge_range(1, 2, 1, 0, str_date, styleSubTitle)
    stylenormal = wb.add_format(xlsStyle.styles_xlsxwriter["table"]['table_normal'])
    current_index=3
    outer_index=data.unstack().index
    dict_={}
    for i in outer_index:
        start_=current_index
        inner_index=data[i].index
        for j in inner_index:
            ws.write(current_index, 1, j, stylenormal)
            ws.write(current_index,2, data[i][j], stylenormal)
            current_index+=1
        ws.write(current_index, 1,"汇总", stylenormal)
        ws.write_formula(current_index, 2, '=SUM({2}{0}:{2}{1})'.format(str(start_ + 1)
                                                                        , str(current_index),str(colnum_to_name(2))), stylenormal)
        ws.merge_range(start_,0 , current_index, 0, i, stylenormal)
        current_index+=1
        dict_[i]=current_index
    
    str_col=str(colnum_to_name(2))
    s_total="+".join({str_col+str(x) for x in dict_.values()})
    del dict_["营达"]
    s_1gs2="+".join({str_col+str(x) for x in dict_.values()})
    del dict_["一公司公备汇总"]
    s_1gs1="+".join({str_col+str(x) for x in dict_.values()})
    start_=current_index
    ws.write(current_index, 1, "一公司营运汇总", stylenormal)
    ws.write_formula(current_index,2, "="+s_1gs1, stylenormal)
    current_index+=1
    ws.write(current_index, 1, "一公司汇总", stylenormal)
    ws.write_formula(current_index,2, "="+s_1gs2, stylenormal)
    current_index+=1
    ws.write(current_index, 1, "全汇总", stylenormal)
    ws.write_formula(current_index,2, "="+s_total, stylenormal)
    ws.merge_range(start_,0 , current_index, 0, "汇总", stylenormal)
    current_index+=1
    
    ws.write(2, 0, u'车队', stylenormal)
    ws.write(2, 1, u'线路', stylenormal)
    ws.write(2, 2, u'油耗', stylenormal)
    ws.set_column("A:B", width_normal)
    ws.set_column("C:C", width_longer)
    wb.close()