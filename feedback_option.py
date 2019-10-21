import feedback_data as fd
import loadexcel as lc
import loaddb as ld
from tools import filetool as ft
import pandas as pd
from xlwt import Workbook


#将F改成正常拼写
replaceF=lambda x:x.replace("Ｆ","F")
#如果以F和D结尾，则补齐6位
zfill_ = lambda x: x.zfill(6) if x.endswith("F")|x.endswith("D") else x
#增加F末尾并补齐
zfill_F = lambda x: "".join([x.zfill(5),"F"])
#增加D末尾并补齐
zfill_D = lambda x: "".join([x.zfill(5),"D"])
#增加A开头
zfill_a=lambda x: "".join(["A",x])
zfill_b=lambda x: "".join(["B",x])
#去掉空格
replaceBk=lambda x:x.replace(" ","")
#去掉路
remove_route=lambda x:x.split("路")[-1].strip()
#去掉线
remove_line=lambda x:x.split("线")[-1].strip()
#去掉斜杆
remove_diagonal=lambda x:x.split("/")[-1].strip()



def execute(source_path,output_path,date,target_by):
	cleaned=lc.excel2Dataframe(path=source_path,target_by=target_by,date_=date)
	#合并list转为dataframe
	data=[]
	#[data.extend(item) for item in cleaned]
	data=pd.concat(data)
	car_value=ld.get_car_info()
	data.index=range(len(data))
	#将所选列转为字符
	data.iloc[:,0]=data.iloc[:,0].apply(str)
	#将所选列转为剔除空格
	data.iloc[:,0]=data.iloc[:,0].str.strip()
	
	data["car_id"]=data.iloc[:,0].apply(replaceF)
	data["car_id"]=data.iloc[:,0].apply(zfill_)
	
	#连接油耗数据与指标
	result_data=pd.merge(data,car_value,left_on='car_id',right_on='sub_car_id',how='left')
	outer_result=result_data[pd.isnull(result_data['sub_car_id'])]
	
	#如果outer_reuslt有值，打印出outer_result,并报错
	if len(outer_result) !=0:
		
	
		#result_data=result_data.drop(outer_result.index)
		#根据日期获取油和电的指标和总指标
		result_data["target_oil_cost"]=result_data["target_value2" if date.month in range(6,10) else "target_value1"]
		result_data["target_elc_cost"]=result_data["target_value4" if date.month in range(6,10) else "target_value3"]
		result_data["total_oil_target"]=result_data["mileage"]*result_data["target_oil_cost"]/100
		result_data["total_elc_target"]=result_data["mileage"]*result_data["target_elc_cost"]/100
		#将index设置为route
		#result_data.index=result_data['route']
		
		teams=result_data["team"].drop_duplicates().values
		for team in teams:
			newWb = Workbook()
			routes=result_data[result_data["team"]==team]["route"].drop_duplicates().values
			for route in routes:
				ws = newWb.add_sheet(str(route)+"路统计表")
				fd.write_car_oil_cost_sum(ws,date,result_data[result_data["route"]==route],team=team,route=route)
			newWb.save("G:\\pythonproject\\bus_sys_dst\\output\\"+team+".xls")