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

def column_to_str(df,col_id):
	df.iloc[:,col_id]=df.iloc[:,col_id].apply(str)

	
def check_car_id(df_source,df_car,df_uncheck):
	for i in range(len(df_uncheck)):
		_car_id=unturn_tmp["car_id_x"].iloc[i]
		match_result=car_value[car_value['sub_car_id'].str.contains(_car_id)]
		if len(match_result)==1:#唯一匹配
			checked_id=match_result.iloc[0]['sub_car_id']
			df_source.iloc[unturn_tmp.iloc[0].name,0]=checked_id
		elif len(match_result)==0:#不存在匹配
			print(_car_id,"该车号不存在")
		
		else:#存在多个匹配
			#todo

		
	