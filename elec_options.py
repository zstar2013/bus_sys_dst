#"example_data\\electric\\8月\\单车能耗汇总-公交集团8月.xlsx", "一公司", "车队"
#"example_data\\electric\\8月\\福州公交八月份  83、161、K2线路车辆明细统计表.xlsx", "福州公交 八月份  83、161、K2线路车辆明细统计表", "车队", date=date
#"example_data\\electric\\8月\\华威8月.xlsx", "公交汇总", "车队",
def tim_data_elec_xny(data,vars=[6,7,13,"新能源"]):
    route_index,car_index,elec_index,source=vars
    #删掉纵轴无关数据
    result=data[data.columns[[car_index,elec_index,route_index]]]
    result.index=range(len(result))
    index_null=result[result[result.columns[0]].isnull()].index
    result=result.drop(index_null)
    index_wrong=result[result[result.columns[0]].str.contains("车号")].index|result[result[result.columns[0]].str.contains("车牌号")].index
    index_wrong2=result[result[result.columns[0]].str.contains("合计")].index
    index_wrong3=result[result[result.columns[0]].str.contains("车辆标识")].index
    result=result.drop(index_wrong|index_wrong2|index_wrong3)
    #将A_Y改成AY
    replace_A_Y=lambda x:x.replace("A-Y","AY")
    #去掉空格
    replaceBk=lambda x:x.replace(" ","")
    #去掉队
    remove_team=lambda x:x.split("队")[-1].strip()
    remove_AY=lambda x:x.replace("闽AY","")
    remove_A=lambda x:x.replace("闽A","")
    replace_K=lambda x:x.replace("K2","k2")
    remove_load=lambda x:x.replace("路","")
    result.iloc[:,0]=result.iloc[:,0].apply(replace_A_Y)
    result.columns=["车号","电量","线路"]
    result["来源公司"]=source 
    result=result.drop(result[result["电量"]==0].index)
    result.iloc[:,0]=result.iloc[:,0].apply(remove_AY)
    result.iloc[:,0]=result.iloc[:,0].apply(remove_A)
    #result.iloc[:,1]=result.iloc[:,1].apply(Decimal)
    
    #填充线路
    result.fillna(method="ffill",inplace =True)
    
    result.iloc[:,2]=result.iloc[:,2].apply(str)
    result.iloc[:,2]=result.iloc[:,2].apply(replaceBk)
    result.iloc[:,2]=result.iloc[:,2].apply(remove_team)
    result.iloc[:,2]=result.iloc[:,2].apply(remove_load)
    result.iloc[:,2]=result.iloc[:,2].apply(replace_K)
    return result
    
def output_electric_cost_list(data):
    newWb = Workbook()
    ws = newWb.add_sheet("反馈汇总表")
    fd.write_monthy_feedback_sum_table(ws, date,result_data)
    ws_ = newWb.add_sheet("单车汇总表")
    fd.write_monthy_feedback_detail_table(ws_, date,result_data)
    createdir("output\\2019\\反馈数据\\8月\\")
    newWb.save("output\\2019\\反馈数据\\8月\\8月反馈汇总表.xls")