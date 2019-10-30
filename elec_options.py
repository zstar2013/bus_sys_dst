#"example_data\\electric\\8月\\单车能耗汇总-公交集团8月.xlsx", "一公司", "车队"
#"example_data\\electric\\8月\\福州公交八月份  83、161、K2线路车辆明细统计表.xlsx", "福州公交 八月份  83、161、K2线路车辆明细统计表", "车队", date=date
#"example_data\\electric\\8月\\华威8月.xlsx", "公交汇总", "车队",

def tim_data_elec_xny(data,vars=[7,13,"新能源"]):
    car_index,elec_index,source=vars
    #删掉纵轴无关数据
    result_1=data[data.columns[[car_index,elec_index]]]
    result_1.index=range(len(result_1))
    index_null=result_1[result_1[result_1.columns[0]].isnull()].index
    result_1=result_1.drop(index_null)
    index_wrong=result_1[result_1[result_1.columns[0]].str.contains("车号")].index
    index_wrong2=result_1[result_1[result_1.columns[0]].str.contains("合计")].index
    index_wrong3=result_1[result_1[result_1.columns[0]].str.contains("车辆标识")].index
    result_1=result_1.drop(index_wrong|index_wrong2|index_wrong3)
    #将F改成正常拼写
    replace_A_Y=lambda x:x.replace("A-Y","AY")
    result_1.iloc[:,0]=result_1.iloc[:,0].apply(replace_A_Y)
    result_1.columns=["车号","电量"]
    result_1["来源公司"]=source
    return result_1