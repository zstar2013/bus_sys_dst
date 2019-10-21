#TODO 编写元组（location,value）,简化代码

backup_cars_num=("109ke","902ka","871ka","569kd","713ka","712ka","091KD","F8G23"
                 )

mergeCells_base_sum=[{"location":[0, 0, 0, 12], "value": "车公里单车燃料消耗统计表", "style": "table_title_without_borders"},
                     {"location":[1,1,0,12],"value":"","style":"table_subtitle_without_borders"},
                     {"location":[2,4,0,0],"value":"单位","style":"table_normal"},
                     {"location":[2,4,1,1],"value":"线路","style":"table_normal"},
                     {"location":[2,4,2,2],"value":"车号","style":"table_normal"},
                     {"location":[2,4,3,3],"value":"车公里","style":"table_normal"},
                     {"location":[2,2,4,7],"value":"柴油消耗(升)","style":"table_normal"},
                     {"location":[3,3,4,5],"value":"指标","style":"table_normal"},
                     {"location":[3,3,6,7],"value":"实绩","style":"table_normal"},
                     {"location":[4,4],"value":"总量","style":"table_normal"},
                     {"location":[4,5],"value":"百公里","style":"table_normal"},
                     {"location":[4,6],"value":"总量","style":"table_normal"},
                     {"location":[4,7],"value":"百公里","style":"table_normal"},
                     {"location":[2, 4, 8, 8],"value":"节约","style":"table_normal"},
                     {"location":[2, 4, 9, 9],"value":"超耗","style":"table_normal"},
                     {"location":[2, 3, 10, 12],"value":"备注","style":"table_normal"},
                     {"location":[4, 10],"value":"二保","style":"table_normal"},
                     {"location":[4, 11],"value":"跟车","style":"table_normal"},
                     {"location":[4, 12],"value":"年审","style":"table_normal"}
                          ]

mergeCells_base_detail=[{"location":[0, 0, 0, 18], "value": "单车运行情况汇总表", "style": "table_title_without_borders"},
                        {"location":[1, 1, 0, 5], "value": "福州市公交第一公司", "style": "table_subtitle_without_borders"},
                        {"location":[1, 1, 6, 18], "value": "   第1页", "style": "table_subtitle_without_borders"},
                        {"location":[2, 4, 0, 0], "value": "车号", "style": "table_normal"},
                        {"location":[2, 2, 1, 5], "value": "车日运用", "style": "table_normal"},
                        {"location":[3, 4, 1, 1], "value": "营运\n车日", "style": "table_normal"},
                        {"location":[3, 4, 2, 2], "value": "修理\n车日", "style": "table_normal"},
                        {"location":[3, 4, 3, 3], "value": "完好\n车日", "style": "table_normal"},
                        {"location":[3, 3, 4, 5], "value": "其中", "style": "table_normal"},
                        {"location":[4, 4], "value": "停驶车日", "style": "table_normal"},
                        {"location":[4, 5], "value": "工作车日", "style": "table_normal"},
                        {"location":[2, 2, 6, 10], "value": "车公里", "style": "table_normal"},
                        {"location":[3, 4, 6, 6], "value": "合计", "style": "table_normal"},
                        {"location":[3, 3, 7, 10], "value": "其中", "style": "table_normal"},
                        {"location":[4, 7], "value": "营业", "style": "table_normal"},
                        {"location":[4, 8], "value": "包车", "style": "table_normal"},
                        {"location":[4, 9], "value": "公用", "style": "table_normal"},
                        {"location":[4, 10], "value": "调车", "style": "table_normal"},
                        {"location":[2, 2, 11, 15], "value": "营运效率", "style": "table_normal"},
                        {"location":[3, 4, 11, 11], "value": "完好车率", "style": "table_normal"},
                        {"location":[3, 4, 12, 12], "value": "工作车率", "style": "table_normal"},
                        {"location":[3, 4, 13, 13], "value": "故障率", "style": "table_normal"},
                        {"location":[3, 3, 14, 15], "value": "故障", "style": "table_normal"},
                        {"location":[4, 14], "value": "次", "style": "table_normal"},
                        {"location":[4, 15], "value": "分", "style": "table_normal"},
                        {"location":[2, 3, 16, 17], "value": "柴油消耗(升)", "style": "table_normal"},
                        {"location":[4, 16], "value": "指标", "style": "table_normal"},
                        {"location":[4, 17], "value": "实际", "style": "table_normal"},
                        {"location":[2, 4, 18, 18], "value": "备注", "style": "table_normal"},
                        ]

mergeCells_base_sum_tail=[{"location":[0, 0, 1, 2], "value": "机关后勤车", "style": "table_normal"},
                          {"location":[0, 0, 3, 4], "value": "汽油(购入金额)", "style": "table_normal"},
                          {"location":[0, 0, 5, 6], "value": "公备车", "style": "table_normal"},
                          {"location":[0, 0, 7, 8], "value": "柴油", "style": "table_normal"},
                          {"location":[0, 9], "value": "", "style": "table_normal"},
                          {"location":[0, 10], "value": "", "style": "table_normal"},
                          {"location":[0, 11], "value": "", "style": "table_normal"},
                          {"location":[0, 12], "value": "", "style": "table_normal"},
                          ]

sum_query_fields=['carInfo__sub_car_id','mileage','oilwear','maintain','follow','inspection']
detail_query_fields=['mileage','work_days','fix_days','stop_days','shunt_mileage','engage_mileage','public_mileage','fault_times','fault_minutes','oilwear','carInfo__sub_car_id']