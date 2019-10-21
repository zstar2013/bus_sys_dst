import sqlalchemy as sqla
import pandas as pd


def get_db():
	return sqla.create_engine('sqlite:///example_data//django_bus_sys')
	
def get_car_info():
	db=get_db()
	car_info=pd.read_sql('select sub_car_id,cartype_id,car_id,route from bus_info_businfo where scrap=0', db)
	type_info=pd.read_sql('select id,target_value1,target_value2,target_value3,target_value4,power_type from bus_info_cartype', db)
	return pd.merge(car_info,type_info,left_on="cartype_id",right_on="id",how='left')