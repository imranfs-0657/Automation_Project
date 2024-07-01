from thinkcellbuilder import Presentation, Template
import pandas as pd
from datetime import datetime
from builder import Builder
from thinkcell import Thinkcell

file_name_1 = r"storage\20240528_Weekly_Leads_Summary_0525_v3.xlsb" 
sheet_name_1 = 'By Marketing Channel (TEMPLATE)'
df1 = Builder().read_excel(file_name_1, sheet_name_1)

custom_column_names_df1 = Builder().generate_columns(df1.shape[1])

df1.columns = custom_column_names_df1

data_for_chart1 = Builder().extract_data(df1, 'C', 'P', 20, 26)

data_for_chart1 = Builder().add_row(df1,data_for_chart1,28,'C','P','D')

data_for_chart1 = Builder().add_row(df1, data_for_chart1, 52, 'C', 'P', 'D')

updated_column_names = Builder().dates(df1,18, 'D','P')

converted_updated_column_names = Builder().convert_to_date_time(updated_column_names)

formated_updated_column_names = [Builder().format_date_time(d) for d in converted_updated_column_names]

data_for_chart1.columns = [data_for_chart1.columns[0]]+formated_updated_column_names




# For Chart 2

data_for_chart2 = Builder().extract_data(df1, 'K', 'P', 32, 38)

column_list = Builder().add_column(df1, 'D', 32,39 )

data_for_chart2.insert(loc=0,column = 'D', value = column_list )

column_names = Builder().add_column(df1, 'C', 32, 39)

data_for_chart2.insert(loc=0,column='C',value=column_names)

updated_column_names_chart2 = Builder().dates(df1,30, 'K','P')
updated_column_names1_chart2 = df1.loc[30, "D"]
updated_column_names1_chart2_list = [updated_column_names1_chart2]

converted_updated_column_names_chart2 = Builder().convert_to_date_time(updated_column_names_chart2)
converted_updated_column_names1_chart2 = Builder().convert_to_date_time(updated_column_names1_chart2_list)

formated_updated_column_names_chart2 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart2]
formated_updated_column_names1_chart2 = [Builder().format_date_time(d) for d in converted_updated_column_names1_chart2]

data_for_chart2.columns = [data_for_chart2.columns[0]]+formated_updated_column_names1_chart2+formated_updated_column_names_chart2



# For Chart3

data_for_chart3 = Builder().extract_data(df1, 'K', 'P', 60, 61)

data_for_chart3 = Builder().add_row(df1, data_for_chart3, 64, 'K','P',None)
data_for_chart3 = Builder().add_row(df1,data_for_chart3,65,'K','P',None)

column_list_chart3 = Builder().add_column(df1, 'D', 60,62 )

column_list_chart3.append(df1.loc[64,'D'])
column_list_chart3.append(df1.loc[65,'D'])


data_for_chart3.insert(loc=0,column = 'D', value = column_list_chart3 )

column_names_chart3 = Builder().add_column(df1, 'C', 60, 62)
column_names_chart3.append(df1.loc[64,'C'])
column_names_chart3.append(df1.loc[65,'C'])

data_for_chart3.insert(loc=0,column='C',value=column_names_chart3)

length = len(data_for_chart3.loc[60])
for i in range(1,length):
    data_for_chart3.iloc[:, i] = data_for_chart3.iloc[:, i].apply(lambda x: float(f"{x * 100:.1f}"))

updated_column_names_chart3 = Builder().dates(df1,58, 'K','P')
updated_column_names1_chart3 = df1.loc[58, "D"]
updated_column_names1_chart3_list = [updated_column_names1_chart3]

converted_updated_column_names_chart3 = Builder().convert_to_date_time(updated_column_names_chart3)
converted_updated_column_names1_chart3 = Builder().convert_to_date_time(updated_column_names1_chart3_list)

formated_updated_column_names_chart3 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart3]
formated_updated_column_names1_chart3 = [Builder().format_date_time(d) for d in converted_updated_column_names1_chart3]

data_for_chart3.columns = [data_for_chart3.columns[0]]+formated_updated_column_names1_chart3+formated_updated_column_names_chart3


#For Chart4

data_for_chart4 = Builder().extract_data(df1, 'C', 'P', 81, 86)

data_for_chart4 = data_for_chart4.drop(index = 82)

updated_column_names_chart4 = Builder().dates(df1,79, 'D','P')

converted_updated_column_names_chart4 = Builder().convert_to_date_time(updated_column_names_chart4)

formated_updated_column_names_chart4 = [Builder().format_date_time(d) for d in converted_updated_column_names_chart4]

data_for_chart4.columns = [data_for_chart4.columns[0]]+formated_updated_column_names_chart4




#Updating chart1

chart_name = "Demand Pacing - Monthly and Weekly - 1"
dataframe = data_for_chart1
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name, dataframe, output_file_name)

#Updating Chart2

chart_name2 = "Demand Pacing - Monthly and Weekly - 2"
dataframe2 = data_for_chart2
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name2, dataframe2, output_file_name)

#Updating Chart3

chart_name3 = "Demand Pacing - Monthly and Weekly - 3"
dataframe3 = data_for_chart3
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name3, dataframe3, output_file_name)

#Updating Chart4

chart_name4 = "Demand Pacing - Monthly and Weekly - 4"
dataframe4 = data_for_chart4
output_file_name = "APR Month End_Digital Performance Update - Copy_Factspan_May (2).ppttc"

Thinkcell().update_chart(chart_name4, dataframe4, output_file_name)

