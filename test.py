import pandas as pd
pd.set_option('display.max_columns', 8)
pd.set_option('display.width', 80)
data_frame = pd.read_excel('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\result\\result_1.xlsx')
#data_frame.dropna(inplace=True)
data_frame = data_frame[~((data_frame['Pokazatel'].str.islower()) | (data_frame['Pokazatel'].str.len() < 3) | (data_frame['Pokazatel'].str.isnumeric()))]
data_frame.to_excel("C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\Result\\Result_3.xlsx", header=True, index=False)
data_frame2 = data_frame[['Pokazatel','form']]
data_frame3 = data_frame2[~(data_frame2['Pokazatel'].str.contains('Наименовани|Код|Примечан|____', regex = True, na=False))]
print(data_frame3)
data_frame3.to_excel("C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\Result\\Result_4.xlsx", header=True, index=False)
#data_frame=data_frame[data_frame['Pokazatel'].str.replace(r'^\d*)', '', regex=True)]