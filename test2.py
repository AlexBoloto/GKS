import pandas as pd
data_frame = pd.read_excel('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\result\\result_3.xlsx', usecols=[0, 3])
data_frame2 = data_frame[['Pokazatel', 'form']]
data_frame3 = data_frame2[~data_frame2['Pokazatel'].str.contains('Наименовани|Код|Примечан|____|стр|Стр', regex=True, na = False)]
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"^\d.?\s",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"^\s\d+$",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"\d+.\d+.?\s",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"\d.\d+.\d+?.?",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"-\s+\d+,?\d+,",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"^\d.\d.",'', regex=True)
data_frame3['Pokazatel'] = data_frame3['Pokazatel'].str.replace(r"^\d+$",'', regex=True)
data_frame3.dropna(inplace=True)
data_frame3.to_excel('C:\\Users\\MikhaylovAV1\\Documents\\GKS_forms_1\\result\\result_5.xlsx', index=False)
data_frame3 = data_frame3[~data_frame3['Pokazatel'].str.contains('^$', regex = True, na = False)]
print(data_frame3['Pokazatel'])

