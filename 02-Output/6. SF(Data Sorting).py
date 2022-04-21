import pandas as pd
import os
import matplotlib.pyplot as plt

#%% 사용자 입력

### Analysis Result 정보불러오기
data_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101' # data 폴더 경로
data_xlsx = 'Analysis Result' # 파일명

# Output 경로 설정
SF_output_xlsx_dir = data_path
SF_output_xlsx = '3.Results_Wall.xlsx'

# # Story 정보 경로 설정
# input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
# input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.2_220214.xlsx'

#%% Analysis Result 불러오기

to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if (data_xlsx in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

# 전단력 불러오기
shear_force_data = pd.DataFrame()

for i in to_load_list:
    shear_force_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Structure Section Forces', skiprows=2, usecols=[0,3,5,6,7,8])
    shear_force_data = shear_force_data.append(shear_force_data_temp)
    
shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)', 'V(kN)']

# 필요없는 전단력 제거(층전단력)
if shear_force_data['Name'].str.contains('_').any() == True: # 이 줄은 없어도 될거같은디..
    shear_force_data = shear_force_data[shear_force_data['Name'].str.contains('_') == True]
    
shear_force_data.reset_index(inplace=True, drop=True)

#%% 부재명, H1, H2 값 뽑기

# 지진파 이름 list 만들기
load_name_list = []
for i in shear_force_data['Load Case'].drop_duplicates():
    new_i = i.split('+')[1]
    new_i = new_i.strip()
    load_name_list.append(new_i)

gravity_load_name = [x for x in load_name_list if ('DE' not in x) and ('MCE' not in x)]
seismic_load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]

seismic_load_name_list.sort()

# 데이터 Grouping
shear_force_H1_data_grouped = pd.DataFrame()
shear_force_H2_data_grouped = pd.DataFrame()

for load_name in seismic_load_name_list:
    shear_force_H1_data_grouped['{}_H1_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                  (shear_force_data['Step Type'] == 'Max')]['H1(kN)'].values
        
    shear_force_H1_data_grouped['{}_H1_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                  (shear_force_data['Step Type'] == 'Min')]['H1(kN)'].values

for load_name in seismic_load_name_list:
    shear_force_H2_data_grouped['{}_H2_max'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                  (shear_force_data['Step Type'] == 'Max')]['H2(kN)'].values
        
    shear_force_H2_data_grouped['{}_H2_min'.format(load_name)] = shear_force_data[(shear_force_data['Load Case'].str.contains('{}'.format(load_name))) &\
                                                                  (shear_force_data['Step Type'] == 'Min')]['H2(kN)'].values   

# all 절대값
shear_force_H1_abs = shear_force_H1_data_grouped.abs()
shear_force_H2_abs = shear_force_H2_data_grouped.abs()

# 최대값 every 4 columns
shear_force_H1_max = shear_force_H1_abs.groupby([[i//4 for i in range(0,56)]], axis=1).max()
shear_force_H2_max = shear_force_H2_abs.groupby([[i//4 for i in range(0,56)]], axis=1).max()

# 1.2 * 평균값
shear_force_H1_avg = 1.2 * shear_force_H1_max.groupby([[i//7 for i in range(0,14)]], axis=1).mean()
shear_force_H2_avg = 1.2 * shear_force_H2_max.groupby([[i//7 for i in range(0,14)]], axis=1).mean()

shear_force_H1_avg.columns = ['1.2 DE', '1.2 MCE']
shear_force_H2_avg.columns = ['1.2 DE', '1.2 MCE']

#%% V(축력) 값 뽑기

# 축력 불러와서 Grouping
axial_force_data = shear_force_data[shear_force_data['Load Case'].str.contains(gravity_load_name[0])]['V(kN)']

# 절대값
axial_force_abs = axial_force_data.abs()

# result
axial_force_abs.reset_index(inplace=True, drop=True)
axial_force = axial_force_abs.groupby([[i//2 for i in range(0, len(axial_force_abs))]], axis=0).max()

#%% 출력

# 출력용 Dataframe 만들기
SF_output = pd.DataFrame()
SF_output['Name'] = shear_force_data['Name'].drop_duplicates()
SF_output.reset_index(inplace=True, drop=True)

SF_output['Nu'] = axial_force
SF_output['1.2_DE_H1'] = shear_force_H1_avg['1.2 DE'].values
SF_output['1.2_DE_H2'] = shear_force_H2_avg['1.2 DE'].values
SF_output['1.2_MCE_H1'] = shear_force_H1_avg['1.2 MCE'].values
SF_output['1.2_MCE_H2'] = shear_force_H2_avg['1.2 MCE'].values

# 엑셀로 출력
SF_output.to_excel(SF_output_xlsx_dir+ '\\'+ SF_output_xlsx, sheet_name = 'Results_Wall', index = False)