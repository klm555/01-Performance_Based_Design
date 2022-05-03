import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
### Idea sketch
'''
1. AS_result_data 에다가 AS_gage_data 에서 i-node ID 만 가져온다.
2. AS_result_data 안에는 DE11~MCE72 까지 총 28개의 지진파에 대한 min/max 값이 모두 들어있을 것
3. 그다음 wall_gage_data 에 있을 i-node ID 와 j-node ID 를 각각 AS_result_data의 i-node ID 와 매칭시킨다.
4. 그러면, 각 벽체가 양쪽 AS gage와 매치되게 된다.
5. 그리고 나서 wall_gage_data 에 height 정보를 입력시킨 wall_gage_total을 만든다.
6. wall_gage_total 가지고 그래프 그리면 끝!
'''
### Initial Setting
# 지진파 개수
DE_num = 14
MCE_num = 14

# 허용기준
min_criteria_DE = -0.002
max_criteria_DE = 0.002
min_criteria_MCE = -0.004/1.2
max_criteria_MCE = 0.004/1.2

### ECHO File/Gage data/Gage result/Node Coordinate data/Story Info 불러오기
data_path = r'C:\Users\khpark\Desktop\21-RM-513 광명 4R구역\퍼폼 모델링\110동\Output\data'
Input_sheets_name = '110D_Input Sheets(Ver.2).xlsx'

# # ECHO File
# ECHO_path = data_path + '\\' + 'ECHO_110.xlsx'
# ECHO_data = pd.read_excel(ECHO_path)

# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\khpark\Desktop\21-RM-513 광명 4R구역\퍼폼 모델링\110동\Output\data'
input_raw_xlsx = '110D_Input Sheets(Ver.2).xlsx'

# 생성할 부록 경로 설정
SF_word_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
SF_word_name = '지진파별 벽체 축변형률(test).docx'

#%% Analysis Result 불러오기
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if 'Analysis Result' in file_name and '~$' not in file_name:
        to_load_list.append(file_name)

# AS Gage data
AS_gage_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Gage Data - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 3, 7, 9])

# AS Gage result data
AS_result_data = pd.DataFrame()
for i in to_load_list:
    AS_result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Gage Results - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 3, 5, 7, 8, 9])
    AS_result_data = AS_result_data.append(AS_result_data_temp)

AS_result_data = AS_result_data.sort_values(by= ['Load Case', 'Element ID', 'Step Type']) # 여러개로 나눠돌릴 경우 순서가 섞여있을 수 있어 DE11~MCE72 순으로 정렬


# SWR Gage data
SWR_gage_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Gage Data - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 3, 7, 9, 11, 13]) # usecols로 원하는 열만 불러오기

# SWR result data
SWR_result_data = pd.DataFrame()
for i in to_load_list:
    SWR_result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Gage Results - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 3, 5, 7, 9, 11])
    SWR_result_data = SWR_result_data.append(SWR_result_data_temp)

# Wall Element data (처음으로 ECHO 말고 Analysis Result 에서 가져와본거)
wall_element_data = pd.DataFrame()
for i in to_load_list:
    wall_element_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Element Data - Shear Wall', skiprows=[0, 2], header=0, usecols=[3, 5, 7, 9, 11, 13])
    wall_element_data = wall_element_data.append(wall_element_data_temp)

# property name split 하기
wall_element
wall_element_data['Property Name'] = wall_element_data['Property Name'].apply(lambda x: x.split('_')[0])



# Node Coord data
Node_coord_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])

# Story Info data
story_info_xlsx_sheet = 'Story Info'
story_info = pd.read_excel(data_path + '\\' + Input_sheets_name, sheet_name='Story Info', keep_default_na=False)
story_height = story_info.loc[:, 'Height(mm)']
story_name = story_info.loc[:, 'Floor Name']

# AS_result data 에다가 AS_gage_data에서 i-node ID 매칭하기
AS_result_data = AS_result_data.join(AS_gage_data.set_index('Element ID')['I-Node ID'], on='Element ID')

### AS_total 만들기
gage_num = len(AS_gage_data)
AS_max = AS_result_data[(AS_result_data['Step Type'] == 'Max') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values # dataframe을 array로
AS_max_Inode = AS_result_data[(AS_result_data['Step Type'] == 'Max') & (AS_result_data['Performance Level'] == 1)][['I-Node ID']].values # dataframe을 array로
AS_max = AS_max.reshape(gage_num, DE_num + MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
AS_max_Inode = AS_max_Inode.reshape(gage_num, DE_num + MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
AS_max = pd.DataFrame(AS_max) # array를 다시 dataframe으로
AS_max_Inode = pd.DataFrame(AS_max_Inode)

AS_min = AS_result_data[(AS_result_data['Step Type'] == 'Min') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values
AS_min_Inode = AS_result_data[(AS_result_data['Step Type'] == 'Min') & (AS_result_data['Performance Level'] == 1)][['I-Node ID']].values
AS_min = AS_min.reshape(gage_num, DE_num + MCE_num, order='F')
AS_min_Inode = AS_min_Inode.reshape(gage_num, DE_num + MCE_num, order='F')
AS_min = pd.DataFrame(AS_min)
AS_min_Inode = pd.DataFrame(AS_min_Inode)
# 만들고 보니 AS_max_Inode랑 AS_min_Inode는 똑같음

AS_total = pd.concat([AS_max_Inode.iloc[:,0], AS_max, AS_min], axis=1)
AS_total.columns = ['I-node ID',
                     'DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']


# SWR_gage_data 좌표 및 direction 결정
SWR_gage_data = SWR_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID'); SWR_gage_data.rename({'H1' : 'i_x', 'H2' : 'i_y', 'V' : 'i_z'}, axis = 1, inplace=True)
SWR_gage_data = SWR_gage_data.join(Node_coord_data.set_index('Node ID')['H1'], on = 'J-Node ID'); SWR_gage_data.rename({'H1' : 'j_x'}, axis=1, inplace=True)
SWR_gage_data['i_x-i_y'] = SWR_gage_data.apply(lambda x: x['i_x']-x['j_x'], axis=1)
SWR_gage_data['direction'] = SWR_gage_data.apply(lambda x: 'h2'
                                                 if (abs(x['i_x-i_y']) <=15)
                                                 else 'h1', axis=1)


# wall element data 좌표 및 direction 결정
wall_element_data = wall_element_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID'); wall_element_data.rename({'H1' : 'i_x', 'H2' : 'i_y', 'V' : 'i_z'}, axis = 1, inplace=True)
wall_element_data = wall_element_data.join(Node_coord_data.set_index('Node ID')['H1'], on = 'J-Node ID'); wall_element_data.rename({'H1' : 'j_x'}, axis=1, inplace=True)
wall_element_data['i_x-i_y'] = wall_element_data.apply(lambda x: x['i_x']-x['j_x'], axis =1)
wall_element_data['direction'] = wall_element_data.apply(lambda x: 'h2'
                                                         if (abs(x['i_x-i_y'])<=15)
                                                         else 'h1', axis=1)
# SWR_gage_data와 wall_element_data 연결
SWR_gage_data = SWR_gage_data.join(wall_element_data.set_index(['I-Node ID', 'direction'])['Property Name'], on=['I-Node ID', 'direction']); SWR_gage_data.columns.values[12] = 'gage_name_full'
SWR_gage_data = SWR_gage_data.join(wall_element_data.set_index(['I-Node ID', 'direction'])['Property Name'], on=['J-Node ID', 'direction']); SWR_gage_data.columns.values[13] = 'gage_name_full_j'

for i in range(len(SWR_gage_data)):
    if pd.isnull(SWR_gage_data.iloc[i, 12]):
        SWR_gage_data.iloc[i, 12] = SWR_gage_data.iloc[i, 13]

SWR_gage_data = SWR_gage_data.iloc[:, 0:13]


# SWR_gage_data 에다가 AS total 하기
SWR_gage_data_I = SWR_gage_data.join(AS_total.set_index('I-node ID')[['DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']]
                                           , on='I-Node ID')
SWR_gage_data_I = SWR_gage_data_I.iloc[:, 12:] # 필요한 부분만 남기기
SWR_gage_data_I = pd.concat([SWR_gage_data_I, SWR_gage_data['i_z']], axis=1)
SWR_gage_data_J = SWR_gage_data.join(AS_total.set_index('I-node ID')[['DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']]
                                           , on='J-Node ID')
SWR_gage_data_J = SWR_gage_data_J.iloc[:, 12:]
SWR_gage_data_J = pd.concat([SWR_gage_data_J, SWR_gage_data['i_z']], axis=1)








# Wall element data 에다가 AS total 하기
wall_element_data_I = wall_element_data.join(AS_total.set_index('I-node ID')[['DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']]
                                           , on='I-Node ID')
wall_element_data_I = wall_element_data_I.drop(['Element ID', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID'], axis=1) # 필요없는 열 삭제
wall_element_data_J = wall_element_data.join(AS_total.set_index('I-node ID')[['DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']]
                                           , on='J-Node ID')
wall_element_data_J = wall_element_data_I.drop(['Element ID', 'I-Node ID', 'J-Node ID', 'K-Node ID', 'L-Node ID'], axis=1)

SWR_gage_data_I = SWR_gage_data_I.sort_values(['gage_name_full', 'i_z']) # 그래프가 깔끔하게 나오기 위해 벽체마다 다시 height 기준으로 재정렬
SWR_gage_data_J = SWR_gage_data_J.sort_values(['gage_name_full', 'i_z']) # 그래프가 깔끔하게 나오기 위해 벽체마다 다시 height 기준으로 재정렬

# I와 J중 더 큰값 or 더 작은 값을 찾아야 하므로 연산이 필요한데 이를 위해 두개를 합쳐야 함.
SWR_gage_data_IJ = pd.concat([SWR_gage_data_I.iloc[:, 0], SWR_gage_data_I.iloc[:, 57], SWR_gage_data_I.iloc[:, 1:57], SWR_gage_data_J.iloc[:, 1:57]], axis=1)
SWR_gage_data_IJ.columns.values[0] = 'gage_name'
SWR_gage_data_IJ.columns.values[1] = 'Height'


Variable_names = ['DE11', 'DE12', 'DE21', 'DE22', 'DE31', 'DE32', 'DE41', 'DE42', 'DE51', 'DE52', 'DE61', 'DE62', 'DE71', 'DE72',
                  'MCE11', 'MCE12', 'MCE21', 'MCE22', 'MCE31', 'MCE32', 'MCE41', 'MCE42', 'MCE51', 'MCE52', 'MCE61', 'MCE62', 'MCE71', 'MCE72']

# SWR_gage_data_IJ 안에는 I것과 J것 각각 두개씩 DE11_max~MCE72_min 이 존재한다. 이 둘 중 max값은 둘 중 더 큰 값을, min값은 둘 중 더 작은 값을 선택하는 부분.
for Variable_name in Variable_names:
    globals()['{}_max'.format(Variable_name)] = SWR_gage_data_IJ.loc[:, '{}_max'.format(Variable_name)].max(axis=1)

for Variable_name in Variable_names:
    globals()['{}_min'.format(Variable_name)] = SWR_gage_data_IJ.loc[:, '{}_min'.format(Variable_name)].max(axis=1)

SWR_gage_data_final = SWR_gage_data_IJ.iloc[:, 0:2]
for Variable_name in Variable_names:
    SWR_gage_data_final = pd.concat([SWR_gage_data_final, globals()['{}_max'.format(Variable_name)]], axis=1)

for Variable_name in Variable_names:
    SWR_gage_data_final = pd.concat([SWR_gage_data_final, globals()['{}_min'.format(Variable_name)]], axis=1)

SWR_gage_data_final.columns = ['gage_name', 'Height', 'DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']


### SWR_avg_data 만들기
DE_max_avg = SWR_gage_data_final.iloc[:, 2:DE_num+2].mean(axis=1) # 2를 더해준 건 앞에 Height와 gage_name이 추가되었기 때문
MCE_max_avg = SWR_gage_data_final.iloc[:, DE_num+2 : DE_num + MCE_num+2].mean(axis=1)
DE_min_avg = SWR_gage_data_final.iloc[:, DE_num+MCE_num+2 : 2*DE_num+MCE_num+2].mean(axis=1)
MCE_min_avg = SWR_gage_data_final.iloc[:, 2*DE_num+MCE_num+2 : 2*DE_num + 2*MCE_num+2].mean(axis=1)
SWR_avg_total = pd.concat([SWR_gage_data_final['Height'], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
SWR_avg_total.columns = ['Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']




### Graph Plotting
# Legend에 쓸 지진파 이름 설정
DE_name = ['DE11', 'DE12', 'DE21', 'DE22', 'DE31', 'DE32', 'DE41', 'DE42', 'DE51', 'DE52', 'DE61', 'DE62', 'DE71', 'DE72']
MCE_name = ['MCE11', 'MCE12', 'MCE21', 'MCE22', 'MCE31', 'MCE32', 'MCE41', 'MCE42', 'MCE51', 'MCE52', 'MCE61', 'MCE62', 'MCE71', 'MCE72']


# 벽체가 바뀌는 지점 찾기
wall_name_change = []
for i in range(len(SWR_gage_data_I)-1):
    if SWR_gage_data_I.iloc[i, 0] != SWR_gage_data_I.iloc[i+1, 0]:
        wall_name_change.append(i)


### Compression AS DE 그래프
for i in range(len(wall_name_change)):
    if i==0:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(min_criteria_DE*1.1, 0)
        plt.grid()
        plt.title('[%s] Compressive Axial Strain(DE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight = 'bold')
        plt.ylabel('Floor',fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0:wall_name_change[i], 2], SWR_gage_data_final.iloc[0: wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[0: wall_name_change[i],j+2+DE_num+MCE_num], SWR_gage_data_final.iloc[0:wall_name_change[i], 1], linewidth = 1, alpha = 0.4)
        plt.axvline(min_criteria_DE, color = 'red', linestyle='--', linewidth = 1.5) # 기준선
        plt.xticks(np.arange(0, min_criteria_DE*1.1, -0.0005))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.2, 0.9), prop={'size': 7.5})
    else:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(min_criteria_DE*1.1, 0)
        plt.grid()
        plt.title('[%s] Compressive Axial Strain(DE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight='bold')
        plt.ylabel('Floor', fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 2], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2+DE_num+MCE_num], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(min_criteria_DE, color='red', linestyle='--', linewidth=1.5)  # 기준선
        plt.xticks(np.arange(0, min_criteria_DE * 1.1, -0.0005))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.2, 0.9),  prop={'size': 7.5})



### Tension AS DE 그래프
for i in range(len(wall_name_change)):
    if i==0:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(0, max_criteria_DE*1.1)
        plt.grid()
        plt.title('[%s] Tensile Axial Strain(DE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight = 'bold')
        plt.ylabel('Floor',fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0:wall_name_change[i], 1], SWR_gage_data_final.iloc[0: wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[0: wall_name_change[i],j+2], SWR_gage_data_final.iloc[0:wall_name_change[i], 1], linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_DE, color = 'red', linestyle='--', linewidth = 1.5) # 기준선
        plt.xticks(np.arange(0, max_criteria_DE*1.1, 0.0005))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.8, 0.9), prop={'size': 7.5})
    else:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(0, max_criteria_DE*1.1)
        plt.grid()
        plt.title('[%s] Tensile Axial Strain(DE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight='bold')
        plt.ylabel('Floor', fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 1], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_DE, color='red', linestyle='--', linewidth=1.5)  # 기준선
        plt.xticks(np.arange(0, max_criteria_DE * 1.1, 0.0005))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.8, 0.9),  prop={'size': 7.5})



### Compression AS MCE 그래프
for i in range(len(wall_name_change)):
    if i==0:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(min_criteria_MCE*1.1, 0)
        plt.grid()
        plt.title('[%s] Compressive Axial Strain(MCE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight = 'bold')
        plt.ylabel('Floor',fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0:wall_name_change[i], 2], SWR_gage_data_final.iloc[0: wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[0: wall_name_change[i],j+2+DE_num+MCE_num+DE_num], SWR_gage_data_final.iloc[0:wall_name_change[i], 1], linewidth = 1, alpha = 0.4)
        plt.axvline(min_criteria_MCE, color = 'red', linestyle='--', linewidth = 1.5) # 기준선
        plt.xticks(np.arange(0, min_criteria_MCE*1.1, -0.00075))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + MCE_name + ['CP'], bbox_to_anchor = (0.2, 0.9), prop={'size': 7.5})
    else:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(min_criteria_MCE*1.1, 0)
        plt.grid()
        plt.title('[%s] Compressive Axial Strain(MCE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight='bold')
        plt.ylabel('Floor', fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 2], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_gage_data_final.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2+DE_num+MCE_num+DE_num], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(min_criteria_MCE, color='red', linestyle='--', linewidth=1.5)  # 기준선
        plt.xticks(np.arange(0, min_criteria_MCE * 1.1, -0.0075))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + MCE_name + ['CP'], bbox_to_anchor = (0.2, 0.9),  prop={'size': 7.5})


### Tension AS MCE 그래프
for i in range(len(wall_name_change)):
    if i==0:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(0, max_criteria_MCE*1.1)
        plt.grid()
        plt.title('[%s] Tensile Axial Strain(MCE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight = 'bold')
        plt.ylabel('Floor',fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0:wall_name_change[i], 1], SWR_gage_data_final.iloc[0: wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(MCE_num):
            plt.plot(SWR_gage_data_final.iloc[0: wall_name_change[i],j+2+DE_num], SWR_gage_data_final.iloc[0:wall_name_change[i], 1], linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_MCE, color = 'red', linestyle='--', linewidth = 1.5) # 기준선
        plt.xticks(np.arange(0, max_criteria_MCE*1.1, 0.00075))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + MCE_name + ['LS'], bbox_to_anchor = (0.8, 0.9), prop={'size': 7.5})
    else:
        plt.figure(dpi=150, figsize=(5, 4.8))
        plt.xlim(0, max_criteria_MCE*1.1)
        plt.grid()
        plt.title('[%s] Tensile Axial Strain(MCE)' % SWR_gage_data_final.iloc[wall_name_change[i], 0], fontweight = 'bold')
        plt.xlabel('Strain(m/m)', fontweight='bold')
        plt.ylabel('Floor', fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 1], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], marker='o', markerfacecolor = 'white', markersize = 3, color='k', linewidth = 3)  # avg plot line
        for j in range(MCE_num):
            plt.plot(SWR_gage_data_final.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2+DE_num], SWR_gage_data_final.iloc[wall_name_change[i-1]+1:wall_name_change[i], 1], linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_MCE, color='red', linestyle='--', linewidth=1.5)  # 기준선
        plt.xticks(np.arange(0, max_criteria_MCE * 1.1, 0.00075))
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.8, 0.9),  prop={'size': 7.5})