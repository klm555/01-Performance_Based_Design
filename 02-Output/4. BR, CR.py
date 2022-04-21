import pandas as pd
import os
import matplotlib.pyplot as plt
import sys

#%% 1. Beam Rotation
#%% 사용자 입력
# 지진파 개수
DE_num = 14
MCE_num = 14

### Initial Setting
# Moment/Shear hinge Gage name
moment_hinge_group_name = 'Beam Rotation(M)'
shear_hinge_group_name = 'Beam Rotation(V)'
column_rotation_H1_group_name = 'Column Rotation(H1)'
column_rotation_H2_group_name = 'Column Rotation(H2)' # 기둥 없으면 그냥 냅두면 됨.

# 허용기준
BR_min_criteria_DE = -0.015/1.2
BR_max_criteria_DE = 0.015/1.2
BR_min_criteria_MCE = -0.03/1.2
BR_max_criteria_MCE = 0.03/1.2

CR_min_criteria_DE = -0.003
CR_max_criteria_DE = 0.003
CR_min_criteria_MCE = -0.004
CR_max_criteria_MCE = 0.004

# 그래프의 x축 limit 값
x_lim = 0.03

# story shear y축 층 간격 정하기
BR_story_yticks = 3 #ex) 3층 간격->3
CR_story_yticks = 1



### 해석 결과 파일 경로 설정
data_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101' # data 폴더 경로
analysis_result = 'Analysis Result' # 해석 결과 파일 이름(확장자명 제외)

# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

# 그림 저장 경로 설정
output_figure_dir = data_path

#%% Analysis Result 불러오기(BR)
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if (analysis_result in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

# Gage data
gage_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Gage Data - Beam Type', skiprows=[0, 2], header=0, usecols=[0, 3, 7, 9]) # usecols로 원하는 열만 불러오기

BR_M_gage_data = gage_data[gage_data['Group Name'] == moment_hinge_group_name]
BR_S_gage_data = gage_data[gage_data['Group Name'] == shear_hinge_group_name]


# Gage result data
result_data = pd.DataFrame()
for i in to_load_list:
    result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Gage Results - Beam Type', skiprows=[0, 2], header=0, usecols=[0, 3, 5, 7, 8, 9])
    result_data = result_data.append(result_data_temp)

result_data = result_data.sort_values(by=['Load Case', 'Element ID', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

result_data = result_data[(result_data['Load Case'].str.contains('DE')) | (result_data['Load Case'].str.contains('MCE'))]

BR_M_result_data = result_data[result_data['Group Name'] == moment_hinge_group_name]
BR_S_result_data = result_data[result_data['Group Name'] == shear_hinge_group_name]

#%% Node, Story 정보 불러오기

# Node Coord data
Node_coord_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])

# Story Info data
story_info_xlsx_sheet = 'Story Data'
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
story_name = story_info.loc[:, 'Story Name']

#%% 데이터 매칭 후 결과뽑기
### Gage data에서 Element ID, I-Node ID 불러와서 v좌표 match하기
BR_M_gage_data = BR_M_gage_data[['Element ID', 'I-Node ID']]; BR_M_gage_num = len(BR_M_gage_data) # gage 개수 얻기
BR_S_gage_data = BR_S_gage_data[['Element ID', 'I-Node ID']]; BR_S_gage_num = len(BR_S_gage_data) # gage 개수 얻기

# I-Node의 v좌표 match해서 추가
gage_data = gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
BR_M_gage_data = BR_M_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
BR_S_gage_data = BR_S_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
BR_S_gage_data.reset_index(drop=True, inplace=True)

### BR_total data 만들기
BR_M_max = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Max') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
BR_M_max = BR_M_max.reshape(BR_M_gage_num, DE_num+MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
BR_M_max = pd.DataFrame(BR_M_max) # array를 다시 dataframe으로
BR_M_min = BR_M_result_data[(BR_M_result_data['Step Type'] == 'Min') & (BR_M_result_data['Performance Level'] == 1)][['Rotation']].values
BR_M_min = BR_M_min.reshape(BR_M_gage_num, DE_num+MCE_num, order='F')
BR_M_min = pd.DataFrame(BR_M_min)
BR_M_total = pd.concat([BR_M_max, BR_M_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### BR_avg_data 만들기
BR_M_DE_max_avg = BR_M_total.iloc[:, 0:DE_num].mean(axis=1)
BR_M_MCE_max_avg = BR_M_total.iloc[:, DE_num : DE_num + MCE_num].mean(axis=1)
BR_M_DE_min_avg = BR_M_total.iloc[:, DE_num+MCE_num : 2*DE_num+MCE_num].mean(axis=1)
BR_M_MCE_min_avg = BR_M_total.iloc[:, 2*DE_num+MCE_num : 2*DE_num + 2*MCE_num].mean(axis=1)
BR_M_avg_total = pd.concat([BR_M_gage_data.iloc[:,[2,3,4]], BR_M_DE_max_avg, BR_M_DE_min_avg, BR_M_MCE_max_avg, BR_M_MCE_min_avg], axis=1)
BR_M_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

### BR_S_total data 만들기
BR_S_max = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Max') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
BR_S_max = BR_S_max.reshape(BR_S_gage_num, DE_num+MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
BR_S_max = pd.DataFrame(BR_S_max) # array를 다시 dataframe으로
BR_S_min = BR_S_result_data[(BR_S_result_data['Step Type'] == 'Min') & (BR_S_result_data['Performance Level'] == 1)][['Rotation']].values
BR_S_min = BR_S_min.reshape(BR_S_gage_num, DE_num+MCE_num, order='F')
BR_S_min = pd.DataFrame(BR_S_min)
BR_S_total = pd.concat([BR_S_max, BR_S_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### BR_S_avg_data 만들기
BR_S_DE_max_avg = BR_S_total.iloc[:, 0:DE_num].mean(axis=1)
BR_S_MCE_max_avg = BR_S_total.iloc[:, DE_num : DE_num + MCE_num].mean(axis=1)
BR_S_DE_min_avg = BR_S_total.iloc[:, DE_num+MCE_num : 2*DE_num+MCE_num].mean(axis=1)
BR_S_MCE_min_avg = BR_S_total.iloc[:, 2*DE_num+MCE_num : 2*DE_num + 2*MCE_num].mean(axis=1)
BR_S_avg_total = pd.concat([BR_S_gage_data.iloc[:,[2,3,4]], BR_S_DE_max_avg, BR_S_DE_min_avg, BR_S_MCE_max_avg, BR_S_MCE_min_avg], axis=1, ignore_index=True)
BR_S_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

#%% ***조작용 코드
# 없애고 싶은 부재의 x좌표 입력
BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -2700)].index)
BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -6.1e-05)].index)
BR_M_avg_total = BR_M_avg_total.drop(BR_M_avg_total[(BR_M_avg_total['X'] == -4725)].index)

#%% BR(Moment Hinge) 그래프

# DE 그래프
plt.figure(1, figsize=(4,5))  # 그래프 사이즈
plt.xlim(-x_lim, x_lim)

plt.scatter(BR_M_avg_total['DE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(BR_M_avg_total['DE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-BR_story_yticks], story_name[::-BR_story_yticks])

# reference line 그려서 허용치 나타내기
plt.axvline(x= BR_min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= BR_max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')

plt.savefig(output_figure_dir + '\\' + 'BR_M_DE')

# MCE 그래프
plt.figure(2, figsize=(4,5))
plt.xlim(-x_lim, x_lim)
plt.scatter(BR_M_avg_total['MCE_min_avg'], BR_M_avg_total['Height'], color = 'k', s=1)
plt.scatter(BR_M_avg_total['MCE_max_avg'], BR_M_avg_total['Height'], color = 'k', s=1)

plt.yticks(story_info['Height(mm)'][::-BR_story_yticks], story_name[::-BR_story_yticks])

plt.axvline(x= BR_min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= BR_max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')

plt.savefig(output_figure_dir + '\\' + 'BR_M_MCE')

#%% BR(Shear Hinge) 그래프
# DE 그래프
plt.figure(3, figsize=(4,5))  # 그래프 사이즈
plt.xlim(-x_lim, x_lim)

plt.scatter(BR_S_avg_total['DE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(BR_S_avg_total['DE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-BR_story_yticks], story_name[::-BR_story_yticks])

# reference line 그려서 허용치 나타내기
plt.axvline(x= BR_min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= BR_max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')

plt.savefig(output_figure_dir + '\\' + 'BR_S_DE')

# MCE 그래프
plt.figure(4, figsize=(4,5))
plt.xlim(-x_lim, x_lim)
plt.scatter(BR_S_avg_total['MCE_min_avg'], BR_S_avg_total['Height'], color = 'k', s=1)
plt.scatter(BR_S_avg_total['MCE_max_avg'], BR_S_avg_total['Height'], color = 'k', s=1)

plt.yticks(story_info['Height(mm)'][::-BR_story_yticks], story_name[::-BR_story_yticks])

plt.axvline(x= BR_min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= BR_max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')

plt.savefig(output_figure_dir + '\\' + 'BR_S_MCE')

plt.show()

#%% 결과 확인할때 필요한 함수들
# 비정상 node 찾는 함수

def ErrorNode(value_min_list, value_max_list, min_criteria, max_criteria):
    error_value_index_list = []
    error_value = []
    for error_value_min, error_value_max, error_value_index in zip(value_min_list, value_max_list, value_min_list.index):
        if (error_value_min <= min_criteria):
            error_value_index_list.append(error_value_index)
            error_value.append(error_value_min)
            
        if (error_value_max >= max_criteria):
            error_value_index_list.append(error_value_index)
            error_value.append(error_value_max)
            
    return error_value_index_list, error_value

# 기준 넘는 node의 좌표 찾는 함수

def ErrorNode_M_coord(error_value_index_list):
    error_x = []
    error_y = []
    error_z = []
    for error_value_index in error_value_index_list:
        error_x_list = BR_M_gage_data['H1']
        error_y_list = BR_M_gage_data['H2']
        error_z_list = BR_M_gage_data['V']

        error_x_list.reset_index(drop=True, inplace=True)
        error_y_list.reset_index(drop=True, inplace=True)
        error_z_list.reset_index(drop=True, inplace=True)

        error_x.append(error_x_list.loc[error_value_index])
        error_y.append(error_y_list.loc[error_value_index])
        error_z.append(error_z_list.loc[error_value_index])

    error_coord = pd.concat([pd.Series(error_x, name='X'), pd.Series(error_y, name='Y'), pd.Series(error_z, name='Z')], axis=1)

    return error_coord

def ErrorNode_S_coord(error_value_index_list):
    error_x = []
    error_y = []
    error_z = []
    for error_value_index in error_value_index_list:
        error_x_list = BR_S_gage_data['H1']
        error_y_list = BR_S_gage_data['H2']
        error_z_list = BR_S_gage_data['V']

        error_x_list.reset_index(drop=True, inplace=True)
        error_y_list.reset_index(drop=True, inplace=True)
        error_z_list.reset_index(drop=True, inplace=True)

        error_x.append(error_x_list.loc[error_value_index])
        error_y.append(error_y_list.loc[error_value_index])
        error_z.append(error_z_list.loc[error_value_index])

    error_coord = pd.concat([pd.Series(error_x, name='X'), pd.Series(error_y, name='Y'), pd.Series(error_z, name='Z')], axis=1)

    return error_coord

def ErrorNode_H1_coord(error_value_index_list):
    error_x = []
    error_y = []
    error_z = []
    for error_value_index in error_value_index_list:
        error_x_list = CR_H1_gage_data['H1']
        error_y_list = CR_H1_gage_data['H2']
        error_z_list = CR_H1_gage_data['V']

        error_x_list.reset_index(drop=True, inplace=True)
        error_y_list.reset_index(drop=True, inplace=True)
        error_z_list.reset_index(drop=True, inplace=True)

        error_x.append(error_x_list.loc[error_value_index])
        error_y.append(error_y_list.loc[error_value_index])
        error_z.append(error_z_list.loc[error_value_index])

    error_coord = pd.concat([pd.Series(error_x, name='X'), pd.Series(error_y, name='Y'), pd.Series(error_z, name='Z')], axis=1)

    return error_coord

def ErrorNode_H2_coord(error_value_index_list):
    error_x = []
    error_y = []
    error_z = []
    for error_value_index in error_value_index_list:
        error_x_list = CR_H2_gage_data['H1']
        error_y_list = CR_H2_gage_data['H2']
        error_z_list = CR_H2_gage_data['V']

        error_x_list.reset_index(drop=True, inplace=True)
        error_y_list.reset_index(drop=True, inplace=True)
        error_z_list.reset_index(drop=True, inplace=True)

        error_x.append(error_x_list.loc[error_value_index])
        error_y.append(error_y_list.loc[error_value_index])
        error_z.append(error_z_list.loc[error_value_index])

    error_coord = pd.concat([pd.Series(error_x, name='X'), pd.Series(error_y, name='Y'), pd.Series(error_z, name='Z')], axis=1)

    return error_coord

#%% 기준 넘는 node의 좌표 출력
error_DE_index_M = ErrorNode(BR_M_avg_total['DE_min_avg'], BR_M_avg_total['DE_max_avg'], BR_min_criteria_DE, BR_max_criteria_DE)
error_DE_index_S = ErrorNode(BR_S_avg_total['DE_min_avg'], BR_S_avg_total['DE_max_avg'], BR_min_criteria_DE, BR_max_criteria_DE)
error_MCE_index_M = ErrorNode(BR_M_avg_total['MCE_min_avg'], BR_M_avg_total['MCE_max_avg'], BR_min_criteria_MCE, BR_max_criteria_MCE)
error_MCE_index_S = ErrorNode(BR_S_avg_total['MCE_min_avg'], BR_S_avg_total['MCE_max_avg'], BR_min_criteria_MCE, BR_max_criteria_MCE)

error_DE_coord_M = ErrorNode_M_coord(error_DE_index_M[0])
error_DE_coord_S = ErrorNode_S_coord(error_DE_index_S[0])
error_MCE_coord_M = ErrorNode_M_coord(error_MCE_index_M[0])
error_MCE_coord_S = ErrorNode_S_coord(error_MCE_index_S[0])

error_DE_coord_M = pd.merge(error_DE_coord_M, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_MCE_coord_M = pd.merge(error_MCE_coord_M, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_DE_coord_S = pd.merge(error_DE_coord_S, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_MCE_coord_S = pd.merge(error_MCE_coord_S, story_info, how='left', left_on='Z', right_on='Height(mm)')

error_DE_coord_M = pd.concat((error_DE_coord_M.iloc[:,[0,1,2,4]], pd.Series(error_DE_index_M[1], name='Value')), axis=1)
error_MCE_coord_M = pd.concat((error_MCE_coord_M.iloc[:,[0,1,2,4]], pd.Series(error_MCE_index_M[1], name='Value')), axis=1)
error_DE_coord_S = pd.concat((error_DE_coord_S.iloc[:,[0,1,2,4]], pd.Series(error_DE_index_S[1], name='Value')), axis=1)
error_MCE_coord_S = pd.concat((error_MCE_coord_S.iloc[:,[0,1,2,4]], pd.Series(error_MCE_index_S[1], name='Value')), axis=1)

#%% 2. Column Rotation
#%% Analysis Result 불러오기(CR)

# Gage data
CR_H1_gage_data = gage_data[gage_data['Group Name'] == column_rotation_H1_group_name]
CR_H2_gage_data = gage_data[gage_data['Group Name'] == column_rotation_H2_group_name]

# Gage result data
CR_H1_result_data = result_data[result_data['Group Name'] == column_rotation_H1_group_name]
CR_H2_result_data = result_data[result_data['Group Name'] == column_rotation_H2_group_name]

# 기둥 없으면 0행렬 나옴 -> 진행안함
if CR_H1_gage_data.shape[0] == 0:
    sys.exit('NO COLUMN FOUND')

#%% 데이터 매칭 후 결과뽑기
### Gage data에서 Element ID, I-Node ID 불러와서 v좌표 match하기
CR_H1_gage_data = CR_H1_gage_data[['Element ID', 'I-Node ID']]; CR_H1_gage_num = len(CR_H1_gage_data) # gage 개수 얻기
CR_H2_gage_data = CR_H2_gage_data[['Element ID', 'I-Node ID']]; CR_H2_gage_num = len(CR_H2_gage_data) # gage 개수 얻기


# I-Node의 v좌표 match해서 추가
CR_H1_gage_data = CR_H1_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
CR_H2_gage_data = CR_H2_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')
CR_H1_gage_data.reset_index(drop=True, inplace=True)
CR_H2_gage_data.reset_index(drop=True, inplace=True)

### CR_total data 만들기
CR_H1_max = CR_H1_result_data[(CR_H1_result_data['Step Type'] == 'Max') & (CR_H1_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
CR_H1_max = CR_H1_max.reshape(CR_H1_gage_num, DE_num+MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
CR_H1_max = pd.DataFrame(CR_H1_max) # array를 다시 dataframe으로
CR_H1_min = CR_H1_result_data[(CR_H1_result_data['Step Type'] == 'Min') & (CR_H1_result_data['Performance Level'] == 1)][['Rotation']].values
CR_H1_min = CR_H1_min.reshape(CR_H1_gage_num, DE_num+MCE_num, order='F')
CR_H1_min = pd.DataFrame(CR_H1_min)
CR_H1_total = pd.concat([CR_H1_max, CR_H1_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### BR_avg_data 만들기
CR_H1_DE_max_avg = CR_H1_total.iloc[:, 0:DE_num].mean(axis=1)
CR_H1_MCE_max_avg = CR_H1_total.iloc[:, DE_num : DE_num + MCE_num].mean(axis=1)
CR_H1_DE_min_avg = CR_H1_total.iloc[:, DE_num+MCE_num : 2*DE_num+MCE_num].mean(axis=1)
CR_H1_MCE_min_avg = CR_H1_total.iloc[:, 2*DE_num+MCE_num : 2*DE_num + 2*MCE_num].mean(axis=1)
CR_H1_avg_total = pd.concat([CR_H1_gage_data.iloc[:,[2,3,4]], CR_H1_DE_max_avg, CR_H1_DE_min_avg, CR_H1_MCE_max_avg, CR_H1_MCE_min_avg], axis=1)
CR_H1_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

### CR_H2_total data 만들기
CR_H2_max = CR_H2_result_data[(CR_H2_result_data['Step Type'] == 'Max') & (CR_H2_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
CR_H2_max = CR_H2_max.reshape(CR_H2_gage_num, DE_num+MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
CR_H2_max = pd.DataFrame(CR_H2_max) # array를 다시 dataframe으로
CR_H2_min = CR_H2_result_data[(CR_H2_result_data['Step Type'] == 'Min') & (CR_H2_result_data['Performance Level'] == 1)][['Rotation']].values
CR_H2_min = CR_H2_min.reshape(CR_H2_gage_num, DE_num+MCE_num, order='F')
CR_H2_min = pd.DataFrame(CR_H2_min)
CR_H2_total = pd.concat([CR_H2_max, CR_H2_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### CR_H2_avg_data 만들기
CR_H2_DE_max_avg = CR_H2_total.iloc[:, 0:DE_num].mean(axis=1)
CR_H2_MCE_max_avg = CR_H2_total.iloc[:, DE_num : DE_num + MCE_num].mean(axis=1)
CR_H2_DE_min_avg = CR_H2_total.iloc[:, DE_num+MCE_num : 2*DE_num+MCE_num].mean(axis=1)
CR_H2_MCE_min_avg = CR_H2_total.iloc[:, 2*DE_num+MCE_num : 2*DE_num + 2*MCE_num].mean(axis=1)
CR_H2_avg_total = pd.concat([CR_H2_gage_data.iloc[:,[2,3,4]], CR_H2_DE_max_avg, CR_H2_DE_min_avg, CR_H2_MCE_max_avg, CR_H2_MCE_min_avg], axis=1, ignore_index=True)
CR_H2_avg_total.columns = ['X', 'Y', 'Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

#%% ***조작용 코드
# 없애고 싶은 부재의 x좌표 입력
CR_H1_avg_total = CR_H1_avg_total.drop(CR_H1_avg_total[(CR_H1_avg_total['X'] == -2700)].index)
CR_H1_avg_total = CR_H1_avg_total.drop(CR_H1_avg_total[(CR_H1_avg_total['X'] == -6.1e-05)].index)
CR_H2_avg_total = CR_H1_avg_total.drop(CR_H1_avg_total[(CR_H1_avg_total['X'] == -4725)].index)

#%% CR(H1) 그래프

# DE 그래프
plt.figure(5, figsize=(4,5))  # 그래프 사이즈
plt.xlim(-0.005, 0.005)

plt.scatter(CR_H1_avg_total['DE_min_avg'], CR_H1_avg_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(CR_H1_avg_total['DE_max_avg'], CR_H1_avg_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-CR_story_yticks], story_name[::-CR_story_yticks])
plt.ylim(CR_H1_avg_total['Height'].min()-5000, CR_H1_avg_total['Height'].max()+5000)

# reference line 그려서 허용치 나타내기
plt.axvline(x= CR_min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= CR_max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')
plt.savefig(output_figure_dir + '\\' + 'CR_H1_DE')

# MCE 그래프
plt.figure(6, figsize=(4,5))
plt.xlim(-0.005, 0.005)

plt.scatter(CR_H1_avg_total['MCE_min_avg'], CR_H1_avg_total['Height'], color = 'k', s=1)
plt.scatter(CR_H1_avg_total['MCE_max_avg'], CR_H1_avg_total['Height'], color = 'k', s=1)

plt.yticks(story_info['Height(mm)'][::-CR_story_yticks], story_name[::-CR_story_yticks])
plt.ylim(CR_H1_avg_total['Height'].min()-5000, CR_H1_avg_total['Height'].max()+5000)

plt.axvline(x= CR_min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= CR_max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')
plt.savefig(output_figure_dir + '\\' + 'CR_H1_MCE')

#%% CR(H2) 그래프
# DE 그래프
plt.figure(7, figsize=(4,5))  # 그래프 사이즈
plt.xlim(-0.005, 0.005)

plt.scatter(CR_H2_avg_total['DE_min_avg'], CR_H2_avg_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(CR_H2_avg_total['DE_max_avg'], CR_H2_avg_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-CR_story_yticks], story_name[::-CR_story_yticks])
plt.ylim(CR_H2_avg_total['Height'].min()-5000, CR_H2_avg_total['Height'].max()+5000)

# reference line 그려서 허용치 나타내기
plt.axvline(x= CR_min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= CR_max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')

# plt.title('Rotation')
plt.savefig(output_figure_dir + '\\' + 'CR_H2_DE')

# MCE 그래프
plt.figure(8, figsize=(4,5))
plt.xlim(-0.005, 0.005)
plt.scatter(CR_H2_avg_total['MCE_min_avg'], CR_H2_avg_total['Height'], color = 'k', s=1)
plt.scatter(CR_H2_avg_total['MCE_max_avg'], CR_H2_avg_total['Height'], color = 'k', s=1)

plt.yticks(story_info['Height(mm)'][::-CR_story_yticks], story_name[::-CR_story_yticks])
plt.ylim(CR_H2_avg_total['Height'].min()-5000, CR_H2_avg_total['Height'].max()+5000)

plt.axvline(x= CR_min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= CR_max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Rotation(rad)')
plt.ylabel('Story')
# plt.title('Rotation')
plt.savefig(output_figure_dir + '\\' + 'CR_H2_MCE')

plt.show()

#%% 기준 넘는 node의 좌표 출력
error_DE_index_H1 = ErrorNode(CR_H1_avg_total['DE_min_avg'], CR_H1_avg_total['DE_max_avg'], CR_min_criteria_DE, CR_max_criteria_DE)
error_DE_index_H2 = ErrorNode(CR_H2_avg_total['DE_min_avg'], CR_H2_avg_total['DE_max_avg'], CR_min_criteria_DE, CR_max_criteria_DE)
error_MCE_index_H1 = ErrorNode(CR_H1_avg_total['MCE_min_avg'], CR_H1_avg_total['MCE_max_avg'], CR_min_criteria_MCE, CR_max_criteria_MCE)
error_MCE_index_H2 = ErrorNode(CR_H2_avg_total['MCE_min_avg'], CR_H2_avg_total['MCE_max_avg'], CR_min_criteria_MCE, CR_max_criteria_MCE)

error_DE_coord_H1 = ErrorNode_H1_coord(error_DE_index_H1[0])
error_DE_coord_H2 = ErrorNode_H2_coord(error_DE_index_H2[0])
error_MCE_coord_H1 = ErrorNode_H1_coord(error_MCE_index_H1[0])
error_MCE_coord_H2 = ErrorNode_H2_coord(error_MCE_index_H2[0])

error_DE_coord_H1 = pd.merge(error_DE_coord_H1, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_MCE_coord_H1 = pd.merge(error_MCE_coord_H1, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_DE_coord_H2 = pd.merge(error_DE_coord_H2, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_MCE_coord_H2 = pd.merge(error_MCE_coord_H2, story_info, how='left', left_on='Z', right_on='Height(mm)')

error_DE_coord_H1 = pd.concat((error_DE_coord_H1.iloc[:,[0,1,2,4]], pd.Series(error_DE_index_H1[1], name='Value')), axis=1)
error_MCE_coord_H1 = pd.concat((error_MCE_coord_H1.iloc[:,[0,1,2,4]], pd.Series(error_MCE_index_H1[1], name='Value')), axis=1)
error_DE_coord_H2 = pd.concat((error_DE_coord_H2.iloc[:,[0,1,2,4]], pd.Series(error_DE_index_H2[1], name='Value')), axis=1)
error_MCE_coord_H2 = pd.concat((error_MCE_coord_H2.iloc[:,[0,1,2,4]], pd.Series(error_MCE_index_H2[1], name='Value')), axis=1)