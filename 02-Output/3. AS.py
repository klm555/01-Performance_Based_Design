import pandas as pd
import os
import matplotlib.pyplot as plt

#%% 사용자 입력
# 지진파 개수
DE_num = 14
MCE_num = 14

# 허용기준
min_criteria_DE = -0.002
max_criteria_DE = 0.04
min_criteria_MCE = -0.002
max_criteria_MCE = 0.04

# story shear y축 층 간격 정하기
story_yticks = 3 #ex) 3층 간격->3

### 해석 결과 파일 경로 설정
data_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101' # data 폴더 경로
analysis_result = 'Analysis Result' # 해석 결과 파일 이름(확장자명 제외)

# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

# 그림 저장 경로 설정
output_figure_dir = data_path

#%% Analysis Result 불러오기
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if (analysis_result in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

# Gage data
AS_gage_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Gage Data - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 3, 7, 9])

# Gage result data
AS_result_data = pd.DataFrame()
for i in to_load_list:
    AS_result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Gage Results - Bar Type', skiprows=[0, 2], header=0, usecols=[0, 3, 5, 7, 8, 9])
    AS_result_data = AS_result_data.append(AS_result_data_temp)

AS_result_data = AS_result_data.sort_values(by= ['Load Case', 'Element ID', 'Step Type']) # 여러개로 나눠돌릴 경우 순서가 섞여있을 수 있어 DE11~MCE72 순으로 정렬

AS_result_data = AS_result_data[(AS_result_data['Load Case'].str.contains('DE')) | (AS_result_data['Load Case'].str.contains('MCE'))]

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
AS_gage_data = AS_gage_data[['Element ID', 'I-Node ID']]; gage_num = len(AS_gage_data) # gage 개수 얻기

# I-Node의 v좌표 match해서 추가
AS_gage_data = AS_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')

### AS_total data 만들기
AS_max = AS_result_data[(AS_result_data['Step Type'] == 'Max') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values # dataframe을 array로
AS_max = AS_max.reshape(gage_num, DE_num + MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
AS_max = pd.DataFrame(AS_max) # array를 다시 dataframe으로
AS_min = AS_result_data[(AS_result_data['Step Type'] == 'Min') & (AS_result_data['Performance Level'] == 1)][['Axial Strain']].values
AS_min = AS_min.reshape(gage_num, DE_num + MCE_num, order='F')
AS_min = pd.DataFrame(AS_min)
AS_total = pd.concat([AS_max, AS_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### AS_avg_data 만들기
DE_max_avg = AS_total.iloc[:, 0:DE_num].mean(axis=1)
MCE_max_avg = AS_total.iloc[:, DE_num : DE_num + MCE_num].mean(axis=1)
DE_min_avg = AS_total.iloc[:, DE_num+MCE_num : 2*DE_num+MCE_num].mean(axis=1)
MCE_min_avg = AS_total.iloc[:, 2*DE_num+MCE_num : 2*DE_num + 2*MCE_num].mean(axis=1)
AS_avg_total = pd.concat([AS_gage_data['V'], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
AS_avg_total.columns = ['Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

### DE11_max ~ MCE72_min 생성
# 자동 변수 생성
Variable_names = ['DE11', 'DE12', 'DE21', 'DE22', 'DE31', 'DE32', 'DE41', 'DE42', 'DE51', 'DE52', 'DE61', 'DE62', 'DE71', 'DE72',
                  'MCE11', 'MCE12', 'MCE21', 'MCE22', 'MCE31', 'MCE32', 'MCE41', 'MCE42', 'MCE51', 'MCE52', 'MCE61', 'MCE62', 'MCE71', 'MCE72']

for Variable_name in Variable_names:
    globals()['{}_max'.format(Variable_name)] = AS_result_data[(AS_result_data['Load Case'] == '[1] + {}'.format(Variable_name)) &
                                                               (AS_result_data['Step Type'] == 'Max') &
                                                               (AS_result_data['Performance Level'] == 1)][['Axial Strain']]

    globals()['{}_min'.format(Variable_name)] = AS_result_data[(AS_result_data['Load Case'] == '[1] + {}'.format(Variable_name)) &
                                                               (AS_result_data['Step Type'] == 'Min') &
                                                               (AS_result_data['Performance Level'] == 1)][['Axial Strain']]
    globals()['{}_max'.format(Variable_name)].reset_index(drop=True, inplace=True)
    globals()['{}_min'.format(Variable_name)].reset_index(drop=True, inplace=True)

# 그냥 AS_total에서 불러다가 써도 됨

#%% ***조작용 코드
# 데이터 없애기 위한 기준값 입력
AS_avg_total = AS_avg_total.drop(AS_avg_total[(AS_avg_total.loc[:,'DE_min_avg'] < -0.002)].index)
AS_avg_total = AS_avg_total.drop(AS_avg_total[(AS_avg_total.loc[:,'MCE_min_avg'] < -0.002)].index)
# .....위와 같은 포맷으로 계속

#%% 그래프

# DE 그래프
# AS_DE_1
plt.figure(1, figsize=(5,4))  # 그래프 사이즈
plt.xlim(-0.003, 0)

plt.scatter(AS_avg_total['DE_min_avg'], AS_avg_total['Height'], color = 'r', s=5) # s=1 : point size
plt.scatter(AS_avg_total['DE_max_avg'], AS_avg_total['Height'], color = 'k', s=5)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

# reference line 그려서 허용치 나타내기
plt.axvline(x= min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Axial Strain(m/m)')
plt.ylabel('Story')

# 그래프 저장
plt.savefig(output_figure_dir + '\\' + 'AS_DE_1')

# AS_DE_2
plt.figure(2, figsize=(5,4))  # 그래프 사이즈
plt.xlim(0, 0.013)
plt.scatter(AS_avg_total['DE_min_avg'], AS_avg_total['Height'], color = 'r', s=5) # s=1 : point size
plt.scatter(AS_avg_total['DE_max_avg'], AS_avg_total['Height'], color = 'k', s=5)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= min_criteria_DE, color='r', linestyle='--')
plt.axvline(x= max_criteria_DE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Axial Strain(m/m)')
plt.ylabel('Story')
# plt.title('Axial Strain')

plt.savefig(output_figure_dir + '\\' + 'AS_DE_2')


# MCE 그래프
# AS_MCE_1
plt.figure(3, figsize=(5,4))
plt.xlim(-0.003, 0)
plt.scatter(AS_avg_total['MCE_min_avg'], AS_avg_total['Height'], color = 'r', s=5)
plt.scatter(AS_avg_total['MCE_max_avg'], AS_avg_total['Height'], color = 'k', s=5)

plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Axial Strain(m/m)')
plt.ylabel('Story')
# plt.title('Axial Strain')

# 그래프 저장
plt.savefig(output_figure_dir + '\\' + 'AS_MCE_1')

# AS_MCE_2
plt.figure(4, figsize=(5,4))
plt.xlim(0, 0.013)
plt.scatter(AS_avg_total['MCE_min_avg'], AS_avg_total['Height'], color = 'r', s=5)
plt.scatter(AS_avg_total['MCE_max_avg'], AS_avg_total['Height'], color = 'k', s=5)

plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= min_criteria_MCE, color='r', linestyle='--')
plt.axvline(x= max_criteria_MCE, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('Axial Strain(m/m)')
plt.ylabel('Story')
# plt.title('Axial Strain')

# 그래프 저장
plt.savefig(output_figure_dir + '\\' + 'AS_MCE_2')

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

def ErrorNode_coord(error_value_index_list):
    error_x = []
    error_y = []
    error_z = []
    for error_value_index in error_value_index_list:
        error_x_list = AS_gage_data['H1']
        error_y_list = AS_gage_data['H2']
        error_z_list = AS_gage_data['V']

        error_x_list.reset_index(drop=True, inplace=True)
        error_y_list.reset_index(drop=True, inplace=True)
        error_z_list.reset_index(drop=True, inplace=True)

        error_x.append(error_x_list.loc[error_value_index])
        error_y.append(error_y_list.loc[error_value_index])
        error_z.append(error_z_list.loc[error_value_index])

    error_coord = pd.concat([pd.Series(error_x, name='X'), pd.Series(error_y, name='Y'), pd.Series(error_z, name='Z')], axis=1)

    return error_coord

# %% 기준 넘는 node의 좌표 출력
error_DE_index = ErrorNode(AS_avg_total['DE_min_avg'], AS_avg_total['DE_max_avg'], min_criteria_DE, max_criteria_DE)
error_MCE_index = ErrorNode(AS_avg_total['MCE_min_avg'], AS_avg_total['MCE_max_avg'], min_criteria_MCE, max_criteria_MCE)

error_DE_coord = ErrorNode_coord(error_DE_index[0])
error_MCE_coord = ErrorNode_coord(error_MCE_index[0])

error_DE_coord = pd.merge(error_DE_coord, story_info, how='left', left_on='Z', right_on='Height(mm)')
error_MCE_coord = pd.merge(error_MCE_coord, story_info, how='left', left_on='Z', right_on='Height(mm)')

error_DE_coord = pd.concat((error_DE_coord.iloc[:,[0,1,2,4]], pd.Series(error_DE_index[1], name='Value')), axis=1)
error_MCE_coord = pd.concat((error_MCE_coord.iloc[:,[0,1,2,4]], pd.Series(error_MCE_index[1], name='Value')), axis=1)