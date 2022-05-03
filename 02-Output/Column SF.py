'''
수정해야할것.......
1. 축력은 중력하중 넣었을 때의 실시간 축력으로 바꾸기
2. sheet 수정 필요 ({보:스터럽}, {기둥:후프})

'''


import numpy as np
import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
from io import StringIO  # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import matplotlib.pyplot as plt
# import datatable as dt
import timeit
import re

# 전체 코드 runtime 측정
time_start = timeit.default_timer()

#%% 사용자 입력

# Input 정보 (xlsx)
input_raw_xlsx_dir = r'D:\이형우\내진성능평가\시티오씨엘 6단지(용현학익)'
input_raw_xlsx = 'Input Sheets(609 전이부재 check용).xlsx'

# 해석 output 파일
data_path = r'D:\이형우\내진성능평가\시티오씨엘 6단지(용현학익)\Output\전이부재 확인용 결과파일' # data 폴더 경로
analysis_result = '609D' # 해석 결과 파일 이름(확장자명 제외)

# 전단력 확인할 부재의 종류
element_type_to_check = 'Beam' # 한번에 한가지만 가능!
input_raw_xlsx_sheet = 'Beam' # 확인할 전이 부재 정보가 있는 시트이름

# 전단력 확인할 부재의 기호
element_name_to_check = ['ATG']

# 전단력 확인할 기둥이 있는 층
element_floor_to_check = ['All']

# # story shear y축 층 간격 정하기
# story_yticks = 4 #ex) 3층 간격->3

# graph_xlim = 3 # x축 limit

# # DCR 기준선 (넘는 부재들 error_*_section에서 확인)
# DCR_criteria = 1.0/1.2

#%% Analysis Result 불러오기

to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if (analysis_result in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

# 전단력 Data
SF_info_data = pd.DataFrame()
for i in to_load_list:
    SF_info_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Frame Results - End Forces', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7, 8, 10]) # usecols로 원하는 열만 불러오기
    SF_info_data = SF_info_data.append(SF_info_data_temp)

SF_info_data = SF_info_data[SF_info_data.iloc[:,0].str.contains(element_type_to_check, flags=re.IGNORECASE)]

SF_info_data = SF_info_data.sort_values(by=['Load Case', 'Element Name', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

# 부재 이름 Matching을 위한 Element 정보
element_info_data = pd.DataFrame()
for i in to_load_list:
    element_info_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Element Data - Frame Types', skiprows=[0, 2], header=0, usecols=[0, 2, 5, 7]) # usecols로 원하는 열만 불러오기
    element_info_data = element_info_data.append(element_info_data_temp)

element_info_data = element_info_data[(element_info_data.iloc[:,0].str.contains(element_type_to_check, flags=re.IGNORECASE)) &\
                                      (element_info_data.iloc[:,2].str.contains('|'.join(element_name_to_check)))]

# 층 정보 Matching을 위한 Node 정보
height_info_data = pd.DataFrame()    
for i in to_load_list:
    height_info_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 4]) # usecols로 원하는 열만 불러오기
    height_info_data = height_info_data.append(height_info_data_temp)

element_info_data = pd.merge(element_info_data, height_info_data, how='left', left_on='I-Node ID', right_on='Node ID')

# 'All' 있으면 모든 부재 다 체크할 수 있게!
if 'All' not in element_floor_to_check:
    element_info_data = element_info_data[(element_info_data.iloc[:,2].str.contains('|'.join(element_floor_to_check)))]

element_info_data = element_info_data.drop_duplicates()

# 전단력, 부재 이름 Matching (by Element Name)
SF_ongoing = pd.merge(element_info_data.iloc[:, [1,2,5]], SF_info_data.iloc[:, 1:6], how='left')

SF_ongoing = SF_ongoing.sort_values(by=['Element Name', 'Load Case', 'Step Type'])

# SF_ongoing['Story'] = SF_ongoing['Property Name'].str.split('_').str[1]
# SF_ongoing['Property Name'] = SF_ongoing['Property Name'].str.split('_').str[0]

SF_ongoing.reset_index(inplace=True, drop=True)

#%% Input Sheet 정보 불러오기

# Story Info Data
story_info_xlsx_sheet = 'Story Info'
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=0, usecols=[0, 1, 2], keep_default_na=False)
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
story_info = story_info.sort_values(by='Height(mm)')
story_name = story_info.loc[:, 'Story Name']
story_name.reset_index(inplace=True, drop=True)

#%% 층이름 없는경우, 층이름 Matching

new_name = []
for i in SF_ongoing['Property Name']:
    if '_' in i:
        new_name.append(i.split('_')[0])
    else: new_name.append(i)
    
SF_ongoing['Property Name'] = new_name

SF_ongoing = pd.merge(SF_ongoing, story_info, how='left', left_on='V', right_on='Height(mm)')
SF_ongoing.rename(columns={'Story Name':'Story'}, inplace=True)

SF_ongoing = SF_ongoing.iloc[:,[0,1,3,4,5,6,8]]
SF_ongoing = SF_ongoing.dropna()
SF_ongoing.reset_index(inplace=True, drop=True)

#%%
# 전이부재 Info Data
transfer_element_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=input_raw_xlsx_sheet, skiprows=1, usecols=[0,1,2,3,4,6,7,8,10,22,23])

transfer_element_info.reset_index(inplace=True, drop=True)
transfer_element_info = transfer_element_info.fillna(method='ffill')

#%% 이름 구분 조건 load & 정리

# 구분 조건 load
naming_criteria_xlsx_sheet = 'etc'
naming_criteria = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                                 sheet_name=naming_criteria_xlsx_sheet, skiprows=1, header=0)


# story_section_idx_list = []
# for i in range(transfer_element_info.shape[0]):
    
#     story_section_idx_sublist = []
#     for j, val in enumerate(story_name):
#         if transfer_element_info.iloc[i,1] in val:
#             story_section_idx_sublist.append(j)
        
#         if transfer_element_info.iloc[i,2]  in val:
#             story_section_idx_sublist.append(j)
    
#     story_section_idx_list.append(story_section_idx_sublist)




# 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
naming_criteria_1_index = []
naming_criteria_2_index = []

for i, j in zip(naming_criteria.iloc[:,4].dropna(), naming_criteria.iloc[:,5].dropna()):
    naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
    naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))

### 구분 조건이 층 순서에 상관없이 작동되게 재정렬

# 구분 조건에 해당하는 콘크리트 강도 재정렬
naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,6].dropna()], axis=1)

naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
naming_criteria_property.sort_values('Story(from) Index', inplace=True)
naming_criteria_property.reset_index(inplace=True)

# 구분 조건 재정렬
naming_criteria_1_index.sort()
naming_criteria_2_index.sort()

#%% 시작층, 끝층 정리

naming_from_index = []
naming_to_index = []

for naming_from, naming_to in zip(transfer_element_info['Story(from)'], transfer_element_info['Story(to)']):
    if isinstance(naming_from, str) == False:
        naming_from = str(naming_from)
    if isinstance(naming_to, str) == False:
        naming_from = str(naming_from)
        
    naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
    naming_to_index.append(pd.Index(story_name).get_loc(naming_to))
    
transfer_element_info['Naming from Index'] = naming_from_index
transfer_element_info['Naming to Index'] = naming_to_index

#%%  층 이름을 etc의 이름 구분 조건에 맞게 나누어서 리스트로 정리

naming_from_index_list = []
naming_to_index_list = []
naming_criteria_property_index_list = []

for current_naming_from_index, current_naming_to_index in zip(naming_from_index, naming_to_index):  # 부재의 시작과 끝 층 loop
    naming_from_index_sublist = [current_naming_from_index]
    naming_to_index_sublist = [current_naming_to_index]
    naming_criteria_property_index_sublist = []
        
    for i, j, k in zip(naming_criteria_1_index, naming_criteria_2_index, naming_criteria_property.index):
        if (i >= current_naming_from_index) and (i <= current_naming_to_index):
            naming_from_index_sublist.append(i)
            naming_criteria_property_index_sublist.append(k)
                        
            if (j >= current_naming_from_index) and (j <= current_naming_to_index):
                naming_to_index_sublist.append(j)
            else:
                naming_to_index_sublist.append(i-1)
                
            if i != current_naming_from_index:
                naming_criteria_property_index_sublist.append(k-1)
                                    
        elif (i < current_naming_from_index) and (j >= current_naming_to_index):
            naming_criteria_property_index_sublist.append(k)
            
        elif (i < current_naming_from_index) and (j <= current_naming_to_index):
            naming_to_index_sublist.append(j)
            
        else:
            if max(naming_criteria_1_index) < current_naming_from_index:
                naming_criteria_property_index_sublist.append(max(naming_criteria_property.index))
                
            elif min(naming_criteria_1_index) > current_naming_to_index:
                    naming_criteria_property_index_sublist.append(min(naming_criteria_property.index))
            
        naming_from_index_sublist = list(set(naming_from_index_sublist))
        naming_to_index_sublist = list(set(naming_to_index_sublist))
        naming_criteria_property_index_sublist = list(set(naming_criteria_property_index_sublist))
                
        # sublist 안의 element들을 내림차순으로 정렬            
        naming_from_index_sublist.sort(reverse = True)
        naming_to_index_sublist.sort(reverse = True)
        naming_criteria_property_index_sublist.sort(reverse = True)
    
    # sublist를 합쳐 list로 완성
    naming_from_index_list.append(naming_from_index_sublist)
    naming_to_index_list.append(naming_to_index_sublist)
    naming_criteria_property_index_list.append(naming_criteria_property_index_sublist)        

# 부재명 만들기, 기타 input sheet의 정보들 부재명에 따라 정리
transfer_element_info_contents = transfer_element_info.iloc[:,3:30]  # input sheet에서 나온 properties
transfer_element_info_contents.reset_index(drop=True, inplace=True)  # ?빼도되나?

name_output = []  # new names
floor_output = []
property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
transfer_element_info_contents_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output

for current_column_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_transfer_element_info_contents_index\
    in zip(transfer_element_info['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, transfer_element_info_contents.index):  
    for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
        if p != q:

            name_output.append(current_column_name)
            floor_output.append(str(story_name[p])+'-'+str(story_name[q]))
            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
            transfer_element_info_contents_output.append(transfer_element_info_contents.iloc[current_transfer_element_info_contents_index])
        
        else:
            name_output.append(current_column_name)
            floor_output.append(str(story_name[p]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
            transfer_element_info_contents_output.append(transfer_element_info_contents.iloc[current_transfer_element_info_contents_index])
        
transfer_element_info_contents_output = pd.DataFrame(transfer_element_info_contents_output)
transfer_element_info_contents_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..

transfer_element_info_contents_output['Concrete(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당

# 중간결과
if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 transfer_element_info_contents를 바로 출력
    transfer_element_ongoing = transfer_element_info_contents
else:
    transfer_element_ongoing = pd.concat([pd.Series(name_output, name = 'Name'), pd.Series(floor_output, name = 'Floor'), transfer_element_info_contents_output], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties        


#%% Analysis Result & Input Sheet Info 결합

SF_output = pd.merge(SF_ongoing, transfer_element_ongoing, how='left', left_on=['Property Name', 'Story'], right_on=['Name', 'Floor'])

#%% 공칭 전단강도 구하기

if 'column' in element_type_to_check.lower():    
    # <철근콘크리트 건축구조물의 성능기반 내진설계를 위한 비선형해석모델> 기준 p.65
    
    # 지름 세부 조정
    SF_output['Hoop(DXX)'].replace(10, 9.53, inplace=True)
    SF_output['Hoop(DXX)'].replace(13, 12.7, inplace=True)
    SF_output['Hoop(DXX)'].replace(16, 15.9, inplace=True)
    
    
    # 유효깊이
    D_eff = (SF_output['D'] - 80 - SF_output['Hoop(DXX)'])
    
    # 공칭 전단강도 계산
    Vs = 0.565 * (SF_output['Hoop(DXX)'] / 2) ** 2 * np.pi * D_eff / SF_output['Spacing']
    
    Vc = 1/6 * (1 + SF_output['P I-End'] / (14 * SF_output['B'] * D_eff)) * np.sqrt(SF_output['Concrete(CXX)']) * SF_output['B'] * D_eff / 1000
    
    Vn = Vc + Vs
    
    # 공칭 전단강도 최대값
    Vn_limit = 0.29 * np.sqrt(SF_output['Concrete(CXX)']) * SF_output['B'] * D_eff / 1000

elif 'beam' in element_type_to_check.lower():
    SF_output['Stirrup(DXX)'].replace(10, 9.53, inplace=True)
    SF_output['Stirrup(DXX)'].replace(13, 12.7, inplace=True)
    SF_output['Stirrup(DXX)'].replace(16, 15.9, inplace=True)
    
    
    # 유효깊이
    D_eff = (SF_output['D'] - 80 - SF_output['Stirrup(DXX)'])
    
    # 공칭 전단강도 계산
    # Vs = 0.565 * ((SF_output['Stirrup(DXX)'] / 2) ** 2 * np.pi) * SF_output['N. of Stirrup(int)'] * D_eff / SF_output['s(int)']
    
    # 스터럽 종류 두가지인 경우
    Vs = (0.565 * ((SF_output['Stirrup(DXX)'] / 2) ** 2 * np.pi) * SF_output['N. of Stirrup(int)']) + (0.666 * ((15.9 / 2) ** 2 * np.pi) * 4) * D_eff / SF_output['s(int)']
    
    Vc = 1/6 * np.sqrt(SF_output['Concrete(CXX)']) * SF_output['B'] * D_eff / 1000
    
    Vn = Vc + Vs
    
    # 공칭 전단강도 최대값
    Vn_limit = 0.29 * np.sqrt(SF_output['Concrete(CXX)']) * SF_output['B'] * D_eff / 1000

    
SF_output['Vc(kN)'] = Vc
SF_output['Vs(kN)'] = Vs
SF_output['Vn(kN)'] = Vn
SF_output['Vn_limit(kN)'] = Vn_limit

Vn_Vn_limit_compare = []
DCR_list = []
for i, j, k in zip(Vn, Vn_limit, SF_output['V2 I-End'].abs()):
    if i >= j:
        Vn_Vn_limit_compare.append('YES')
        DCR_list.append(k / j)
    else: 
        Vn_Vn_limit_compare.append('NO')
        DCR_list.append(k / i)

SF_output['Vn > Vn_limit?'] = Vn_Vn_limit_compare
SF_output['DCR'] = DCR_list

# 필요한 정보만 가진 SF_output 데이터프레임
SF_output_sum = SF_output.iloc[:,[0,1,6,2,3,4,5,19,20,21,22,23,24]]

#%% 부재별, 지진파별 DCR 평균

# 지진파 이름 자동 생성
load_name_list = SF_ongoing['Load Case'].drop_duplicates()
load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
load_name_list = [x.split('+')[1].strip() for x in load_name_list]

load_name_list.sort()

# 각 지진파들로 변수 생성 후, 값 대입
DCR_max = SF_output_sum.groupby(['Element Name', 'Property Name', 'Story', 'Load Case'], as_index=False)['DCR'].max()

DCR_max_avg_DE = DCR_max[DCR_max['Load Case'].str.contains('DE')]\
                .groupby(['Element Name', 'Property Name', 'Story'], as_index=False)['DCR'].mean()

DCR_max_avg_MCE = DCR_max[DCR_max['Load Case'].str.contains('MCE')]\
                .groupby(['Element Name', 'Property Name', 'Story'], as_index=False)['DCR'].mean() 

# 내림차순 정리
SF_output_sum_sorted = SF_output_sum.sort_values(by=['DCR'], ascending=False)
DCR_max_avg_DE_sorted = DCR_max_avg_DE.sort_values(by=['DCR'], ascending=False)
DCR_max_avg_MCE_sorted = DCR_max_avg_MCE.sort_values(by=['DCR'], ascending=False)

