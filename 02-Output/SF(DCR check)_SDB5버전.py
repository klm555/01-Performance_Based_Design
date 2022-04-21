"""
###  update(2022.01.24)  ###
1. [H1 DE], [H2 DE], [H1 MCE], [H2 MCE] 순
2. N.G 섹션(error_*_section) 확인 시, DCR 값도 함께 확인 가능.
"""

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

#%% 파일 경로 지정

# Wall input 정보 (xlsx)
wall_raw_xlsx_dir = r'D:\이형우\내진성능평가\송도 B5\B동'
wall_raw_xlsx = 'Input Sheets(B)_plate modified_조작.xlsx'

wall_raw_xlsx_sheet = 'Wall'
story_info_xlsx_sheet = 'Story Info'
naming_criteria_xlsx_sheet = 'etc'
SF_ongoing_modified_sheet = 'SF_ongoing_modified'

# 해석 output 파일 (txt)
H1_txt_dir = r'D:\이형우\내진성능평가\송도 B5\비선형해석 결과자료\B동(18)\H1'
H2_txt_dir = r'D:\이형우\내진성능평가\송도 B5\비선형해석 결과자료\B동(18)\H2'
V_txt_dir = r'D:\이형우\내진성능평가\송도 B5\비선형해석 결과자료\B동(18)\V'

DE11_txt = 'DE11.txt'
DE12_txt = 'DE12.txt'
DE21_txt = 'DE21.txt'
DE22_txt = 'DE22.txt'
DE31_txt = 'DE31.txt'
DE32_txt = 'DE32.txt'
DE41_txt = 'DE41.txt'
DE42_txt = 'DE42.txt'
DE51_txt = 'DE51.txt'
DE52_txt = 'DE52.txt'
DE61_txt = 'DE61.txt'
DE62_txt = 'DE62.txt'
DE71_txt = 'DE71.txt'
DE72_txt = 'DE72.txt'
MCE11_txt = 'MCE11.txt'
MCE12_txt = 'MCE12.txt'
MCE21_txt = 'MCE21.txt'
MCE22_txt = 'MCE22.txt'
MCE31_txt = 'MCE31.txt'
MCE32_txt = 'MCE32.txt'
MCE41_txt = 'MCE41.txt'
MCE42_txt = 'MCE42.txt'
MCE51_txt = 'MCE51.txt'
MCE52_txt = 'MCE52.txt'
MCE61_txt = 'MCE61.txt'
MCE62_txt = 'MCE62.txt'
MCE71_txt = 'MCE71.txt'
MCE72_txt = 'MCE72.txt'

# story shear y축 층 간격 정하기
story_yticks = 4 #ex) 3층 간격->3

graph_xlim = 3 # x축 limit

# DCR 기준선 (넘는 부재들 error_*_section에서 확인)
DCR_criteria = 1.0/1.2+0.05

#%% 데이터베이스

# 철근, 콘크리트 정보
steel_geometry_database = pd.DataFrame({'Name': [10, 13, 16, 19, 22,
                                                 25, 29],
                                        'Diameter(mm)': [9.53, 12.7, 15.9,
                                                         19.1, 22.2, 25.4,
                                                         28.6],
                                        'Area(mm^2)': [71.33, 126.7, 198.6,
                                                       286.52, 387.08, 506.71,
                                                       642.42]})

concrete_database = pd.DataFrame({'Name': ['C15', 'C21', 'C24', 'C27', 'C30', 'C35', 'C40', 'C45', 'C49'],
                                  'Strength(MPa)': [15*1.2, 21*1.2, 24*1.1, 27*1.1, 30*1.1, 35*1.1, 40*1.1, 45, 49],
                                  'Elastic Modulus(kN/mm^2)': [18.369, 24.920,
                                                               25.310, 26.323,
                                                               27.264, 28.702,
                                                               30.008, 31.185, 32.083],
                                  'Poisson\'s Ratio': [0.25]*9,\
                                  'Weight Density(kN/mm^3)': [2.354*10**(-8)]*9})

steel_database = pd.DataFrame({'Name': ['SD240', 'SD400', 'SD500', 'SD600'],
                               'Strength(MPa)': [240, 400*1.17, 500*1.13, 600*1.11],
                               'Elastic Modulus(kN/mm^2)': [200, 200, 200, 200]})

#%% 파일 load

# Wall 정보 load
wall = pd.read_excel(wall_raw_xlsx_dir+'\\'+wall_raw_xlsx,
                     sheet_name=wall_raw_xlsx_sheet, skiprows=1, header=0)

wall = wall.dropna(axis=0, how='all')
wall.reset_index(inplace=True, drop=True)
wall = wall.fillna(method='ffill')

# Story 정보 load
story_info = pd.read_excel(wall_raw_xlsx_dir+'\\'+wall_raw_xlsx,
                           sheet_name=story_info_xlsx_sheet, keep_default_na=False)

# Story 정보에서 층이름만 뽑아내기
story_name = story_info.loc[:, 'Floor Name']
story_name = story_name[::-1]  # 층 이름 재배열
story_name.reset_index(drop=True, inplace=True)

#%% 글자 나누는 함수 (12F~15F, D10@300)

def str_div(temp_list):
    
    first = []
    second = []
    
    for i in temp_list:
        if '~' in i:
            first.append(i.split('~')[0])
            second.append(i.split('~')[1])
        elif '@' in i:
            first.append(i.split('@')[0])
            second.append(i.split('@')[1])
        elif '-' in i:
            second.append(i.split('-')[0])
            first.append(i.split('-')[1])
        else:
            first.append(i)
            second.append(i)
    
    first = pd.Series(first).str.strip()
    second = pd.Series(second).str.strip()
    
    return first, second

#%% 철근 지름 앞의 D 떼주는 함수 (D10...)

def str_extract(sth_str):
    result = int(re.findall(r'[0-9]+', sth_str)[0])
    
    return result

#%% 철근 강도 구하는 함수

def rebar_det(rebar_diameter, rebar_det_criteria):       
    
    for i in rebar_det_criteria.itertuples():
        if i[2] == 'under':
            if rebar_diameter <= int(re.findall(r'[0-9]+', i[1])[0]):
                a = i[3]
        else:
            if rebar_diameter >= int(re.findall(r'[0-9]+', i[1])[0]):
                a = i[3]
                continue
    return a

#%% 불러온 wall 정보 정리하기

# 글자가 합쳐져 있을 경우 글자 나누기 (12F~15F, D10@300)
# 층 나누기

if wall['Story(to)'].isnull().any() == True:
    wall['Story(to)'] = str_div(wall['Story(from)'])[1]
    wall['Story(from)'] = str_div(wall['Story(from)'])[0]
else: pass

# V. Rebar 나누기
if wall['V. Rebar Space'].isnull().any() == True:
    wall['V. Rebar Space'] = str_div(wall['Vertical Rebar(DXX)'])[1].astype(int)
    wall['Vertical Rebar(DXX)'] = str_div(wall['Vertical Rebar(DXX)'])[0]
else: pass

# H. Rebar 나누기
if wall['H. Rebar Space'].isnull().any() == True:
    wall['H. Rebar Space'] = str_div(wall['Horizontal Rebar(DXX)'])[1].astype(int)
    wall['Horizontal Rebar(DXX)'] = str_div(wall['Horizontal Rebar(DXX)'])[0]
else: pass

# 철근의 앞에붙은 D 떼어주기
new_v_rebar = []
new_h_rebar = []

for i in wall['Vertical Rebar(DXX)']:
    if isinstance(i, int):
        new_v_rebar.append(i)
    else:
        new_v_rebar.append(str_extract(i))
        
for j in wall['Horizontal Rebar(DXX)']:
    if isinstance(j, int):
        new_h_rebar.append(j)
    else:
        new_h_rebar.append(str_extract(j))
        
wall['Vertical Rebar(DXX)'] = new_v_rebar
wall['Horizontal Rebar(DXX)'] = new_h_rebar

# Rebar 강도 정하기
# 구분 조건 load
naming_criteria = pd.read_excel(wall_raw_xlsx_dir+'\\'+wall_raw_xlsx,
                                 sheet_name=naming_criteria_xlsx_sheet, skiprows=1, header=0)

rebar_criteria = naming_criteria.iloc[:,[8,9,10]]
rebar_criteria = rebar_criteria.dropna(axis=0, how='all')

# 철근 강도 구하는 함수로 철근 강도 구하기
v_rebar_strength = []
h_rebar_strength = []

for i, j in zip(wall['Vertical Rebar(DXX)'], wall['Horizontal Rebar(DXX)']):
    v_rebar_strength.append(rebar_det(i, rebar_criteria))
    h_rebar_strength.append(rebar_det(j, rebar_criteria))
    
wall['V. Rebar Strength(SDXXX)'] = v_rebar_strength
wall['H. Rebar Strength(SDXXX)'] = h_rebar_strength

#%% **부록**에 부재 정보 입력을 위한 열 추가

h_rebar_info_appendix = []
        
for i in wall.itertuples():
    if i[10] <= 30:
        h_rebar_info_appendix.append(str(i[10]) + '-' + 'D' + str(i[9]))
    else:
        h_rebar_info_appendix.append('D' + str(i[9]) + '@' + str(i[10]))    
# i[-3]: cover thickness // i[-2]: length
        
wall['H. Rebar Info'] = h_rebar_info_appendix

#%% Rebar Spacing 다시 정리하기 (12-D10 등...)

new_v_rebar_spacing = []
new_h_rebar_spacing = []

for i in wall.itertuples():
    if i[7] <= 30:
        new_v_rebar_spacing.append(((i[13]-(2*i[12]+steel_geometry_database.loc[steel_geometry_database['Name'] == i[6], 'Diameter(mm)']))/(i[7]/2-1)).iloc[0])
    else:
        new_v_rebar_spacing.append(i[7])
        
for i in wall.itertuples():
    if i[10] <= 30:
        new_h_rebar_spacing.append(((i[13]-(2*i[12]+steel_geometry_database.loc[steel_geometry_database['Name'] == i[9], 'Diameter(mm)']))/(i[10]/2-1)).iloc[0])
    else:
        new_h_rebar_spacing.append(i[10])     
# i[-3]: cover thickness // i[-2]: length
        
wall['V. Rebar Space'] = new_v_rebar_spacing
wall['H. Rebar Space'] = new_h_rebar_spacing

# new = wall['V. Rebar Space'].apply(lambda x: (wall['Length']-(wall['Cover Thickness']*2+wall['Vertical Rebar(DXX)'])))

#%% 이름 구분 조건 load & 정리

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

for naming_from, naming_to in zip(wall['Story(from)'], wall['Story(to)']):
    if isinstance(naming_from, str) == False:
        naming_from = str(naming_from)
    if isinstance(naming_to, str) == False:
        naming_from = str(naming_from)
        
    naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
    naming_to_index.append(pd.Index(story_name).get_loc(naming_to))
    
wall['Naming from Index'] = naming_from_index
wall['Naming to Index'] = naming_to_index

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
wall_info = wall.iloc[:,3:30]  # input sheet에서 나온 properties
wall_info.reset_index(drop=True, inplace=True)  # ?빼도되나?

name_output = []  # new names
floor_output = []
property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
wall_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output

for current_wall_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_wall_info_index\
    in zip(wall['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, wall_info.index):  
    for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
        if p != q:
            for s in reversed(range(p, q+1)):
                name_output.append(current_wall_name)
                floor_output.append(str(story_name[s]))
                property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                wall_info_output.append(wall_info.iloc[current_wall_info_index])
        
        else:
            name_output.append(current_wall_name)
            floor_output.append(str(story_name[p]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
            wall_info_output.append(wall_info.iloc[current_wall_info_index])
        
wall_info_output = pd.DataFrame(wall_info_output)
wall_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..

wall_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당

# 중간결과
if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 wall_info를 바로 출력
    wall_ongoing = wall_info
else:
    wall_ongoing = pd.concat([pd.Series(name_output, name = 'Name'), pd.Series(floor_output, name = 'Floor'), wall_info_output], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties        

###############################################################################################################
##############################   여기까진 Wall 계산.py 랑 똑같음   #############################################
############################################################################################################### 

#%% Max값의 행, 열 index 찾는 함수
def FindPosition_max(series_max, txt_name):  # Series형태로 입력
    time_history = pd.read_csv(txt_name, skiprows=len(series_max)+7, header=None, usecols=[i for i in range(1,len(series_max)+1)])[:-2]  # header 어케하는게 좋을까?    
    
    row_index = []
    for current_col_index in series_max.index:
        row_index.append(time_history.loc[:, current_col_index].idxmax()) 
    
    # row_index = pd.Series(series_max.index).apply(lambda x: time_history.loc[:, x].idxmax())
    # row_index = time_history.idxmax(axis='rows') 
           
    return row_index, series_max.index.tolist()
        # return pd.concat([pd.Series(row_index), pd.Series(series_max_or_min.index)], axis=1)

#%% Min값의 행, 열 index 찾는 함수
def FindPosition_min(series_min, txt_name):  # Series형태로 입력
    time_history = pd.read_csv(txt_name, skiprows=len(series_min)+7, header=None, usecols=[i for i in range(1, len(series_min)+1)])[:-2]  # header 어케하는게 좋을까?    

    row_index = []
    for current_col_index in series_min.index:
        row_index.append(time_history.loc[:, current_col_index].idxmin())
        
    return row_index, series_min.index.tolist()
        # return pd.concat([pd.Series(row_index), pd.Series(series_max_or_min.index)], axis=1)

#%% H1, H2의 max와 같은 위치에 있는 V값 찾는 함수
def MatchV(series_max_or_min, max_or_min_position, v_txt_name):
    v_time_history = pd.read_csv(v_txt_name, skiprows=len(series_max_or_min)+7, header=None, usecols=[i for i in range(1, len(series_max_or_min)+1)])[:-2]
    
    V = []
    for row_index, col_index in zip(max_or_min_position.iloc[0,:], max_or_min_position.iloc[1,:]):
        V.append(v_time_history.loc[row_index, col_index])
        
    return V
        
# %% Section Name에 맞는 Property 불러오기

# H1, H2의 텍스트 파일 load (element별 height이 포함된 행만 load)
H1_DE11_raw_txt = H1_txt_dir+'\\'+DE11_txt

# Section name열 추출
section_name = pd.read_csv(H1_DE11_raw_txt, skiprows=7, error_bad_lines=False, header=None).iloc[:, -1]
section_name.reset_index(drop=True, inplace=True)
section_name.name = 'Name'

# Section Name을 각각 Wall name, Floor 로 나누기 (이름의 형태에 따라 달라질수있음!!!)
wall_name = []
floor = []
for i in section_name:
    i = i.strip()  # section_name 앞에 있는 blank 제거
    
    if '_' in i:
        wall_name.append(i.split('_')[0])
        floor.append(i.split('_')[-1])
        
    else:
        wall_name.append(i)
        floor.append(None)
           
SF_ongoing = pd.concat([pd.Series(wall_name, name = 'Name'), pd.Series(floor, name = 'Floor')], axis=1)

# Section Name에 맞는 properties를 Input Sheet에서 불러온 wall 정보와 매치
SF_ongoing = pd.merge(SF_ongoing, wall_ongoing, how='left', on = ['Name', 'Floor'])  # Name과 Floor 두 열을 기준으로 SF_ongoing에 wall_ongoing을 붙여넣음

# SF_ongoing = SF_ongoing.dropna()

##### 섹션 네이밍 특이해서 이번 프로젝트에서만 사용 #######
# 제일 쉬운법 : section name추출하는 지진파 txt파일(여기서는 DE11)의 section name을 바꿔서 불러옴

SF_ongoing = pd.merge(SF_ongoing, concrete_database.iloc[:,0:2], how='left',\
                      left_on='Concrete Strength(CXX)', right_on='Name')

SF_ongoing = pd.merge(SF_ongoing, steel_database.iloc[:,0:2], how='left',\
                  left_on='H. Rebar Strength(SDXXX)', right_on='Name')
    
SF_ongoing = pd.merge(SF_ongoing, steel_geometry_database.iloc[:,[0,2]], how='left',\
                  left_on='Horizontal Rebar(DXX)', right_on='Name')   

#%% 길이 수정
# SF_ongoing['Length'] = length_tobe_added   

#%% Maximum, Minimum 행(마지막 2행) 추출

# H1의 min, max 추출
H1_file_names = os.listdir(H1_txt_dir)

H1_minmax = pd.DataFrame()
for file_name in H1_file_names:
     H1_txt_path = H1_txt_dir + '\\%s' %file_name
     with open(H1_txt_path) as f:
          q = deque(f, 2)
          H1_minmax = H1_minmax.append(pd.read_csv(StringIO(''.join(q)), header=None, index_col=0))

# H2의 min, max 추출
H2_file_names = os.listdir(H2_txt_dir)

H2_minmax = pd.DataFrame()
for file_name in H2_file_names:
     H2_txt_path = H2_txt_dir + '\\%s' %file_name
     with open(H2_txt_path) as f:
          q = deque(f, 2)
          H2_minmax = H2_minmax.append(pd.read_csv(StringIO(''.join(q)), header=None, index_col=0))

# Min, max값 분리
H1_max = H1_minmax.iloc[0:55:2, :]
H1_min = H1_minmax.iloc[1:56:2, :]

H2_max = H2_minmax.iloc[0:55:2, :]
H2_min = H2_minmax.iloc[1:56:2, :]

#%% DE11

# Max와 Min의 행, 열 index 찾기
H1_DE11_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[0, :], H1_txt_dir+'\\'+DE11_txt))  # tuple -> dataframe
H1_DE11_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[0, :], H1_txt_dir+'\\'+DE11_txt))

H2_DE11_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[0, :], H2_txt_dir+'\\'+DE11_txt))
H2_DE11_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[0, :], H2_txt_dir+'\\'+DE11_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE11_V_max = MatchV(H1_max.iloc[0, :], H1_DE11_max_pos, V_txt_dir+'\\'+DE11_txt)  # 최종적으로 계산에 사용할 축력
H1_DE11_V_min = MatchV(H1_min.iloc[0, :], H1_DE11_min_pos, V_txt_dir+'\\'+DE11_txt)

H2_DE11_V_max = MatchV(H2_max.iloc[0, :], H2_DE11_max_pos, V_txt_dir+'\\'+DE11_txt)
H2_DE11_V_min = MatchV(H2_min.iloc[0, :], H2_DE11_min_pos, V_txt_dir+'\\'+DE11_txt)

# max, min transpose하고 index 리셋
H1_DE11_V_minmax = pd.concat([pd.Series(H1_DE11_V_max), pd.Series(H1_DE11_V_min)], axis=1)
H2_DE11_V_minmax = pd.concat([pd.Series(H2_DE11_V_max), pd.Series(H2_DE11_V_min)], axis=1)

# %% DE12

# Max와 Min의 행, 열 index 찾기
H1_DE12_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[1, :], H1_txt_dir+'\\'+DE12_txt))  # tuple -> dataframe
H1_DE12_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[1, :], H1_txt_dir+'\\'+DE12_txt))

H2_DE12_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[1, :], H2_txt_dir+'\\'+DE12_txt))
H2_DE12_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[1, :], H2_txt_dir+'\\'+DE12_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE12_V_max = MatchV(H1_max.iloc[1, :], H1_DE12_max_pos, V_txt_dir+'\\'+DE12_txt)  # 최종적으로 계산에 사용할 축력
H1_DE12_V_min = MatchV(H1_min.iloc[1, :], H1_DE12_min_pos, V_txt_dir+'\\'+DE12_txt)

H2_DE12_V_max = MatchV(H2_max.iloc[1, :], H2_DE12_max_pos, V_txt_dir+'\\'+DE12_txt)
H2_DE12_V_min = MatchV(H2_min.iloc[1, :], H2_DE12_min_pos, V_txt_dir+'\\'+DE12_txt)

# max, min transpose하고 index 리셋
H1_DE12_V_minmax = pd.concat([pd.Series(H1_DE12_V_max), pd.Series(H1_DE12_V_min)], axis=1)
H2_DE12_V_minmax = pd.concat([pd.Series(H2_DE12_V_max), pd.Series(H2_DE12_V_min)], axis=1)


# V값을 min, max에 따라 분리
H1_V_min = pd.concat([H1_DE11_V_minmax.iloc[:,1], H1_DE12_V_minmax.iloc[:,1]], keys=['DE11', 'DE12'], axis=1)
H1_V_max = pd.concat([H1_DE11_V_minmax.iloc[:,0], H1_DE12_V_minmax.iloc[:,0]], keys=['DE11', 'DE12'], axis=1)

H2_V_min = pd.concat([H2_DE11_V_minmax.iloc[:,1], H2_DE12_V_minmax.iloc[:,1]], keys=['DE11', 'DE12'], axis=1)
H2_V_max = pd.concat([H2_DE11_V_minmax.iloc[:,0], H2_DE12_V_minmax.iloc[:,0]], keys=['DE11', 'DE12'], axis=1)

# %% DE21

# Max와 Min의 행, 열 index 찾기
H1_DE21_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[2, :], H1_txt_dir+'\\'+DE21_txt))  # tuple -> dataframe
H1_DE21_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[2, :], H1_txt_dir+'\\'+DE21_txt))

H2_DE21_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[2, :], H2_txt_dir+'\\'+DE21_txt))
H2_DE21_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[2, :], H2_txt_dir+'\\'+DE21_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE21_V_max = MatchV(H1_max.iloc[2, :], H1_DE21_max_pos, V_txt_dir+'\\'+DE21_txt)  # 최종적으로 계산에 사용할 축력
H1_DE21_V_min = MatchV(H1_min.iloc[2, :], H1_DE21_min_pos, V_txt_dir+'\\'+DE21_txt)

H2_DE21_V_max = MatchV(H2_max.iloc[2, :], H2_DE21_max_pos, V_txt_dir+'\\'+DE21_txt)
H2_DE21_V_min = MatchV(H2_min.iloc[2, :], H2_DE21_min_pos, V_txt_dir+'\\'+DE21_txt)

# max, min transpose하고 index 리셋
H1_DE21_V_minmax = pd.concat([pd.Series(H1_DE21_V_max), pd.Series(H1_DE21_V_min)], axis=1)
H2_DE21_V_minmax = pd.concat([pd.Series(H2_DE21_V_max), pd.Series(H2_DE21_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE21'] = H1_DE21_V_minmax.iloc[:, 1]
H1_V_max['DE21'] = H1_DE21_V_minmax.iloc[:, 0]

H2_V_min['DE21'] = H2_DE21_V_minmax.iloc[:, 1]
H2_V_max['DE21'] = H2_DE21_V_minmax.iloc[:, 0]

# %% DE22

# Max와 Min의 행, 열 index 찾기
H1_DE22_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[3, :], H1_txt_dir+'\\'+DE22_txt))  # tuple -> dataframe
H1_DE22_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[3, :], H1_txt_dir+'\\'+DE22_txt))

H2_DE22_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[3, :], H2_txt_dir+'\\'+DE22_txt))
H2_DE22_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[3, :], H2_txt_dir+'\\'+DE22_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE22_V_max = MatchV(H1_max.iloc[3, :], H1_DE22_max_pos, V_txt_dir+'\\'+DE22_txt)  # 최종적으로 계산에 사용할 축력
H1_DE22_V_min = MatchV(H1_min.iloc[3, :], H1_DE22_min_pos, V_txt_dir+'\\'+DE22_txt)

H2_DE22_V_max = MatchV(H2_max.iloc[3, :], H2_DE22_max_pos, V_txt_dir+'\\'+DE22_txt)
H2_DE22_V_min = MatchV(H2_min.iloc[3, :], H2_DE22_min_pos, V_txt_dir+'\\'+DE22_txt)

# max, min transpose하고 index 리셋
H1_DE22_V_minmax = pd.concat([pd.Series(H1_DE22_V_max), pd.Series(H1_DE22_V_min)], axis=1)
H2_DE22_V_minmax = pd.concat([pd.Series(H2_DE22_V_max), pd.Series(H2_DE22_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE22'] = H1_DE22_V_minmax.iloc[:, 1]
H1_V_max['DE22'] = H1_DE22_V_minmax.iloc[:, 0]

H2_V_min['DE22'] = H2_DE22_V_minmax.iloc[:, 1]
H2_V_max['DE22'] = H2_DE22_V_minmax.iloc[:, 0]

# %% DE31

# Max와 Min의 행, 열 index 찾기
H1_DE31_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[4, :], H1_txt_dir+'\\'+DE31_txt))  # tuple -> dataframe
H1_DE31_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[4, :], H1_txt_dir+'\\'+DE31_txt))

H2_DE31_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[4, :], H2_txt_dir+'\\'+DE31_txt))
H2_DE31_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[4, :], H2_txt_dir+'\\'+DE31_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE31_V_max = MatchV(H1_max.iloc[4, :], H1_DE31_max_pos, V_txt_dir+'\\'+DE31_txt)  # 최종적으로 계산에 사용할 축력
H1_DE31_V_min = MatchV(H1_min.iloc[4, :], H1_DE31_min_pos, V_txt_dir+'\\'+DE31_txt)

H2_DE31_V_max = MatchV(H2_max.iloc[4, :], H2_DE31_max_pos, V_txt_dir+'\\'+DE31_txt)
H2_DE31_V_min = MatchV(H2_min.iloc[4, :], H2_DE31_min_pos, V_txt_dir+'\\'+DE31_txt)

# max, min transpose하고 index 리셋
H1_DE31_V_minmax = pd.concat([pd.Series(H1_DE31_V_max), pd.Series(H1_DE31_V_min)], axis=1)
H2_DE31_V_minmax = pd.concat([pd.Series(H2_DE31_V_max), pd.Series(H2_DE31_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE31'] = H1_DE31_V_minmax.iloc[:, 1]
H1_V_max['DE31'] = H1_DE31_V_minmax.iloc[:, 0]

H2_V_min['DE31'] = H2_DE31_V_minmax.iloc[:, 1]
H2_V_max['DE31'] = H2_DE31_V_minmax.iloc[:, 0]

# %% DE32

# Max와 Min의 행, 열 index 찾기
H1_DE32_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[5, :], H1_txt_dir+'\\'+DE32_txt))  # tuple -> dataframe
H1_DE32_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[5, :], H1_txt_dir+'\\'+DE32_txt))

H2_DE32_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[5, :], H2_txt_dir+'\\'+DE32_txt))
H2_DE32_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[5, :], H2_txt_dir+'\\'+DE32_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE32_V_max = MatchV(H1_max.iloc[5, :], H1_DE32_max_pos, V_txt_dir+'\\'+DE32_txt)  # 최종적으로 계산에 사용할 축력
H1_DE32_V_min = MatchV(H1_min.iloc[5, :], H1_DE32_min_pos, V_txt_dir+'\\'+DE32_txt)

H2_DE32_V_max = MatchV(H2_max.iloc[5, :], H2_DE32_max_pos, V_txt_dir+'\\'+DE32_txt)
H2_DE32_V_min = MatchV(H2_min.iloc[5, :], H2_DE32_min_pos, V_txt_dir+'\\'+DE32_txt)

# max, min transpose하고 index 리셋
H1_DE32_V_minmax = pd.concat([pd.Series(H1_DE32_V_max), pd.Series(H1_DE32_V_min)], axis=1)
H2_DE32_V_minmax = pd.concat([pd.Series(H2_DE32_V_max), pd.Series(H2_DE32_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE32'] = H1_DE32_V_minmax.iloc[:, 1]
H1_V_max['DE32'] = H1_DE32_V_minmax.iloc[:, 0]

H2_V_min['DE32'] = H2_DE32_V_minmax.iloc[:, 1]
H2_V_max['DE32'] = H2_DE32_V_minmax.iloc[:, 0]

# %% DE41

# Max와 Min의 행, 열 index 찾기
H1_DE41_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[6, :], H1_txt_dir+'\\'+DE41_txt))  # tuple -> dataframe
H1_DE41_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[6, :], H1_txt_dir+'\\'+DE41_txt))

H2_DE41_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[6, :], H2_txt_dir+'\\'+DE41_txt))
H2_DE41_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[6, :], H2_txt_dir+'\\'+DE41_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE41_V_max = MatchV(H1_max.iloc[6, :], H1_DE41_max_pos, V_txt_dir+'\\'+DE41_txt)  # 최종적으로 계산에 사용할 축력
H1_DE41_V_min = MatchV(H1_min.iloc[6, :], H1_DE41_min_pos, V_txt_dir+'\\'+DE41_txt)

H2_DE41_V_max = MatchV(H2_max.iloc[6, :], H2_DE41_max_pos, V_txt_dir+'\\'+DE41_txt)
H2_DE41_V_min = MatchV(H2_min.iloc[6, :], H2_DE41_min_pos, V_txt_dir+'\\'+DE41_txt)

# max, min transpose하고 index 리셋
H1_DE41_V_minmax = pd.concat([pd.Series(H1_DE41_V_max), pd.Series(H1_DE41_V_min)], axis=1)
H2_DE41_V_minmax = pd.concat([pd.Series(H2_DE41_V_max), pd.Series(H2_DE41_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE41'] = H1_DE41_V_minmax.iloc[:, 1]
H1_V_max['DE41'] = H1_DE41_V_minmax.iloc[:, 0]

H2_V_min['DE41'] = H2_DE41_V_minmax.iloc[:, 1]
H2_V_max['DE41'] = H2_DE41_V_minmax.iloc[:, 0]

# %% DE42

# Max와 Min의 행, 열 index 찾기
H1_DE42_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[7, :], H1_txt_dir+'\\'+DE42_txt))  # tuple -> dataframe
H1_DE42_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[7, :], H1_txt_dir+'\\'+DE42_txt))

H2_DE42_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[7, :], H2_txt_dir+'\\'+DE42_txt))
H2_DE42_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[7, :], H2_txt_dir+'\\'+DE42_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE42_V_max = MatchV(H1_max.iloc[7, :], H1_DE42_max_pos, V_txt_dir+'\\'+DE42_txt)  # 최종적으로 계산에 사용할 축력
H1_DE42_V_min = MatchV(H1_min.iloc[7, :], H1_DE42_min_pos, V_txt_dir+'\\'+DE42_txt)

H2_DE42_V_max = MatchV(H2_max.iloc[7, :], H2_DE42_max_pos, V_txt_dir+'\\'+DE42_txt)
H2_DE42_V_min = MatchV(H2_min.iloc[7, :], H2_DE42_min_pos, V_txt_dir+'\\'+DE42_txt)

# max, min transpose하고 index 리셋
H1_DE42_V_minmax = pd.concat([pd.Series(H1_DE42_V_max), pd.Series(H1_DE42_V_min)], axis=1)
H2_DE42_V_minmax = pd.concat([pd.Series(H2_DE42_V_max), pd.Series(H2_DE42_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE42'] = H1_DE42_V_minmax.iloc[:, 1]
H1_V_max['DE42'] = H1_DE42_V_minmax.iloc[:, 0]

H2_V_min['DE42'] = H2_DE42_V_minmax.iloc[:, 1]
H2_V_max['DE42'] = H2_DE42_V_minmax.iloc[:, 0]

# %% DE51

# Max와 Min의 행, 열 index 찾기
H1_DE51_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[8, :], H1_txt_dir+'\\'+DE51_txt))  # tuple -> dataframe
H1_DE51_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[8, :], H1_txt_dir+'\\'+DE51_txt))

H2_DE51_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[8, :], H2_txt_dir+'\\'+DE51_txt))
H2_DE51_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[8, :], H2_txt_dir+'\\'+DE51_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE51_V_max = MatchV(H1_max.iloc[8, :], H1_DE51_max_pos, V_txt_dir+'\\'+DE51_txt)  # 최종적으로 계산에 사용할 축력
H1_DE51_V_min = MatchV(H1_min.iloc[8, :], H1_DE51_min_pos, V_txt_dir+'\\'+DE51_txt)

H2_DE51_V_max = MatchV(H2_max.iloc[8, :], H2_DE51_max_pos, V_txt_dir+'\\'+DE51_txt)
H2_DE51_V_min = MatchV(H2_min.iloc[8, :], H2_DE51_min_pos, V_txt_dir+'\\'+DE51_txt)

# max, min transpose하고 index 리셋
H1_DE51_V_minmax = pd.concat([pd.Series(H1_DE51_V_max), pd.Series(H1_DE51_V_min)], axis=1)
H2_DE51_V_minmax = pd.concat([pd.Series(H2_DE51_V_max), pd.Series(H2_DE51_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE51'] = H1_DE51_V_minmax.iloc[:, 1]
H1_V_max['DE51'] = H1_DE51_V_minmax.iloc[:, 0]

H2_V_min['DE51'] = H2_DE51_V_minmax.iloc[:, 1]
H2_V_max['DE51'] = H2_DE51_V_minmax.iloc[:, 0]

# %% DE52

# Max와 Min의 행, 열 index 찾기
H1_DE52_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[9, :], H1_txt_dir+'\\'+DE52_txt))  # tuple -> dataframe
H1_DE52_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[9, :], H1_txt_dir+'\\'+DE52_txt))

H2_DE52_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[9, :], H2_txt_dir+'\\'+DE52_txt))
H2_DE52_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[9, :], H2_txt_dir+'\\'+DE52_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE52_V_max = MatchV(H1_max.iloc[9, :], H1_DE52_max_pos, V_txt_dir+'\\'+DE52_txt)  # 최종적으로 계산에 사용할 축력
H1_DE52_V_min = MatchV(H1_min.iloc[9, :], H1_DE52_min_pos, V_txt_dir+'\\'+DE52_txt)

H2_DE52_V_max = MatchV(H2_max.iloc[9, :], H2_DE52_max_pos, V_txt_dir+'\\'+DE52_txt)
H2_DE52_V_min = MatchV(H2_min.iloc[9, :], H2_DE52_min_pos, V_txt_dir+'\\'+DE52_txt)

# max, min transpose하고 index 리셋
H1_DE52_V_minmax = pd.concat([pd.Series(H1_DE52_V_max), pd.Series(H1_DE52_V_min)], axis=1)
H2_DE52_V_minmax = pd.concat([pd.Series(H2_DE52_V_max), pd.Series(H2_DE52_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE52'] = H1_DE52_V_minmax.iloc[:, 1]
H1_V_max['DE52'] = H1_DE52_V_minmax.iloc[:, 0]

H2_V_min['DE52'] = H2_DE52_V_minmax.iloc[:, 1]
H2_V_max['DE52'] = H2_DE52_V_minmax.iloc[:, 0]

# %% DE61

# Max와 Min의 행, 열 index 찾기
H1_DE61_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[10, :], H1_txt_dir+'\\'+DE61_txt))  # tuple -> dataframe
H1_DE61_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[10, :], H1_txt_dir+'\\'+DE61_txt))

H2_DE61_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[10, :], H2_txt_dir+'\\'+DE61_txt))
H2_DE61_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[10, :], H2_txt_dir+'\\'+DE61_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE61_V_max = MatchV(H1_max.iloc[10, :], H1_DE61_max_pos, V_txt_dir+'\\'+DE61_txt)  # 최종적으로 계산에 사용할 축력
H1_DE61_V_min = MatchV(H1_min.iloc[10, :], H1_DE61_min_pos, V_txt_dir+'\\'+DE61_txt)

H2_DE61_V_max = MatchV(H2_max.iloc[10, :], H2_DE61_max_pos, V_txt_dir+'\\'+DE61_txt)
H2_DE61_V_min = MatchV(H2_min.iloc[10, :], H2_DE61_min_pos, V_txt_dir+'\\'+DE61_txt)

# max, min transpose하고 index 리셋
H1_DE61_V_minmax = pd.concat([pd.Series(H1_DE61_V_max), pd.Series(H1_DE61_V_min)], axis=1)
H2_DE61_V_minmax = pd.concat([pd.Series(H2_DE61_V_max), pd.Series(H2_DE61_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE61'] = H1_DE61_V_minmax.iloc[:, 1]
H1_V_max['DE61'] = H1_DE61_V_minmax.iloc[:, 0]

H2_V_min['DE61'] = H2_DE61_V_minmax.iloc[:, 1]
H2_V_max['DE61'] = H2_DE61_V_minmax.iloc[:, 0]

# %% DE62

# Max와 Min의 행, 열 index 찾기
H1_DE62_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[11, :], H1_txt_dir+'\\'+DE62_txt))  # tuple -> dataframe
H1_DE62_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[11, :], H1_txt_dir+'\\'+DE62_txt))

H2_DE62_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[11, :], H2_txt_dir+'\\'+DE62_txt))
H2_DE62_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[11, :], H2_txt_dir+'\\'+DE62_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE62_V_max = MatchV(H1_max.iloc[11, :], H1_DE62_max_pos, V_txt_dir+'\\'+DE62_txt)  # 최종적으로 계산에 사용할 축력
H1_DE62_V_min = MatchV(H1_min.iloc[11, :], H1_DE62_min_pos, V_txt_dir+'\\'+DE62_txt)

H2_DE62_V_max = MatchV(H2_max.iloc[11, :], H2_DE62_max_pos, V_txt_dir+'\\'+DE62_txt)
H2_DE62_V_min = MatchV(H2_min.iloc[11, :], H2_DE62_min_pos, V_txt_dir+'\\'+DE62_txt)

# max, min transpose하고 index 리셋
H1_DE62_V_minmax = pd.concat([pd.Series(H1_DE62_V_max), pd.Series(H1_DE62_V_min)], axis=1)
H2_DE62_V_minmax = pd.concat([pd.Series(H2_DE62_V_max), pd.Series(H2_DE62_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE62'] = H1_DE62_V_minmax.iloc[:, 1]
H1_V_max['DE62'] = H1_DE62_V_minmax.iloc[:, 0]

H2_V_min['DE62'] = H2_DE62_V_minmax.iloc[:, 1]
H2_V_max['DE62'] = H2_DE62_V_minmax.iloc[:, 0]

# %% DE71

# Max와 Min의 행, 열 index 찾기
H1_DE71_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[12, :], H1_txt_dir+'\\'+DE71_txt))  # tuple -> dataframe
H1_DE71_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[12, :], H1_txt_dir+'\\'+DE71_txt))

H2_DE71_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[12, :], H2_txt_dir+'\\'+DE71_txt))
H2_DE71_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[12, :], H2_txt_dir+'\\'+DE71_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE71_V_max = MatchV(H1_max.iloc[12, :], H1_DE71_max_pos, V_txt_dir+'\\'+DE71_txt)  # 최종적으로 계산에 사용할 축력
H1_DE71_V_min = MatchV(H1_min.iloc[12, :], H1_DE71_min_pos, V_txt_dir+'\\'+DE71_txt)

H2_DE71_V_max = MatchV(H2_max.iloc[12, :], H2_DE71_max_pos, V_txt_dir+'\\'+DE71_txt)
H2_DE71_V_min = MatchV(H2_min.iloc[12, :], H2_DE71_min_pos, V_txt_dir+'\\'+DE71_txt)

# max, min transpose하고 index 리셋
H1_DE71_V_minmax = pd.concat([pd.Series(H1_DE71_V_max), pd.Series(H1_DE71_V_min)], axis=1)
H2_DE71_V_minmax = pd.concat([pd.Series(H2_DE71_V_max), pd.Series(H2_DE71_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE71'] = H1_DE71_V_minmax.iloc[:, 1]
H1_V_max['DE71'] = H1_DE71_V_minmax.iloc[:, 0]

H2_V_min['DE71'] = H2_DE71_V_minmax.iloc[:, 1]
H2_V_max['DE71'] = H2_DE71_V_minmax.iloc[:, 0]

# %% DE72

# Max와 Min의 행, 열 index 찾기
H1_DE72_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[13, :], H1_txt_dir+'\\'+DE72_txt))  # tuple -> dataframe
H1_DE72_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[13, :], H1_txt_dir+'\\'+DE72_txt))

H2_DE72_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[13, :], H2_txt_dir+'\\'+DE72_txt))
H2_DE72_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[13, :], H2_txt_dir+'\\'+DE72_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_DE72_V_max = MatchV(H1_max.iloc[13, :], H1_DE72_max_pos, V_txt_dir+'\\'+DE72_txt)  # 최종적으로 계산에 사용할 축력
H1_DE72_V_min = MatchV(H1_min.iloc[13, :], H1_DE72_min_pos, V_txt_dir+'\\'+DE72_txt)

H2_DE72_V_max = MatchV(H2_max.iloc[13, :], H2_DE72_max_pos, V_txt_dir+'\\'+DE72_txt)
H2_DE72_V_min = MatchV(H2_min.iloc[13, :], H2_DE72_min_pos, V_txt_dir+'\\'+DE72_txt)

# max, min transpose하고 index 리셋
H1_DE72_V_minmax = pd.concat([pd.Series(H1_DE72_V_max), pd.Series(H1_DE72_V_min)], axis=1)
H2_DE72_V_minmax = pd.concat([pd.Series(H2_DE72_V_max), pd.Series(H2_DE72_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['DE72'] = H1_DE72_V_minmax.iloc[:, 1]
H1_V_max['DE72'] = H1_DE72_V_minmax.iloc[:, 0]

H2_V_min['DE72'] = H2_DE72_V_minmax.iloc[:, 1]
H2_V_max['DE72'] = H2_DE72_V_minmax.iloc[:, 0]

# %% MCE11

# Max와 Min의 행, 열 index 찾기
H1_MCE11_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[0, :], H1_txt_dir+'\\'+MCE11_txt))  # tuple -> dataframe
H1_MCE11_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[0, :], H1_txt_dir+'\\'+MCE11_txt))

H2_MCE11_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[0, :], H2_txt_dir+'\\'+MCE11_txt))
H2_MCE11_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[0, :], H2_txt_dir+'\\'+MCE11_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE11_V_max = MatchV(H1_max.iloc[0, :], H1_MCE11_max_pos, V_txt_dir+'\\'+MCE11_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE11_V_min = MatchV(H1_min.iloc[0, :], H1_MCE11_min_pos, V_txt_dir+'\\'+MCE11_txt)

H2_MCE11_V_max = MatchV(H2_max.iloc[0, :], H2_MCE11_max_pos, V_txt_dir+'\\'+MCE11_txt)
H2_MCE11_V_min = MatchV(H2_min.iloc[0, :], H2_MCE11_min_pos, V_txt_dir+'\\'+MCE11_txt)

# max, min transpose하고 index 리셋
H1_MCE11_V_minmax = pd.concat([pd.Series(H1_MCE11_V_max), pd.Series(H1_MCE11_V_min)], axis=1)
H2_MCE11_V_minmax = pd.concat([pd.Series(H2_MCE11_V_max), pd.Series(H2_MCE11_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE11'] = H1_MCE11_V_minmax.iloc[:, 1]
H1_V_max['MCE11'] = H1_MCE11_V_minmax.iloc[:, 0]

H2_V_min['MCE11'] = H2_MCE11_V_minmax.iloc[:, 1]
H2_V_max['MCE11'] = H2_MCE11_V_minmax.iloc[:, 0]

# %% MCE12

# Max와 Min의 행, 열 index 찾기
H1_MCE12_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[1, :], H1_txt_dir+'\\'+MCE12_txt))  # tuple -> dataframe
H1_MCE12_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[1, :], H1_txt_dir+'\\'+MCE12_txt))

H2_MCE12_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[1, :], H2_txt_dir+'\\'+MCE12_txt))
H2_MCE12_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[1, :], H2_txt_dir+'\\'+MCE12_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE12_V_max = MatchV(H1_max.iloc[1, :], H1_MCE12_max_pos, V_txt_dir+'\\'+MCE12_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE12_V_min = MatchV(H1_min.iloc[1, :], H1_MCE12_min_pos, V_txt_dir+'\\'+MCE12_txt)

H2_MCE12_V_max = MatchV(H2_max.iloc[1, :], H2_MCE12_max_pos, V_txt_dir+'\\'+MCE12_txt)
H2_MCE12_V_min = MatchV(H2_min.iloc[1, :], H2_MCE12_min_pos, V_txt_dir+'\\'+MCE12_txt)

# max, min transpose하고 index 리셋
H1_MCE12_V_minmax = pd.concat([pd.Series(H1_MCE12_V_max), pd.Series(H1_MCE12_V_min)], axis=1)
H2_MCE12_V_minmax = pd.concat([pd.Series(H2_MCE12_V_max), pd.Series(H2_MCE12_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE12'] = H1_MCE12_V_minmax.iloc[:, 1]
H1_V_max['MCE12'] = H1_MCE12_V_minmax.iloc[:, 0]

H2_V_min['MCE12'] = H2_MCE12_V_minmax.iloc[:, 1]
H2_V_max['MCE12'] = H2_MCE12_V_minmax.iloc[:, 0]

# %% MCE21

# Max와 Min의 행, 열 index 찾기
H1_MCE21_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[2, :], H1_txt_dir+'\\'+MCE21_txt))  # tuple -> dataframe
H1_MCE21_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[2, :], H1_txt_dir+'\\'+MCE21_txt))

H2_MCE21_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[2, :], H2_txt_dir+'\\'+MCE21_txt))
H2_MCE21_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[2, :], H2_txt_dir+'\\'+MCE21_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE21_V_max = MatchV(H1_max.iloc[2, :], H1_MCE21_max_pos, V_txt_dir+'\\'+MCE21_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE21_V_min = MatchV(H1_min.iloc[2, :], H1_MCE21_min_pos, V_txt_dir+'\\'+MCE21_txt)

H2_MCE21_V_max = MatchV(H2_max.iloc[2, :], H2_MCE21_max_pos, V_txt_dir+'\\'+MCE21_txt)
H2_MCE21_V_min = MatchV(H2_min.iloc[2, :], H2_MCE21_min_pos, V_txt_dir+'\\'+MCE21_txt)

# max, min transpose하고 index 리셋
H1_MCE21_V_minmax = pd.concat([pd.Series(H1_MCE21_V_max), pd.Series(H1_MCE21_V_min)], axis=1)
H2_MCE21_V_minmax = pd.concat([pd.Series(H2_MCE21_V_max), pd.Series(H2_MCE21_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE21'] = H1_MCE21_V_minmax.iloc[:, 1]
H1_V_max['MCE21'] = H1_MCE21_V_minmax.iloc[:, 0]

H2_V_min['MCE21'] = H2_MCE21_V_minmax.iloc[:, 1]
H2_V_max['MCE21'] = H2_MCE21_V_minmax.iloc[:, 0]

# %% MCE22

# Max와 Min의 행, 열 index 찾기
H1_MCE22_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[3, :], H1_txt_dir+'\\'+MCE22_txt))  # tuple -> dataframe
H1_MCE22_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[3, :], H1_txt_dir+'\\'+MCE22_txt))

H2_MCE22_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[3, :], H2_txt_dir+'\\'+MCE22_txt))
H2_MCE22_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[3, :], H2_txt_dir+'\\'+MCE22_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE22_V_max = MatchV(H1_max.iloc[3, :], H1_MCE22_max_pos, V_txt_dir+'\\'+MCE22_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE22_V_min = MatchV(H1_min.iloc[3, :], H1_MCE22_min_pos, V_txt_dir+'\\'+MCE22_txt)

H2_MCE22_V_max = MatchV(H2_max.iloc[3, :], H2_MCE22_max_pos, V_txt_dir+'\\'+MCE22_txt)
H2_MCE22_V_min = MatchV(H2_min.iloc[3, :], H2_MCE22_min_pos, V_txt_dir+'\\'+MCE22_txt)

# max, min transpose하고 index 리셋
H1_MCE22_V_minmax = pd.concat([pd.Series(H1_MCE22_V_max), pd.Series(H1_MCE22_V_min)], axis=1)
H2_MCE22_V_minmax = pd.concat([pd.Series(H2_MCE22_V_max), pd.Series(H2_MCE22_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE22'] = H1_MCE22_V_minmax.iloc[:, 1]
H1_V_max['MCE22'] = H1_MCE22_V_minmax.iloc[:, 0]

H2_V_min['MCE22'] = H2_MCE22_V_minmax.iloc[:, 1]
H2_V_max['MCE22'] = H2_MCE22_V_minmax.iloc[:, 0]

# %% MCE31

# Max와 Min의 행, 열 index 찾기
H1_MCE31_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[4, :], H1_txt_dir+'\\'+MCE31_txt))  # tuple -> dataframe
H1_MCE31_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[4, :], H1_txt_dir+'\\'+MCE31_txt))

H2_MCE31_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[4, :], H2_txt_dir+'\\'+MCE31_txt))
H2_MCE31_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[4, :], H2_txt_dir+'\\'+MCE31_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE31_V_max = MatchV(H1_max.iloc[4, :], H1_MCE31_max_pos, V_txt_dir+'\\'+MCE31_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE31_V_min = MatchV(H1_min.iloc[4, :], H1_MCE31_min_pos, V_txt_dir+'\\'+MCE31_txt)

H2_MCE31_V_max = MatchV(H2_max.iloc[4, :], H2_MCE31_max_pos, V_txt_dir+'\\'+MCE31_txt)
H2_MCE31_V_min = MatchV(H2_min.iloc[4, :], H2_MCE31_min_pos, V_txt_dir+'\\'+MCE31_txt)

# max, min transpose하고 index 리셋
H1_MCE31_V_minmax = pd.concat([pd.Series(H1_MCE31_V_max), pd.Series(H1_MCE31_V_min)], axis=1)
H2_MCE31_V_minmax = pd.concat([pd.Series(H2_MCE31_V_max), pd.Series(H2_MCE31_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE31'] = H1_MCE31_V_minmax.iloc[:, 1]
H1_V_max['MCE31'] = H1_MCE31_V_minmax.iloc[:, 0]

H2_V_min['MCE31'] = H2_MCE31_V_minmax.iloc[:, 1]
H2_V_max['MCE31'] = H2_MCE31_V_minmax.iloc[:, 0]

# %% MCE32

# Max와 Min의 행, 열 index 찾기
H1_MCE32_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[5, :], H1_txt_dir+'\\'+MCE32_txt))  # tuple -> dataframe
H1_MCE32_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[5, :], H1_txt_dir+'\\'+MCE32_txt))

H2_MCE32_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[5, :], H2_txt_dir+'\\'+MCE32_txt))
H2_MCE32_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[5, :], H2_txt_dir+'\\'+MCE32_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE32_V_max = MatchV(H1_max.iloc[5, :], H1_MCE32_max_pos, V_txt_dir+'\\'+MCE32_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE32_V_min = MatchV(H1_min.iloc[5, :], H1_MCE32_min_pos, V_txt_dir+'\\'+MCE32_txt)

H2_MCE32_V_max = MatchV(H2_max.iloc[5, :], H2_MCE32_max_pos, V_txt_dir+'\\'+MCE32_txt)
H2_MCE32_V_min = MatchV(H2_min.iloc[5, :], H2_MCE32_min_pos, V_txt_dir+'\\'+MCE32_txt)

# max, min transpose하고 index 리셋
H1_MCE32_V_minmax = pd.concat([pd.Series(H1_MCE32_V_max), pd.Series(H1_MCE32_V_min)], axis=1)
H2_MCE32_V_minmax = pd.concat([pd.Series(H2_MCE32_V_max), pd.Series(H2_MCE32_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE32'] = H1_MCE32_V_minmax.iloc[:, 1]
H1_V_max['MCE32'] = H1_MCE32_V_minmax.iloc[:, 0]

H2_V_min['MCE32'] = H2_MCE32_V_minmax.iloc[:, 1]
H2_V_max['MCE32'] = H2_MCE32_V_minmax.iloc[:, 0]

# %% MCE41

# Max와 Min의 행, 열 index 찾기
H1_MCE41_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[6, :], H1_txt_dir+'\\'+MCE41_txt))  # tuple -> dataframe
H1_MCE41_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[6, :], H1_txt_dir+'\\'+MCE41_txt))

H2_MCE41_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[6, :], H2_txt_dir+'\\'+MCE41_txt))
H2_MCE41_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[6, :], H2_txt_dir+'\\'+MCE41_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE41_V_max = MatchV(H1_max.iloc[6, :], H1_MCE41_max_pos, V_txt_dir+'\\'+MCE41_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE41_V_min = MatchV(H1_min.iloc[6, :], H1_MCE41_min_pos, V_txt_dir+'\\'+MCE41_txt)

H2_MCE41_V_max = MatchV(H2_max.iloc[6, :], H2_MCE41_max_pos, V_txt_dir+'\\'+MCE41_txt)
H2_MCE41_V_min = MatchV(H2_min.iloc[6, :], H2_MCE41_min_pos, V_txt_dir+'\\'+MCE41_txt)

# max, min transpose하고 index 리셋
H1_MCE41_V_minmax = pd.concat([pd.Series(H1_MCE41_V_max), pd.Series(H1_MCE41_V_min)], axis=1)
H2_MCE41_V_minmax = pd.concat([pd.Series(H2_MCE41_V_max), pd.Series(H2_MCE41_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE41'] = H1_MCE41_V_minmax.iloc[:, 1]
H1_V_max['MCE41'] = H1_MCE41_V_minmax.iloc[:, 0]

H2_V_min['MCE41'] = H2_MCE41_V_minmax.iloc[:, 1]
H2_V_max['MCE41'] = H2_MCE41_V_minmax.iloc[:, 0]

# %% MCE42

# Max와 Min의 행, 열 index 찾기
H1_MCE42_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[7, :], H1_txt_dir+'\\'+MCE42_txt))  # tuple -> dataframe
H1_MCE42_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[7, :], H1_txt_dir+'\\'+MCE42_txt))

H2_MCE42_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[7, :], H2_txt_dir+'\\'+MCE42_txt))
H2_MCE42_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[7, :], H2_txt_dir+'\\'+MCE42_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE42_V_max = MatchV(H1_max.iloc[7, :], H1_MCE42_max_pos, V_txt_dir+'\\'+MCE42_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE42_V_min = MatchV(H1_min.iloc[7, :], H1_MCE42_min_pos, V_txt_dir+'\\'+MCE42_txt)

H2_MCE42_V_max = MatchV(H2_max.iloc[7, :], H2_MCE42_max_pos, V_txt_dir+'\\'+MCE42_txt)
H2_MCE42_V_min = MatchV(H2_min.iloc[7, :], H2_MCE42_min_pos, V_txt_dir+'\\'+MCE42_txt)

# max, min transpose하고 index 리셋
H1_MCE42_V_minmax = pd.concat([pd.Series(H1_MCE42_V_max), pd.Series(H1_MCE42_V_min)], axis=1)
H2_MCE42_V_minmax = pd.concat([pd.Series(H2_MCE42_V_max), pd.Series(H2_MCE42_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE42'] = H1_MCE42_V_minmax.iloc[:, 1]
H1_V_max['MCE42'] = H1_MCE42_V_minmax.iloc[:, 0]

H2_V_min['MCE42'] = H2_MCE42_V_minmax.iloc[:, 1]
H2_V_max['MCE42'] = H2_MCE42_V_minmax.iloc[:, 0]

# %% MCE51

# Max와 Min의 행, 열 index 찾기
H1_MCE51_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[8, :], H1_txt_dir+'\\'+MCE51_txt))  # tuple -> dataframe
H1_MCE51_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[8, :], H1_txt_dir+'\\'+MCE51_txt))

H2_MCE51_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[8, :], H2_txt_dir+'\\'+MCE51_txt))
H2_MCE51_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[8, :], H2_txt_dir+'\\'+MCE51_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE51_V_max = MatchV(H1_max.iloc[8, :], H1_MCE51_max_pos, V_txt_dir+'\\'+MCE51_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE51_V_min = MatchV(H1_min.iloc[8, :], H1_MCE51_min_pos, V_txt_dir+'\\'+MCE51_txt)

H2_MCE51_V_max = MatchV(H2_max.iloc[8, :], H2_MCE51_max_pos, V_txt_dir+'\\'+MCE51_txt)
H2_MCE51_V_min = MatchV(H2_min.iloc[8, :], H2_MCE51_min_pos, V_txt_dir+'\\'+MCE51_txt)

# max, min transpose하고 index 리셋
H1_MCE51_V_minmax = pd.concat([pd.Series(H1_MCE51_V_max), pd.Series(H1_MCE51_V_min)], axis=1)
H2_MCE51_V_minmax = pd.concat([pd.Series(H2_MCE51_V_max), pd.Series(H2_MCE51_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE51'] = H1_MCE51_V_minmax.iloc[:, 1]
H1_V_max['MCE51'] = H1_MCE51_V_minmax.iloc[:, 0]

H2_V_min['MCE51'] = H2_MCE51_V_minmax.iloc[:, 1]
H2_V_max['MCE51'] = H2_MCE51_V_minmax.iloc[:, 0]

# %% MCE52

# Max와 Min의 행, 열 index 찾기
H1_MCE52_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[9, :], H1_txt_dir+'\\'+MCE52_txt))  # tuple -> dataframe
H1_MCE52_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[9, :], H1_txt_dir+'\\'+MCE52_txt))

H2_MCE52_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[9, :], H2_txt_dir+'\\'+MCE52_txt))
H2_MCE52_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[9, :], H2_txt_dir+'\\'+MCE52_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE52_V_max = MatchV(H1_max.iloc[9, :], H1_MCE52_max_pos, V_txt_dir+'\\'+MCE52_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE52_V_min = MatchV(H1_min.iloc[9, :], H1_MCE52_min_pos, V_txt_dir+'\\'+MCE52_txt)

H2_MCE52_V_max = MatchV(H2_max.iloc[9, :], H2_MCE52_max_pos, V_txt_dir+'\\'+MCE52_txt)
H2_MCE52_V_min = MatchV(H2_min.iloc[9, :], H2_MCE52_min_pos, V_txt_dir+'\\'+MCE52_txt)

# max, min transpose하고 index 리셋
H1_MCE52_V_minmax = pd.concat([pd.Series(H1_MCE52_V_max), pd.Series(H1_MCE52_V_min)], axis=1)
H2_MCE52_V_minmax = pd.concat([pd.Series(H2_MCE52_V_max), pd.Series(H2_MCE52_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE52'] = H1_MCE52_V_minmax.iloc[:, 1]
H1_V_max['MCE52'] = H1_MCE52_V_minmax.iloc[:, 0]

H2_V_min['MCE52'] = H2_MCE52_V_minmax.iloc[:, 1]
H2_V_max['MCE52'] = H2_MCE52_V_minmax.iloc[:, 0]

# %% MCE61

# Max와 Min의 행, 열 index 찾기
H1_MCE61_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[10, :], H1_txt_dir+'\\'+MCE61_txt))  # tuple -> dataframe
H1_MCE61_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[10, :], H1_txt_dir+'\\'+MCE61_txt))

H2_MCE61_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[10, :], H2_txt_dir+'\\'+MCE61_txt))
H2_MCE61_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[10, :], H2_txt_dir+'\\'+MCE61_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE61_V_max = MatchV(H1_max.iloc[10, :], H1_MCE61_max_pos, V_txt_dir+'\\'+MCE61_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE61_V_min = MatchV(H1_min.iloc[10, :], H1_MCE61_min_pos, V_txt_dir+'\\'+MCE61_txt)

H2_MCE61_V_max = MatchV(H2_max.iloc[10, :], H2_MCE61_max_pos, V_txt_dir+'\\'+MCE61_txt)
H2_MCE61_V_min = MatchV(H2_min.iloc[10, :], H2_MCE61_min_pos, V_txt_dir+'\\'+MCE61_txt)

# max, min transpose하고 index 리셋
H1_MCE61_V_minmax = pd.concat([pd.Series(H1_MCE61_V_max), pd.Series(H1_MCE61_V_min)], axis=1)
H2_MCE61_V_minmax = pd.concat([pd.Series(H2_MCE61_V_max), pd.Series(H2_MCE61_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE61'] = H1_MCE61_V_minmax.iloc[:, 1]
H1_V_max['MCE61'] = H1_MCE61_V_minmax.iloc[:, 0]

H2_V_min['MCE61'] = H2_MCE61_V_minmax.iloc[:, 1]
H2_V_max['MCE61'] = H2_MCE61_V_minmax.iloc[:, 0]

# %% MCE62

# Max와 Min의 행, 열 index 찾기
H1_MCE62_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[11, :], H1_txt_dir+'\\'+MCE62_txt))  # tuple -> dataframe
H1_MCE62_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[11, :], H1_txt_dir+'\\'+MCE62_txt))

H2_MCE62_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[11, :], H2_txt_dir+'\\'+MCE62_txt))
H2_MCE62_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[11, :], H2_txt_dir+'\\'+MCE62_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE62_V_max = MatchV(H1_max.iloc[11, :], H1_MCE62_max_pos, V_txt_dir+'\\'+MCE62_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE62_V_min = MatchV(H1_min.iloc[11, :], H1_MCE62_min_pos, V_txt_dir+'\\'+MCE62_txt)

H2_MCE62_V_max = MatchV(H2_max.iloc[11, :], H2_MCE62_max_pos, V_txt_dir+'\\'+MCE62_txt)
H2_MCE62_V_min = MatchV(H2_min.iloc[11, :], H2_MCE62_min_pos, V_txt_dir+'\\'+MCE62_txt)

# max, min transpose하고 index 리셋
H1_MCE62_V_minmax = pd.concat([pd.Series(H1_MCE62_V_max), pd.Series(H1_MCE62_V_min)], axis=1)
H2_MCE62_V_minmax = pd.concat([pd.Series(H2_MCE62_V_max), pd.Series(H2_MCE62_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE62'] = H1_MCE62_V_minmax.iloc[:, 1]
H1_V_max['MCE62'] = H1_MCE62_V_minmax.iloc[:, 0]

H2_V_min['MCE62'] = H2_MCE62_V_minmax.iloc[:, 1]
H2_V_max['MCE62'] = H2_MCE62_V_minmax.iloc[:, 0]

# %% MCE71

# Max와 Min의 행, 열 index 찾기
H1_MCE71_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[12, :], H1_txt_dir+'\\'+MCE71_txt))  # tuple -> dataframe
H1_MCE71_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[12, :], H1_txt_dir+'\\'+MCE71_txt))

H2_MCE71_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[12, :], H2_txt_dir+'\\'+MCE71_txt))
H2_MCE71_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[12, :], H2_txt_dir+'\\'+MCE71_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE71_V_max = MatchV(H1_max.iloc[12, :], H1_MCE71_max_pos, V_txt_dir+'\\'+MCE71_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE71_V_min = MatchV(H1_min.iloc[12, :], H1_MCE71_min_pos, V_txt_dir+'\\'+MCE71_txt)

H2_MCE71_V_max = MatchV(H2_max.iloc[12, :], H2_MCE71_max_pos, V_txt_dir+'\\'+MCE71_txt)
H2_MCE71_V_min = MatchV(H2_min.iloc[12, :], H2_MCE71_min_pos, V_txt_dir+'\\'+MCE71_txt)

# max, min transpose하고 index 리셋
H1_MCE71_V_minmax = pd.concat([pd.Series(H1_MCE71_V_max), pd.Series(H1_MCE71_V_min)], axis=1)
H2_MCE71_V_minmax = pd.concat([pd.Series(H2_MCE71_V_max), pd.Series(H2_MCE71_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE71'] = H1_MCE71_V_minmax.iloc[:, 1]
H1_V_max['MCE71'] = H1_MCE71_V_minmax.iloc[:, 0]

H2_V_min['MCE71'] = H2_MCE71_V_minmax.iloc[:, 1]
H2_V_max['MCE71'] = H2_MCE71_V_minmax.iloc[:, 0]

# %% MCE72

# Max와 Min의 행, 열 index 찾기
H1_MCE72_max_pos = pd.DataFrame(FindPosition_max(H1_max.iloc[13, :], H1_txt_dir+'\\'+MCE72_txt))  # tuple -> dataframe
H1_MCE72_min_pos = pd.DataFrame(FindPosition_min(H1_min.iloc[13, :], H1_txt_dir+'\\'+MCE72_txt))

H2_MCE72_max_pos = pd.DataFrame(FindPosition_max(H2_max.iloc[13, :], H2_txt_dir+'\\'+MCE72_txt))
H2_MCE72_min_pos = pd.DataFrame(FindPosition_min(H2_min.iloc[13, :], H2_txt_dir+'\\'+MCE72_txt))

# Max, Min의 행, 열 index와 같은 V값 찾기
H1_MCE72_V_max = MatchV(H1_max.iloc[13, :], H1_MCE72_max_pos, V_txt_dir+'\\'+MCE72_txt)  # 최종적으로 계산에 사용할 축력
H1_MCE72_V_min = MatchV(H1_min.iloc[13, :], H1_MCE72_min_pos, V_txt_dir+'\\'+MCE72_txt)

H2_MCE72_V_max = MatchV(H2_max.iloc[13, :], H2_MCE72_max_pos, V_txt_dir+'\\'+MCE72_txt)
H2_MCE72_V_min = MatchV(H2_min.iloc[13, :], H2_MCE72_min_pos, V_txt_dir+'\\'+MCE72_txt)

# max, min transpose하고 index 리셋
H1_MCE72_V_minmax = pd.concat([pd.Series(H1_MCE72_V_max), pd.Series(H1_MCE72_V_min)], axis=1)
H2_MCE72_V_minmax = pd.concat([pd.Series(H2_MCE72_V_max), pd.Series(H2_MCE72_V_min)], axis=1)

# 추출된 Maximum, Minimum 값을 분리시켜 각각의 데이터프레임에 저장
H1_V_min['MCE72'] = H1_MCE72_V_minmax.iloc[:, 1]
H1_V_max['MCE72'] = H1_MCE72_V_minmax.iloc[:, 0]

H2_V_min['MCE72'] = H2_MCE72_V_minmax.iloc[:, 1]
H2_V_max['MCE72'] = H2_MCE72_V_minmax.iloc[:, 0]

#%% 수동입력한 길이 입력

# Wall 정보 load
SF_ongoing_modified = pd.read_excel(wall_raw_xlsx_dir+'\\'+wall_raw_xlsx,
                     sheet_name=SF_ongoing_modified_sheet, header=0)

length_modified = SF_ongoing_modified.loc[:,'Length']
length_modified.reset_index(inplace=True, drop=True)

SF_ongoing['Length'] = length_modified

# %% 계산~

# 공칭전단강도(Vn) 계산
Vs = SF_ongoing['Area(mm^2)']*SF_ongoing['Strength(MPa)_y']*0.8*2*SF_ongoing['Length']\
    /1000/SF_ongoing['H. Rebar Space']

# Vs = list(map(lambda x : x[0]*x[1]*0.8*2*x[2]/1000/x[3],\
#           SF_ongoing[['Area(mm^2)', 'Strength(MPa)_y', 'Length', 'H. Rebar Space']].values))

#%% 
Vc_simple = 1/6 * np.sqrt(SF_ongoing['Strength(MPa)_x']) * SF_ongoing['Length'] * SF_ongoing['Thickness'] / 1000

Vn_simple = Vc_simple + Vs

Vn_limit = 5 * Vn_simple

#%%   
Vn_H1_min = H1_V_min.apply(lambda x: (0.28*1*np.sqrt(SF_ongoing['Strength(MPa)_x'])\
                            *SF_ongoing['Thickness']*0.8*SF_ongoing['Length']/1000\
                            -x*0.8/4)+Vs)
                           
Vn_H1_max = H1_V_max.apply(lambda x: (0.28*1*np.sqrt(SF_ongoing['Strength(MPa)_x'])\
                            *SF_ongoing['Thickness']*0.8*SF_ongoing['Length']/1000\
                            -x*0.8/4)+Vs)
                           
Vn_H2_min = H2_V_min.apply(lambda x: (0.28*1*np.sqrt(SF_ongoing['Strength(MPa)_x'])\
                            *SF_ongoing['Thickness']*0.8*SF_ongoing['Length']/1000\
                            -x*0.8/4)+Vs)
                           
Vn_H2_max = H2_V_max.apply(lambda x: (0.28*1*np.sqrt(SF_ongoing['Strength(MPa)_x'])\
                            *SF_ongoing['Thickness']*0.8*SF_ongoing['Length']/1000\
                            -x*0.8/4)+Vs)  

H1_min = H1_min.T
H1_max = H1_max.T
H2_min = H2_min.T
H2_max = H2_max.T

H1_min.columns = Vn_H1_min.columns
H1_max.columns = Vn_H1_min.columns
H2_min.columns = Vn_H1_min.columns
H2_max.columns = Vn_H1_min.columns

H1_min.reset_index(drop=True, inplace=True)
H1_max.reset_index(drop=True, inplace=True)
H2_min.reset_index(drop=True, inplace=True)
H2_max.reset_index(drop=True, inplace=True)

# DCR 계산
DCR_H1_min = H1_min/Vn_H1_min
DCR_H1_max = H1_max/Vn_H1_max

DCR_H2_min = H2_min/Vn_H2_min
DCR_H2_max = H2_max/Vn_H2_max

#%%
# 보강 시 낮출 수 있는 DCR 최소값
DCR_H1_min_limit = H1_min.div(Vn_limit, axis=0)
DCR_H1_max_limit = H1_max.div(Vn_limit, axis=0)

DCR_H2_min_limit = H2_min.div(Vn_limit, axis=0)
DCR_H2_max_limit = H2_max.div(Vn_limit, axis=0)

#%%                   
# DCR 평균값 구하기
DCR_total = pd.concat([DCR_H1_min.iloc[:, 0:13].mean(axis=1), DCR_H1_max.iloc[:, 0:13].mean(axis=1),\
                       DCR_H1_min.iloc[:, 14:27].mean(axis=1), DCR_H1_max.iloc[:, 14:27].mean(axis=1),\
                       DCR_H2_min.iloc[:, 0:13].mean(axis=1), DCR_H2_max.iloc[:, 0:13].mean(axis=1),\
                       DCR_H2_min.iloc[:, 14:27].mean(axis=1), DCR_H2_max.iloc[:, 14:27].mean(axis=1)],\
                      keys=['H1 DE Min Average', 'H1 DE Max Average', 'H1 MCE Min Average',\
                            'H1 MCE Max Average', 'H2 DE Min Average', 'H2 DE Max Average',\
                            'H2 MCE Min Average', 'H2 MCE Max Average'], axis=1).abs()  # 절대값 씌우기
    
# Story 정보에서 height와 층이름만 뽑아내기
height = story_info.loc[:, 'Height(mm)']
story_name = story_info.loc[:, 'Floor Name']

# 최종 데이터프레임에 height열 추가
# SF_ongoing['Floor'] = SF_ongoing['Floor'].astype(str)
story_info.iloc[:,1] = story_info.iloc[:,1].astype(str)
SF_ongoing = pd.merge(SF_ongoing, story_info.iloc[:, [1,2]], how='left',\
                      left_on='Floor', right_on='Floor Name') 
    
DCR_total['Height'] = SF_ongoing['Height(mm)']

SF_output = pd.concat([section_name, DCR_total], axis=1)

DCR_total = DCR_total.dropna()
SF_output = SF_output.dropna()

#%% 조작용 코드
DCR_total = DCR_total[SF_output['Name'].str.contains('BW') == False]
DCR_total = DCR_total[(SF_output['Name'].str.contains('29') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= DCR_criteria)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('28') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('27') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('26') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW22_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW23') == False) | (SF_output['Name'].str.contains('PIT') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W3_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W11_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W13_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W2_') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW12A') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW14_') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW3_') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW4A') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW3_') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW0') == False)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW12_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW13') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW14') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
# DCR_total = DCR_total[(SF_output['Name'].str.contains('SW2A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW15_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW11A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW1_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW1A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= DCR_criteria)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW13A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW5_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW3A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW4') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW5A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
# DCR_total = DCR_total[(SF_output['Name'].str.contains('W1A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW15A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW11A_1_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW11A_2_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W1_11_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('DW11_1_8F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW23_5_9F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W3A_5_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('SW21_3_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
DCR_total = DCR_total[(SF_output['Name'].str.contains('W1_11_9F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]

# DCR_total = DCR_total[(SF_output.iloc[:,1:8].max(axis=1) <= 1.9)]


SF_output = SF_output[SF_output.iloc[:,0].str.contains('BW') == False]
SF_output = SF_output[(SF_output['Name'].str.contains('29') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= DCR_criteria)]
SF_output = SF_output[(SF_output['Name'].str.contains('28') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('27') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('26') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW22_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW23') == False) | (SF_output['Name'].str.contains('PIT') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('W3_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('W11_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('W13_') == False) | (SF_output['Name'].str.contains('PIT') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('W2_') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW12A') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW14_') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW3_') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW4A') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW3_') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW0') == False)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW12_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW13') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW14') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
# SF_output = SF_output[(SF_output['Name'].str.contains('SW2A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW15_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW11A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW1_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW1A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= DCR_criteria)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW13A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW5_') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW3A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW4') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW5A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
# SF_output = SF_output[(SF_output['Name'].str.contains('W1A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW15A') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW11A_1_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW11A_2_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('W1_11_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('DW11_1_8F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW23_5_9F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('W3A_5_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('SW21_3_PIT') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]
SF_output = SF_output[(SF_output['Name'].str.contains('W1_11_9F') == False) | (SF_output.iloc[:,1:9].max(axis=1) <= 1.05)]


# SF_output = SF_output[(SF_output.iloc[:,1:8].max(axis=1) <= 1.9)]

# DCR_total.reset_index(drop=True, inplace=True)
# SF_output.reset_index(drop=True, inplace=True)

                            
#%%
DCR_total_limit = pd.concat([DCR_H1_min_limit.iloc[:, 0:14].mean(axis=1), DCR_H1_max_limit.iloc[:, 0:14].mean(axis=1),\
                       DCR_H1_min_limit.iloc[:, 14:28].mean(axis=1), DCR_H1_max_limit.iloc[:, 14:28].mean(axis=1),\
                       DCR_H2_min_limit.iloc[:, 0:14].mean(axis=1), DCR_H2_max_limit.iloc[:, 0:14].mean(axis=1),\
                       DCR_H2_min_limit.iloc[:, 14:28].mean(axis=1), DCR_H2_max_limit.iloc[:, 14:28].mean(axis=1)],\
                      keys=['H1 DE Min Average', 'H1 DE Max Average', 'H1 MCE Min Average',\
                            'H1 MCE Max Average', 'H2 DE Min Average', 'H2 DE Max Average',\
                            'H2 MCE Min Average', 'H2 MCE Max Average'], axis=1).abs()  # 절대값 씌우기
    
DCR_total_limit['Height'] = SF_ongoing['Height(mm)']

DCR_total_limit = DCR_total_limit.dropna()

#%% 그래프 그리기

### H1 DE 그래프 ###
plt.figure(1, (4,5))
plt.xlim(0, graph_xlim)
plt.scatter(DCR_total['H1 DE Min Average'], DCR_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(DCR_total['H1 DE Max Average'], DCR_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(height[::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(wall_raw_xlsx_dir + '\\' + 'SF_H1_DE')


### H2 DE 그래프 ###
plt.figure(2, (4,5))
plt.xlim(0, graph_xlim)
plt.scatter(DCR_total['H2 DE Min Average'], DCR_total['Height'], color = 'k', s=1) # s=1 : point size
plt.scatter(DCR_total['H2 DE Max Average'], DCR_total['Height'], color = 'k', s=1)

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(height[::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(wall_raw_xlsx_dir + '\\' + 'SF_H2_DE')


### H1 MCE 그래프 ###
plt.figure(3, figsize=(4,5))
plt.xlim(0, graph_xlim)
plt.scatter(DCR_total['H1 MCE Min Average'], DCR_total['Height'], color = 'k', s=1)
plt.scatter(DCR_total['H1 MCE Max Average'], DCR_total['Height'], color = 'k', s=1)

plt.yticks(height[::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(wall_raw_xlsx_dir + '\\' + 'SF_H1_MCE')



### H2 MCE 그래프 ###
plt.figure(4, figsize=(4,5))
plt.xlim(0, graph_xlim)
plt.scatter(DCR_total['H2 MCE Min Average'], DCR_total['Height'], color = 'k', s=1)
plt.scatter(DCR_total['H2 MCE Max Average'], DCR_total['Height'], color = 'k', s=1)

plt.yticks(height[::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')

plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(wall_raw_xlsx_dir + '\\' + 'SF_H2_MCE')

#%% 결과 확인할때 필요한 함수
# 기준 점 넘는 node 찾는 함수

def ErrorSection(value_min_list, value_max_list, max_criteria):
    
    section_list = SF_output['Name']
    error_section = []
    error_value = []
    
    max_value = value_min_list.where(value_min_list > value_max_list, value_max_list)
    
    for error_max_value, error_value_index in zip(max_value, value_min_list.index):
        
        if (error_max_value >= max_criteria):            
            error_section.append(section_list.loc[error_value_index].strip())
            error_value.append(error_max_value) 
            
    return error_section, error_value

#%% 기준 넘는 section 출력 (데이터프레임 열어서 확인)

error_H1_DE_section = ErrorSection(DCR_total['H1 DE Min Average'], DCR_total['H1 DE Max Average'], DCR_criteria)
error_H1_MCE_section = ErrorSection(DCR_total['H1 MCE Min Average'], DCR_total['H1 MCE Max Average'], DCR_criteria)
error_H2_DE_section = ErrorSection(DCR_total['H2 DE Min Average'], DCR_total['H2 DE Max Average'], DCR_criteria)
error_H2_MCE_section = ErrorSection(DCR_total['H2 MCE Min Average'], DCR_total['H2 MCE Max Average'], DCR_criteria)

error_H1_DE_section = pd.DataFrame({'Section' : error_H1_DE_section[0], 'Value' : error_H1_DE_section[1]})
error_H1_MCE_section = pd.DataFrame({'Section' : error_H1_MCE_section[0], 'Value' : error_H1_MCE_section[1]})
error_H2_DE_section = pd.DataFrame({'Section' : error_H2_DE_section[0], 'Value' : error_H2_DE_section[1]})
error_H2_MCE_section = pd.DataFrame({'Section' : error_H2_MCE_section[0], 'Value' : error_H2_MCE_section[1]})


#%% 부재별 DCR 및 NG여부 출력(부록 형식)
# DCR_total에 층, 배근, NG여부 합쳐서 

DCR_output = pd.concat([SF_ongoing.iloc[:,[0,1,13]].dropna(), \
                        round(DCR_total.iloc[:,[2,3,6,7]].max(axis=1), 4), round(DCR_total_limit.iloc[:,[2,3,6,7]].max(axis=1), 4)], axis=1)

DCR_output['Result'] = np.where(DCR_output.iloc[:,3] >= DCR_criteria, 'NG', 'OK')

DCR_output.columns = ['부재명', '층', '수평 배근', '최대 DCR', '보강시 낮출 수 있는 DCR limit','결과']

#%% 전체 코드 runtime 측정

time_end = timeit.default_timer()
time_run = (time_end-time_start)/60
print('total time = %0.7f min' %(time_run))