import numpy as np
from openpyxl import load_workbook
import pandas as pd
import math
import re

#%% 사용자 경로 설정

# Input 경로 설정
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

# Output 경로 설정
output_xlsx_dir = input_raw_xlsx_dir
output_xlsx = 'Wall,Beam Output Sheets.xlsx'

#%% 파일 load
wall_raw_xlsx_sheet = 'Wall Properties'
beam_raw_xlsx_sheet = 'C.Beam Properties'
story_info_xlsx_sheet = 'Story Data'
naming_criteria_xlsx_sheet = 'ETC'
wall_naming_xlsx_sheet = 'Wall Naming'
beam_naming_xlsx_sheet = 'Beam Naming'

# Wall 정보 load
wall = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                     sheet_name=wall_raw_xlsx_sheet, skiprows=3, header=0)

wall = wall.iloc[:,np.r_[0:11, 21,22]]
wall.columns = ['Name', 'Story(from)', 'Story(to)', 'Thickness', 'Vertical Rebar(DXX)',\
                'V. Rebar Space', 'Horizontal Rebar(DXX)', 'H. Rebar Space', 'Type', 'Length', 'Element length', 'Fibers(Concrete)', 'Fibers(Rebar)']

wall = wall.dropna(axis=0, how='all')
wall.reset_index(inplace=True, drop=True)
wall = wall.fillna(method='ffill')

# Beam 정보 load
beam = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                     sheet_name=beam_raw_xlsx_sheet, skiprows=3, header=0)

beam = beam.iloc[:,0:20]
beam.columns = ['Name', 'Story(from)', 'Story(to)', 'Length(mm)', 'b(mm)',\
                'h(mm)', 'Cover Thickness(mm)', 'Type', '배근', '내진상세 여부',\
                'Main Rebar(DXX)', 'Stirrup Rebar(DXX)', 'X-Bracing Rebar', 'Top(1)', 'Top(2)',\
                'EA(Stirrup)', 'Spacing(Stirrup)', 'EA(Diagonal)', 'Degree(Diagonal)', 'D(mm)']

beam = beam.dropna(axis=0, how='all')
beam.reset_index(inplace=True, drop=True)
beam = beam.fillna(method='ffill')

# 구분 조건 load
naming_criteria = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                                 sheet_name=naming_criteria_xlsx_sheet, skiprows=3, header=0)

# Story 정보 load
story_info = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                           sheet_name=story_info_xlsx_sheet, skiprows=3, keep_default_na=False)

# Story 정보에서 층이름만 뽑아내기
story_info_xlsx_sheet = 'Story Data'
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
story_name = story_info.loc[:, 'Story Name']
story_name = story_name[::-1]  # 층 이름 재배열
story_name.reset_index(drop=True, inplace=True)

# 벽체 개수 load
num_of_wall = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                            sheet_name=wall_naming_xlsx_sheet, skiprows=3)

num_of_wall = num_of_wall.drop(['from', 'to'], axis=1)
num_of_wall.columns = ['Name', 'EA']

# 보 개수 load
num_of_beam = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                            sheet_name=beam_naming_xlsx_sheet, skiprows=3)

num_of_beam = num_of_beam.drop(['from', 'to'], axis=1)
num_of_beam.columns = ['Name', 'EA']

#%% 1. Wall
#%% 벽체 이름 설정할 때 필요한 함수들

# 층, 철근 나누는 함수 (12F~15F, D10@300)

def str_div(temp_list):
    first = []
    second = []
    
    for i in temp_list:
        if '~' in i:
            first.append(i.split('~')[0])
            second.append(i.split('~')[1])
        elif '-' in i:
            second.append(i.split('-')[0])
            first.append(i.split('-')[1])
        else:
            first.append(i)
            second.append(i)
    
    first = pd.Series(first).str.strip()
    second = pd.Series(second).str.strip()
    
    return first, second

# 층, 철근 나누는 함수 (12F~15F, D10@300)

def rebar_div(temp_list1, temp_list2):
    first = []
    second = []
    third = []
    
    for i, j in zip(temp_list1, temp_list2):
        if isinstance(i, str) : # string인 경우
            if '@' in i:
                first.append(i.split('@')[0].strip())
                second.append(i.split('@')[1])
                third.append(np.nan)
            elif '-' in i:
                third.append(i.split('-')[0])
                first.append(i.split('-')[1].strip())
                second.append(np.nan)
            else: 
                first.append(i.strip())
                second.append(j)
                third.append(np.nan)
        else: # string 아닌 경우
            first.append(i)
            second.append(j)
            third.append(np.nan)
    
    # first = pd.Series(first).str.strip()
    # second = pd.Series(second).str.strip()
    
    return first, second, third

# 철근 지름 앞의 D 떼주는 함수 (D10...)

def str_extract(sth_str):
    result = int(re.findall(r'[0-9]+', sth_str)[0])
    
    return result

#%% 데이터베이스
steel_geometry_database = naming_criteria.iloc[:,[0,1,2]].dropna()
steel_geometry_database.columns = ['Name', 'Diameter(mm)', 'Area(mm^2)']

new_steel_geometry_name = []

for i in steel_geometry_database['Name']:
    if isinstance(i, int):
        new_steel_geometry_name.append(i)
    else:
        new_steel_geometry_name.append(str_extract(i))

steel_geometry_database['Name'] = new_steel_geometry_name

#%% 불러온 wall 정보 정리하기

# 글자가 합쳐져 있을 경우 글자 나누기 (12F~15F, D10@300)
# 층 나누기

if wall['Story(to)'].isnull().any() == True:
    wall['Story(to)'] = str_div(wall['Story(from)'])[1]
    wall['Story(from)'] = str_div(wall['Story(from)'])[0]
else: pass

# V. Rebar 나누기
wall['V. Rebar EA'] = 0 # 'V. Rebar EA 열 만들기
whynotworking = rebar_div(wall['Vertical Rebar(DXX)'], wall['V. Rebar Space'])
wall['Vertical Rebar(DXX)'] = whynotworking[0]
wall['V. Rebar Space'] = whynotworking[1]
wall['V. Rebar EA'] = whynotworking[2]

# H. Rebar 나누기
# wall['Horizontal Rebar(DXX)'] = rebar_div(wall['Horizontal Rebar(DXX)'], wall['H. Rebar Space'])[0].copy()

whynotworking2 = rebar_div(wall['Horizontal Rebar(DXX)'], wall['H. Rebar Space']) # 아니 이거 왜 안되는거야, 어이없네 ㅎㅎ
wall['Horizontal Rebar(DXX)'] = whynotworking2[0]
wall['H. Rebar Space'] = whynotworking2[1]

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

# Rebar Space 데이터값 모두 float로 바꿔주기
v_rebar_spacing_float = []
h_rebar_spacing_float = []
v_rebar_ea_float = []

for i, j, k in zip(wall['V. Rebar Space'], wall['H. Rebar Space'], wall['V. Rebar EA']):
    
    if not isinstance(i, float):
        v_rebar_spacing_float.append(float(i))
        
    else: v_rebar_spacing_float.append(i)
        
    if not isinstance(j, float):
        h_rebar_spacing_float.append(float(j))
        
    else: h_rebar_spacing_float.append(j)
    
    if not isinstance(k, float):
        v_rebar_ea_float.append(float(k))
        
    else: v_rebar_ea_float.append(k)
    
wall['V. Rebar Space'] = v_rebar_spacing_float
wall['H. Rebar Space'] = h_rebar_spacing_float
wall['V. Rebar EA'] = v_rebar_ea_float

#%% 이름 구분 조건 load & 정리

# 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
naming_criteria_1_index = []
naming_criteria_2_index = []

for i, j in zip(naming_criteria.iloc[:,5].dropna(), naming_criteria.iloc[:,6].dropna()):
    naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
    naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))

### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
# 구분 조건에 해당하는 콘크리트 강도 재정렬
naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,7].dropna()], axis=1)

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
wall_info = wall.copy()  # input sheet에서 나온 properties
wall_info.reset_index(drop=True, inplace=True)  # ?빼도되나?

name_output = []  # new names
property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
wall_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output

count = 1000
count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기

for i, j in zip(num_of_wall['Name'], num_of_wall['EA']):
    
    for k in range(1,j+1):

        for current_wall_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_wall_info_index\
                    in zip(wall['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, wall_info.index):

            if i == current_wall_name:
                
                
                
                for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                    if p != q:
                        for s in range(p, q+1):

                            count_list.append(count + s)
                            
                            name_output.append(current_wall_name + '_' + str(k) + '_' + str(story_name[s]))
                            
                            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                            wall_info_output.append(wall_info.iloc[current_wall_info_index])
                            
                    else:
                        count_list.append(count + q)
                        
                        name_output.append(current_wall_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                        
                        property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                        wall_info_output.append(wall_info.iloc[current_wall_info_index])  
                        
        count += 1000
        
wall_info_output = pd.DataFrame(wall_info_output)
wall_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..

wall_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당

# 중간결과
if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 wall_info를 바로 출력
    wall_ongoing = wall_info
else:
    wall_ongoing = pd.concat([pd.Series(name_output, name='Name'), wall_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties

wall_ongoing = wall_ongoing.sort_values(by=['Count'])
wall_ongoing.reset_index(inplace=True, drop=True)

# 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
wall_output = wall_ongoing.iloc[:,[0,10,4,15,9,5,6,14,7,8,12,13]]  

# 철근지름에 다시 D붙이기
wall_output['Vertical Rebar(DXX)'] = 'D' + wall_output['Vertical Rebar(DXX)'].astype(str)
wall_output['Horizontal Rebar(DXX)'] = 'D' + wall_output['Horizontal Rebar(DXX)'].astype(str)

#%% 2. Beam
#%% 불러온 Beam 정보 정리

# 글자가 합쳐져 있을 경우 글자 나누기 (12F~15F, D10@300)
# 층 나누기

if beam['Story(to)'].isnull().any() == True:
    beam['Story(to)'] = str_div(beam['Story(from)'])[1]
    beam['Story(from)'] = str_div(beam['Story(from)'])[0]
else: pass

# 철근의 앞에붙은 D 떼어주기
new_m_rebar = []
new_s_rebar = []

for i in beam['Main Rebar(DXX)']:
    if isinstance(i, int):
        new_m_rebar.append(i)
    else:
        new_m_rebar.append(str_extract(i))
        
for j in beam['Stirrup Rebar(DXX)']:
    if isinstance(j, int):
        new_s_rebar.append(j)
    else:
        new_s_rebar.append(str_extract(j))
        
beam['Main Rebar(DXX)'] = new_m_rebar
beam['Stirrup Rebar(DXX)'] = new_s_rebar

#%% 이름 구분 조건 load & 정리

# 층 구분 조건에  story_name의 index 매칭시켜서 새로 열 만들기
naming_criteria_1_index = []
naming_criteria_2_index = []

for i, j in zip(naming_criteria.iloc[:,8].dropna(), naming_criteria.iloc[:,9].dropna()):
    naming_criteria_1_index.append(pd.Index(story_name).get_loc(i))
    naming_criteria_2_index.append(pd.Index(story_name).get_loc(j))

### 구분 조건이 층 순서에 상관없이 작동되게 재정렬
# 구분 조건에 해당하는 콘크리트 강도 재정렬
naming_criteria_property = pd.concat([pd.Series(naming_criteria_1_index, name='Story(from) Index'), naming_criteria.iloc[:,10].dropna()], axis=1)

naming_criteria_property['Story(from) Index'] = pd.Categorical(naming_criteria_property['Story(from) Index'], naming_criteria_1_index.sort())
naming_criteria_property.sort_values('Story(from) Index', inplace=True)
naming_criteria_property.reset_index(inplace=True)

# 구분 조건 재정렬
naming_criteria_1_index.sort()
naming_criteria_2_index.sort()

#%% 시작층, 끝층 정리

naming_from_index = []
naming_to_index = []

for naming_from, naming_to in zip(beam['Story(from)'], beam['Story(to)']):
    if isinstance(naming_from, str) == False:
        naming_from = str(naming_from)
    if isinstance(naming_to, str) == False:
        naming_from = str(naming_from)
        
    naming_from_index.append(pd.Index(story_name).get_loc(naming_from))
    naming_to_index.append(pd.Index(story_name).get_loc(naming_to))

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
beam_info = beam.copy()  # input sheet에서 나온 properties
beam_info.reset_index(drop=True, inplace=True)  # ?빼도되나?

name_output = []  # new names
property_output = []  # 이름 구분 조건에 따라 할당되는 properties를 새로운 부재명에 맞게 다시 정리한 output
beam_info_output = []  # input sheet에서 나온 properties를 새로운 부재명에 맞게 다시 정리한 output

count = 1000
count_list = [] # 벽체이름을 오름차순으로 바꾸기 위한 index 만들기

for i, j in zip(num_of_beam['Name'], num_of_beam['EA']):
    
    for k in range(1,j+1):

        for current_beam_name, current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list, current_beam_info_index\
                    in zip(beam['Name'], naming_from_index_list, naming_to_index_list, naming_criteria_property_index_list, beam_info.index):

            if i == current_beam_name:
                
                
                
                for p, q, r in zip(current_naming_from_index_list, current_naming_to_index_list, current_naming_criteria_property_index_list):
                    if p != q:
                        for s in range(p, q+1):

                            count_list.append(count + s)
                            
                            name_output.append(current_beam_name + '_' + str(k) + '_' + str(story_name[s]))
                            
                            property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                            beam_info_output.append(beam_info.iloc[current_beam_info_index])
                            
                    else:
                        count_list.append(count + q)
                        
                        name_output.append(current_beam_name + '_' + str(k) + '_' + str(story_name[q]))  # 시작과 끝층이 같으면 둘 중 한 층만 표기
                        
                        property_output.append(naming_criteria_property.iloc[:,-1][r])  # 각 이름에 맞게 property 할당 (index의 index 사용하였음)
                        beam_info_output.append(beam_info.iloc[current_beam_info_index])  
                        
        count += 1000
        
beam_info_output = pd.DataFrame(beam_info_output)
beam_info_output.reset_index(drop=True, inplace=True)  # 왜인지는 모르겠는데 index가 이상해져서..

beam_info_output['Concrete Strength(CXX)'] = property_output  # 이름 구분 조건에 따른 property를 중간결과물에 재할당

# 중간결과
if (len(name_output) == 0) or (len(property_output) == 0):  # 구분 조건없이 을 경우는 beam_info를 바로 출력
    beam_ongoing = beam_info
else:
    beam_ongoing = pd.concat([pd.Series(name_output, name='Name'), beam_info_output, pd.Series(count_list, name='Count')], axis = 1)  # 중간결과물 : 부재명 변경, 콘크리트 강도 추가, 부재명과 콘크리트 강도에 따른 properties

beam_ongoing = beam_ongoing.sort_values(by=['Count'])
beam_ongoing.reset_index(inplace=True, drop=True)

# 최종 sheet에 미리 넣을 수 있는 것들도 넣어놓기
beam_output = beam_ongoing.iloc[:,[0,4,5,6,20,21,8,9,10,11,12,13,14,15,16,17,18,19]]  

# 철근지름에 다시 D붙이기
beam_output['Main Rebar(DXX)'] = 'D' + beam_output['Main Rebar(DXX)'].astype(str)
beam_output['Stirrup Rebar(DXX)'] = 'D' + beam_output['Stirrup Rebar(DXX)'].astype(str)



#%% Output 뽑기
with pd.ExcelWriter(output_xlsx_dir+ '\\'+ output_xlsx) as writer:  
    beam_output.to_excel(writer, sheet_name='Output_C.Beam Properties', index = False)
    wall_output.to_excel(writer, sheet_name='Output_Wall_Properties', index = False)