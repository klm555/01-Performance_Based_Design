# -*- coding: utf-8 -*-
"""
###  update(2022.01.11)  ###
1. 층별 section 출력되어 나옴.

###  update(2021.10.22)  ###
1. 경로 설정 편하게 다 앞으로 몰아놨음.
2. drift도 naming돼서 출력됨.
3. Base Section도 출력되어 나옴.
"""

import pandas as pd
import numpy as np
import openpyxl

# Input 경로 설정
section_info_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
section_info_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

# Output 경로 설정
name_output_xlsx_dir = section_info_xlsx_dir
name_output_xlsx = 'Naming Output Sheets.xlsx'

# Drift의 위치, 방향 지정
position_list = [2, 5, 7, 11]
direction_list = ['X', 'Y']

#%% section, frame 이름 만들기 위한 정보 load
section_info_xlsx_sheet = 'Wall Naming' # section naming 관련된 정보만 들어있는 시트
beam_info_xlsx_sheet = 'Beam Naming'
story_info_xlsx_sheet = 'Story Data' # 층 정보 sheet
drift_info_xlsx_sheet = 'ETC' # Drift 정보 sheet

section_info = pd.read_excel(section_info_xlsx_dir + '\\' + section_info_xlsx, sheet_name = section_info_xlsx_sheet, skiprows = 3)

# Section Naming 시트에서 불러온 데이터 열 분리
wall_name = section_info.iloc[:,0].values  # .values : series -> numpy array
story_from = section_info.iloc[:,1].tolist() # .tolist() : series -> list
story_to = section_info.iloc[:,2].tolist()
amount = section_info.iloc[:,3].values

# Beam에 대해서도 똑같이...
beam_info = pd.read_excel(section_info_xlsx_dir + '\\' + section_info_xlsx, sheet_name = beam_info_xlsx_sheet, skiprows = 3)

beam_name = beam_info.iloc[:,0].values
beam_story_from = beam_info.iloc[:,1].tolist()
beam_story_to = beam_info.iloc[:,2].tolist()
beam_amount = beam_info.iloc[:,3].values

#%% story 정보 load (section info에서 불러온 story_from, story_to를 연동시키기 위함)
story_info = pd.read_excel(section_info_xlsx_dir + '\\' + section_info_xlsx, sheet_name = story_info_xlsx_sheet, skiprows = 3)

# 불러온 Story info 시트에서 층 이름(story name)만 불러오기
story_name = story_info.iloc[:,1].tolist() # 마찬가지로 series -> list
story_name = story_name[::-1] # 배열이 내가 원하는 방향과 반대로 되어있어서, 리스트 거꾸로만들었음

#%% Section 이름 뽑기

# for문으로 section naming에 사용할 섹션 이름(section_name_output) 뽑기
section_name_output = [] # 결과로 나올 section_name_output 리스트 미리 정의

for wall_name_parameter, amount_parameter, story_from_parameter, story_to_parameter in zip(wall_name, amount, story_from, story_to):  # for 문에 조건 여러개 달고싶을 때는 zip으로 묶어서~ 
    story_from_index = story_name.index(story_from_parameter)  # story_from이 문자열이라 story_from을 사용해서 slicing이 안되기 때문에(내 지식선에서) .index로 story_from의 index만 뽑음
    story_to_index = story_name.index(story_to_parameter)  # 마찬가지로 story_to의 index만 뽑음
    story_window = story_name[story_from_index : story_to_index+1]  # 내가 원하는 층 구간(story_from부터 story_to까지)만 뽑아서 리스트로 만들기
    for i in np.arange(1, amount_parameter+1):  # (벽체 개수(amount))에 맞게 numbering하기 위해 1,2,3,4...amount[i]개의 배열을 만듦. 첫 시작을 1로 안하면 index 시작은 0이 default값이기 때문에 1씩 더해줌
        for current_story_name in story_window:
            if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                current_story_name = str(current_story_name)
            else:
                pass
            
            section_name_output.append(wall_name_parameter + '_' + i.astype(str) + '_' + current_story_name)  # 반복될때마다 생성되는 section 이름을 .append를 이용하여 리스트의 끝에 하나씩 쌓아줌. i값은 숫자라 .astype(str)로 string으로 바꿔줌

# Base section 추가하기
section_name_output.append('Base')

# 각 층 전단력 확인을 위한 각 층 section 추가하기
for i in story_name[1:len(story_name)]:
    section_name_output.append(i + '_Shear')

#%% Frame 이름 뽑기

# Wall Frame 이름 뽑기
frame_wall_name_output = []

for wall_name_parameter, amount_parameter, story_from_parameter, story_to_parameter in zip(wall_name, amount, story_from, story_to):  # for 문에 조건 여러개 달고싶을 때는 zip으로 묶어서~ 
    story_from_index = story_name.index(story_from_parameter)  
    story_to_index = story_name.index(story_to_parameter) 
    story_window = story_name[story_from_index:story_to_index+1]  
    
    for i in np.arange(1, amount_parameter+1):  
        frame_wall_name_output.append(wall_name_parameter + '_' +i.astype(str))
        
# Beam Frame 이름 뽑기
frame_beam_name_output = []

for i, j, k, l in zip(beam_name, beam_amount, beam_story_from, beam_story_to):
    beam_story_from_index = story_name.index(k)
    beam_story_to_index = story_name.index(l)
    beam_story_window = story_name[beam_story_from_index : beam_story_to_index+1]
    
    for i in np.arange(1, j+1):
        frame_beam_name_output.append(i + '_' + )
        
#%% Constraints 이름 뽑기


#%% Drift 이름 뽑기

# 시작층, 끝층의 index 뽑고, 그 구간의 index 리스트 만듦
drift_from_index = []
drift_to_index = []

for i, j in zip(story_from, story_to):
    drift_from_index.append(story_name.index(i))
    drift_to_index.append(story_name.index(j))

# Drift로 만들 층 list
drift_story_window = story_name[min(drift_from_index) : max(drift_to_index)+1]

# Drift 이름 만들기
drift_name_output = []

for position in position_list:
    for direction in direction_list:
        for current_story_name in drift_story_window:
            if isinstance(current_story_name, str) == False:  # 층이름이 int인 경우, 이름조합을 위해 str로 바꿈
                current_story_name = str(current_story_name)
            drift_name_output.append(current_story_name + '_' + str(int(position)) + '_' + direction)
                
#%% 출력

name_output = pd.DataFrame(({'Frame(Beam) Name': pd.Series(frame_beam_name_output),\
                             'Frame(Wall) Name': pd.Series(frame_wall_name_output),\
                             'Constraints': ???
                             'Section Name': pd.Series(section_name_output),\
                             'Drift Name': pd.Series(drift_name_output)}))

name_output.to_excel(name_output_xlsx_dir+ '\\'+ name_output_xlsx, sheet_name = 'Name List', index = False)

