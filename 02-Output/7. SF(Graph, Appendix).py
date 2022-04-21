"""
###  update(2022.01.24)  ###
1. [H1 DE], [H1 MCE], [H2 DE], [H2 MCE] 순
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

#%% 사용자 입력

# Wall input 정보 (xlsx)
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

wall_result_raw_xlsx_sheet = 'Results_Wall'
story_info_xlsx_sheet = 'Story Data'

# story shear y축 층 간격 정하기
story_yticks = 4 #ex) 3층 간격->3

graph_xlim = 3 # x축 limit

# DCR 기준선 (넘는 부재들 error_*_section에서 확인)
DCR_criteria = 1.0

#%% 파일 load

# Wall 정보 load
wall_result = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                     sheet_name=wall_result_raw_xlsx_sheet, skiprows=3, header=0)

wall_result = wall_result.iloc[:, [0, 25, 27, 29, 31]]
wall_result.columns = ['Name', 'DE_H1', 'DE_H2', 'MCE_H1', 'MCE_H2']
wall_result = wall_result.dropna()
wall_result.reset_index(inplace=True, drop=True)

# Story 정보 load
story_info = pd.read_excel(input_raw_xlsx_dir+'\\'+input_raw_xlsx,
                           sheet_name=story_info_xlsx_sheet, skiprows=3, keep_default_na=False)

story_info.columns = ['Index', 'Floor', 'Height(mm)', 'Story Height(mm)', 'EA']

# Story 정보에서 층이름만 뽑아내기
story_name = story_info.iloc[:, 1]
story_name.reset_index(drop=True, inplace=True)

#%% ***조작용 코드
wall_name_to_delete = ['84A-W1_1','84A-W3_1_40F'] 
# 지우고싶은 층들을 대괄호 안에 입력(벽 이름만 입력하면 벽 전체 다 없어짐, 벽+층 이름 입력하면 특정 층의 벽만 없어짐)

for i in wall_name_to_delete:
    wall_result = wall_result[wall_result['Name'].str.contains(i) == False]

#%% 벽체 해당하는 층 높이 할당
floor = []
for i in wall_result['Name']:
    floor.append(i.split('_')[-1])

wall_result['Floor'] = floor

wall_result_output = pd.merge(wall_result, story_info.iloc[:,[1,2]], how='left')

#%% 그래프

### H1 DE 그래프 ###
plt.figure(1, (4,5))
plt.xlim(0, graph_xlim)
plt.scatter(wall_result_output['DE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')
plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(input_raw_xlsx_dir + '\\' + 'SF_H1_DE')


### H2 DE 그래프 ###
plt.figure(2, (4,5))
plt.xlim(0, graph_xlim)
plt.scatter(wall_result_output['MCE_H1'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')
plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(input_raw_xlsx_dir + '\\' + 'SF_H2_MCE')


### H1 MCE 그래프 ###
plt.figure(3, figsize=(4,5))
plt.xlim(0, graph_xlim)
plt.scatter(wall_result_output['DE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')
plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(input_raw_xlsx_dir + '\\' + 'SF_H1_DE')


### H2 MCE 그래프 ###
plt.figure(4, figsize=(4,5))
plt.xlim(0, graph_xlim)
plt.scatter(wall_result_output['MCE_H2'], wall_result_output['Height(mm)'], color = 'k', s=1) # s=1 : point size

# height값에 대응되는 층 이름으로 y축 눈금 작성
plt.yticks(story_info['Height(mm)'][::-story_yticks], story_name[::-story_yticks])

plt.axvline(x= DCR_criteria, color='r', linestyle='--')
plt.grid(linestyle='-.')
plt.xlabel('D/C Ratios')
plt.ylabel('Story')
# plt.title('Shear Wall Rotation')
plt.savefig(input_raw_xlsx_dir + '\\' + 'SF_H2_MCE')

plt.show()

#%% 전체 코드 runtime 측정

time_end = timeit.default_timer()
time_run = (time_end-time_start)/60
print('total time = %0.7f min' %(time_run))

#%% Appendix (to be continued...)