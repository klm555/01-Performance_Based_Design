"""
###  update(2022.01.11)  ###
1. [H1 DE], [H2 DE], [H1 MCE], [H2 MCE] 순
"""

import numpy as np
import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
from io import StringIO  # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import matplotlib.pyplot as plt

# 전체 코드 runtime 측정
time_start = timeit.default_timer()

#%% 파일 경로 지정

# Wall input 정보 (xlsx)
wall_raw_xlsx_dir = r'D:\이형우\내진성능평가\광명 4R\101'
wall_raw_xlsx = 'Input Sheets(101)'

story_info_xlsx_sheet = 'Story Info'

# 해석 output 파일 (txt)
H1_txt_dir = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101\H1'
H2_txt_dir = r'D:\이형우\내진성능평가\광명 4R\해석 결과\101\H2'

# story shear y축 층 간격 정하기
story_yticks = 2 #ex) 3층 간격->3

# 그래프 size(가로길이,세로길이)

figsize_x, figsize_y = 4,6 
# 원하는 부재명
input_wall_name = 'CW9_1' # 필요한 부재명, 모든 부재 = all

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

#%% H1, H2의 텍스트 파일 load (element별 height이 포함된 행만 load)
H1_DE11_raw_txt = H1_txt_dir+'\\%s' %file_name

# Section name열 추출
section_name = pd.read_csv(H1_DE11_raw_txt, skiprows=7, error_bad_lines=False, header=None).iloc[:, -1]
section_name.reset_index(drop=True, inplace=True)
section_name.name = 'Name'

# if input_wall_name == 'all':
    
# Section으로 지정한 각 층을 추출하기 (이름의 형태에 따라 달라질수있음!!!)
floor = []
for i in section_name:
    i = i.strip()  # section_name 앞에 있는 blank 제거
    
    if i.startswith(input_wall_name):
        floor.append(i.split('_')[2])
        
    else:
        floor.append(None)
  
# 추출한 층을 각 전단력 dataframe의 column 인덱스로 지정
H1_max.columns = floor
H1_min.columns = floor
H2_max.columns = floor
H2_min.columns = floor

# 각 지진파 이름을 dataframe의 row 인덱스로 지정
load_name = []
for i in H1_file_names:
    load_name.append(i.split('.')[0])

H1_max.index = load_name
H1_min.index = load_name
H2_max.index = load_name
H2_min.index = load_name

# 해당 부재의 전단력을 제외한 나머지 부재들의 전단력 drop
H1_max.drop([None], axis=1, inplace=True)
H1_min.drop([None], axis=1, inplace=True)
H2_max.drop([None], axis=1, inplace=True)
H2_min.drop([None], axis=1, inplace=True)

# Min과 Max의 절대값 중 큰 값
story_shear_H1 = H1_max.where(H1_max.abs() > H1_min.abs(), H1_min.abs()).T
story_shear_H2 = H2_max.where(H2_max.abs() > H2_min.abs(), H2_min.abs()).T

#%% Story 정보 load
# story_info = pd.read_excel(wall_raw_xlsx_dir+'\\'+wall_raw_xlsx,
#                            sheet_name=story_info_xlsx_sheet, keep_default_na=False)

# # Story 정보에서 층이름만 뽑아내기
# story_name = story_info.loc[:, 'Floor Name']
# story_name = story_name[::-1]  # 층 이름 재배열
# story_name.reset_index(drop=True, inplace=True)

# # Story 정렬하기
# story_name_window = story_shear_H1.index
# story_name_window[0] = 
# story_name_window_reordered = [x for x in story_name.tolist() \
#                                if x in story_name_window.tolist()]  # story name를 reference로 해서 정렬


#%% 부재별 SF 그래프 그리기
### H1_DE
plt.figure(1, figsize=(figsize_x, figsize_y), dpi=150)

# 지진파별 plot
for i in range(14):
    plt.plot(story_shear_H1.iloc[:,i], range(story_shear_H1.shape[0]), label=load_name[i], linewidth=0.7)
    
# 평균 plot
plt.plot(story_shear_H1.iloc[:,0:14].mean(axis=1), range(story_shear_H1.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(story_shear_H1.shape[0])[::story_yticks], story_shear_H1.index[::story_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Shear Force(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
plt.title(input_wall_name.split('_')[0] + ' (X-Dir.)')
# plt.savefig(data_path + '\\' + 'Story_SF_H1_DE')


# H2_DE
plt.figure(2, figsize=(figsize_x, figsize_y), dpi=150)

# 지진파별 plot
for i in range(14):
    plt.plot(story_shear_H2.iloc[:,i], range(story_shear_H2.shape[0]), label=load_name[i], linewidth=0.7)
    
# 평균 plot
plt.plot(story_shear_H2.iloc[:,0:14].mean(axis=1), range(story_shear_H2.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(story_shear_H2.shape[0])[::story_yticks], story_shear_H2.index[::story_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Shear Force(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
plt.title(input_wall_name.split('_')[0] + ' (Y-Dir.)')
# plt.savefig(data_path + '\\' + 'Story_SF_H2_DE')

### H1_MCE
plt.figure(3, figsize=(figsize_x, figsize_y), dpi=150)

# 지진파별 plot
for i in range(14):
    plt.plot(story_shear_H1.iloc[:,i+14], range(story_shear_H1.shape[0]), label=load_name[i+14], linewidth=0.7)
    
# 평균 plot
plt.plot(story_shear_H1.iloc[:,14:28].mean(axis=1), range(story_shear_H1.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(story_shear_H1.shape[0])[::story_yticks], story_shear_H1.index[::story_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Shear Force(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
plt.title(input_wall_name.split('_')[0] + ' (X-Dir.)')
# plt.savefig(data_path + '\\' + 'Story_SF_H1_MCE')

# H2_MCE
plt.figure(4, figsize=(figsize_x, figsize_y), dpi=150)

# 지진파별 plot
for i in range(14):
    plt.plot(story_shear_H2.iloc[:,i+14], range(story_shear_H2.shape[0]), label=load_name[i+14], linewidth=0.7)
    
# 평균 plot
plt.plot(story_shear_H2.iloc[:,14:28].mean(axis=1), range(story_shear_H2.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(story_shear_H2.shape[0])[::story_yticks], story_shear_H2.index[::story_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Shear Force(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
plt.title(input_wall_name.split('_')[0] + ' (Y-Dir.)')
# plt.savefig(data_path + '\\' + 'Story_SF_H2_MCE')


plt.show()

#%% 전체 코드 runtime 측정

time_end = timeit.default_timer()
time_run = (time_end-time_start)/60
print('total time = %0.7f min' %(time_run))