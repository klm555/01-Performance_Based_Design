import numpy as np
import pandas as pd
import os
from collections import deque  # Double-ended Queue : 자료의 앞, 뒤 양 방향에서 자료를 추가하거나 제거가능
from io import StringIO  # 파일처럼 취급되는 문자열 객체 생성(메모리 낭비 down)
import matplotlib.pyplot as plt

#%% 사용자 입력

# x, y축 limit값(그래프 y축 크기) 정하기
base_shear_ylim = 50000 #kN
story_shear_xlim = 70000 #kN

# story shear y축 층 간격 정하기
story_shear_yticks = 2 #ex) 3층 간격->3

### Analysis Result 정보불러오기
data_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\103' # data 폴더 경로
data_xlsx = 'Analysis Result' # 파일명


# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

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
                               sheet_name='Structure Section Forces', skiprows=2, usecols=[0,3,5,6,7])
    shear_force_data = shear_force_data.append(shear_force_data_temp)
    
shear_force_data.columns = ['Name', 'Load Case', 'Step Type', 'H1(kN)', 'H2(kN)']

# 필요없는 전단력 제거(층전단력)
if shear_force_data['Name'].str.contains('_').all() == False: # 이 줄은 없어도 될거같은디..
    shear_force_data = shear_force_data[shear_force_data['Name'].str.contains('_') == False]
  
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

DE_load_name_list = [x for x in load_name_list if 'DE' in x] # base shear로 사용할 지진파 개수 산정을 위함
MCE_load_name_list = [x for x in load_name_list if 'MCE' in x]

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

# Min, Max 중 최대값
shear_force_H1_max = shear_force_H1_abs.groupby([[i//2 for i in range(0,56)]], axis=1).max()
shear_force_H2_max = shear_force_H2_abs.groupby([[i//2 for i in range(0,56)]], axis=1).max()

shear_force_H1_max.columns = seismic_load_name_list
shear_force_H2_max.columns = seismic_load_name_list

shear_force_H1_max.index = shear_force_data['Name'].drop_duplicates()
shear_force_H2_max.index = shear_force_data['Name'].drop_duplicates()

#%% Story 정보 load

# Story 정보에서 층이름만 뽑아내기
story_info_xlsx_sheet = 'Story Data'
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
story_name = story_info.loc[:, 'Story Name']

# Story 정렬하기
# story_name_window = story_shear_H1.index
# story_name_window[0] = 
# story_name_window_reordered = [x for x in story_name.tolist() \
#                                 if x in story_name_window.tolist()]  # story name를 reference로 해서 정렬

#%% Base Shear 그래프 그리기
# Base Shear
base_shear_H1 = shear_force_H1_max[shear_force_H1_max.index.str.contains('base', case=False)]
base_shear_H2 = shear_force_H2_max[shear_force_H2_max.index.str.contains('base', case=False)]

# H1_DE
plt.figure(1)
plt.ylim(0, base_shear_ylim)

plt.bar(range(14), base_shear_H1.iloc[0, 0:14], color='darkblue', edgecolor='k', label = 'Max. Base Shear')
plt.axhline(y= base_shear_H1.iloc[0, 0:14].mean(), color='r', linestyle='-', label='Average')
plt.xticks(range(14), range(1,15))
# plt.xticks(range(14), load_name[0:14], fontsize=8.5)

plt.xlabel('Ground Motion No.')
plt.ylabel('Base Shear(kN)')
plt.legend(loc = 2)

# plt.savefig(data_path + '\\' + 'Base_SF_H1_DE')

# H2_DE
plt.figure(2)
plt.ylim(0, base_shear_ylim)

plt.bar(range(14), base_shear_H2.iloc[0, 0:14], color='darkblue', edgecolor='k', label = 'Max. Base Shear')
plt.axhline(y= base_shear_H2.iloc[0, 0:14].mean(), color='r', linestyle='-', label='Average')
plt.xticks(range(14), range(1,15))
# plt.xticks(range(14), load_name[0:14], fontsize=8.5)

plt.xlabel('Ground Motion No.')
plt.ylabel('Base Shear(kN)')
plt.legend(loc = 2)

# plt.savefig(data_path + '\\' + 'Base_SF_H2_DE')

# H1_MCE
plt.figure(3)
plt.ylim(0, base_shear_ylim)

plt.bar(range(14), base_shear_H1.iloc[0, 14:28], color='darkblue', edgecolor='k', label = 'Max. Base Shear')
plt.axhline(y= base_shear_H1.iloc[0, 14:28].mean(), color='r', linestyle='-', label='Average')
plt.xticks(range(14), range(1,15))
# plt.xticks(range(14), load_name[0:14], fontsize=8.5)

plt.xlabel('Ground Motion No.')
plt.ylabel('Base Shear(kN)')
plt.legend(loc = 2)

# plt.savefig(data_path + '\\' + 'Base_SF_H1_MCE')

# H2_MCE
plt.figure(4)
plt.ylim(0, base_shear_ylim)

plt.bar(range(14), base_shear_H2.iloc[0, 14:28], color='darkblue', edgecolor='k', label = 'Max. Base Shear')
plt.axhline(y= base_shear_H2.iloc[0, 14:28].mean(), color='r', linestyle='-', label='Average')
plt.xticks(range(14), range(1,15))
# plt.xticks(range(14), load_name[0:14], fontsize=8.5)

plt.xlabel('Ground Motion No.')
plt.ylabel('Base Shear(kN)')
plt.legend(loc = 2)

# plt.savefig(data_path + '\\' + 'Base_SF_H2_MCE')

plt.show()

print('base_shear_avg(H1_DE) =', round(base_shear_H1.iloc[0, 0:14].mean(), 2),
      '\nbase_shear_avg(H2_DE) =', round(base_shear_H2.iloc[0, 0:14].mean(), 2),
      '\nbase_shear_avg(H1_MCE) =', round(base_shear_H1.iloc[0, 14:28].mean(), 2),
      '\nbase_shear_avg(H2_MCE) =', round(base_shear_H2.iloc[0, 14:28].mean(), 2))

#%% ***temp***
temp_story_list = shear_force_data['Name'].drop_duplicates().tolist()
del temp_story_list[-1] # 층 list에서 맨 마지막 층 지우기
temp_story_list.insert(0, 'Base') # 층 list에서 맨 첫 층 추가

shear_force_H1_max = shear_force_H1_max.apply(np.roll, shift=1) # 맨 마지막 1열을 첫번쨰 열로 굴리기
shear_force_H2_max = shear_force_H2_max.apply(np.roll, shift=1)

shear_force_H1_max.index = temp_story_list
shear_force_H2_max.index = temp_story_list

#%% Story Shear 그래프 그리기
### H1_DE
plt.figure(5, dpi=150)
plt.xlim(0, story_shear_xlim)

# 지진파별 plot
for i in range(14):
    plt.plot(shear_force_H1_max.iloc[:,i], range(shear_force_H1_max.shape[0]), label=seismic_load_name_list[i], linewidth=0.7)
    
# 평균 plot
plt.plot(shear_force_H1_max.iloc[:,0:14].mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(shear_force_H1_max.shape[0])[::story_shear_yticks], shear_force_H1_max.index[::story_shear_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Story Shear(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'Story_SF_H1_DE')

# H2_DE
plt.figure(6, dpi=150)
plt.xlim(0, story_shear_xlim)

# 지진파별 plot
for i in range(14):
    plt.plot(shear_force_H2_max.iloc[:,i], range(shear_force_H2_max.shape[0]), label=seismic_load_name_list[i], linewidth=0.7)
    
# 평균 plot
plt.plot(shear_force_H2_max.iloc[:,0:14].mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(shear_force_H2_max.shape[0])[::story_shear_yticks], shear_force_H2_max.index[::story_shear_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Story Shear(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'Story_SF_H2_DE')

### H1_MCE
plt.figure(7, dpi=150)
plt.xlim(0, story_shear_xlim)

# 지진파별 plot
for i in range(14):
    plt.plot(shear_force_H1_max.iloc[:,i+14], range(shear_force_H1_max.shape[0]), label=seismic_load_name_list[i+14], linewidth=0.7)
    
# 평균 plot
plt.plot(shear_force_H1_max.iloc[:,14:28].mean(axis=1), range(shear_force_H1_max.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(shear_force_H1_max.shape[0])[::story_shear_yticks], shear_force_H1_max.index[::story_shear_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Story Shear(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'Story_SF_H1_MCE')

# H2_MCE
plt.figure(8, dpi=150)
plt.xlim(0, story_shear_xlim)

# 지진파별 plot
for i in range(14):
    plt.plot(shear_force_H2_max.iloc[:,i+14], range(shear_force_H2_max.shape[0]), label=seismic_load_name_list[i+14], linewidth=0.7)
    
# 평균 plot
plt.plot(shear_force_H2_max.iloc[:,14:28].mean(axis=1), range(shear_force_H2_max.shape[0]), color='k', label='Average', linewidth=2)

plt.yticks(range(shear_force_H2_max.shape[0])[::story_shear_yticks], shear_force_H2_max.index[::story_shear_yticks], fontsize=8.5)
# plt.xticks(range(14), range(1,15))

# 기타
plt.grid(linestyle='-.')
plt.xlabel('Story Shear(kN)')
plt.ylabel('Story')
plt.legend(loc=1, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'Story_SF_H2_MCE')

plt.show()