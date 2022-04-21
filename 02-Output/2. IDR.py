import pandas as pd
import os
import matplotlib.pyplot as plt

#%% 사용자 입력

# 층간변위비 허용기준
min_criteria_DE = -0.015
max_criteria_DE = 0.015
min_criteria_MCE = -0.02
max_criteria_MCE = 0.02

# y축 층수 간격
story_gap = 2

### Gage data/Gage result/Node Coordinate data/Story Info 불러오기
data_path = r'C:\Users\hwlee\Desktop\Python\내진성능설계' # data 폴더 경로
analysis_result = 'Analysis Result' # 해석 결과 파일 이름(확장자명 제외)

# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
input_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0.xlsx'

# drift_rsa_xlsx_sheet = 'Drift RSA'

#%% Analysis Result 불러오기
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if (analysis_result in file_name) and ('~$' not in file_name):
        to_load_list.append(file_name)

# Gage data
IDR_result_data = pd.DataFrame()
for i in to_load_list:
    IDR_result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Drift Output', skiprows=[0, 2], header=0, usecols=[0, 1, 3, 5, 6]) # usecols로 원하는 열만 불러오기
    IDR_result_data = IDR_result_data.append(IDR_result_data_temp)

IDR_result_data = IDR_result_data.sort_values(by=['Load Case', 'Drift ID', 'Step Type']) # 지진파 순서가 섞여있을 때 sort

# Story Info data
story_info_xlsx_sheet = 'Story Data'
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, skiprows=3, usecols=[0, 1, 2], keep_default_na=False)
story_info.columns = ['Index', 'Story Name', 'Height(mm)']
story_name = story_info.loc[:, 'Story Name']

#%% Drift Name에서 story, direction 뽑아내기
drift_name = IDR_result_data['Drift Name']

story = []
direction = []
position = []
for i in drift_name:
    i = i.strip()  # drift_name 앞뒤에 있는 blank 제거

    if i.count('_') == 2:
        story.append(i.split('_')[0])
        direction.append(i.split('_')[-1])
        position.append(i.split('_')[1].split('_')[0])
    else:
        story.append(None)
        direction.append(None)

# Load Case에서 지진파 이름만 뽑아서 다시 naming
load_striped = []        
for i in IDR_result_data['Load Case']:
    load_striped.append(i.strip().split(' ')[-1])
    
IDR_result_data['Load Case'] = load_striped
    

IDR_result_data.reset_index(inplace=True, drop=True)
IDR_result_data = pd.concat([pd.Series(story, name='Name'), pd.Series(direction, name='Direction'), pd.Series(position, name='Position'), IDR_result_data], axis=1)

#%% IDR값(방향에 따른)
### 지진파별 평균

# 지진파 이름 자동 생성
load_name_list = IDR_result_data['Load Case'].drop_duplicates()
load_name_list = [x for x in load_name_list if ('DE' in x) or ('MCE' in x)]
load_name_list.sort()

# 각 지진파들로 변수 생성 후, 값 대입
for load_name in load_name_list:
    globals()['IDR_x_max_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                  (IDR_result_data['Direction'] == 'X') &\
                                                                  (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift']\
                                                                  .agg(**{'X Max avg':'mean'}).groupby('Name').max()
    
    globals()['IDR_x_min_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                  (IDR_result_data['Direction'] == 'X') &\
                                                                  (IDR_result_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                  .agg(**{'X Min avg':'mean'}).groupby('Name').min()
        
    globals()['IDR_y_max_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                  (IDR_result_data['Direction'] == 'Y') &\
                                                                  (IDR_result_data['Step Type'] == 'Max')].groupby(['Name'])['Drift']\
                                                                  .agg(**{'X Max avg':'mean'}).groupby('Name').max()
    
    globals()['IDR_y_min_{}_avg'.format(load_name)] = IDR_result_data[(IDR_result_data['Load Case'] == '{}'.format(load_name)) &\
                                                                  (IDR_result_data['Direction'] == 'Y') &\
                                                                  (IDR_result_data['Step Type'] == 'Min')].groupby(['Name'])['Drift']\
                                                                  .agg(**{'X Min avg':'mean'}).groupby('Name').min()
        
    globals()['IDR_x_max_{}_avg'.format(load_name)].reset_index(inplace=True)
    globals()['IDR_x_min_{}_avg'.format(load_name)].reset_index(inplace=True)
    globals()['IDR_y_max_{}_avg'.format(load_name)].reset_index(inplace=True)
    globals()['IDR_y_min_{}_avg'.format(load_name)].reset_index(inplace=True)
    
# # Story 정렬하기
# story_name_window = globals()['IDR_x_max_{}_avg'.format(load_name_list[0])]['Name']
# story_name_window_reordered = [x for x in story_name[::-1].tolist() \
#                                if x in story_name_window.tolist()]  # story name를 reference로 해서 정렬

# 정렬된 Story에 따라 IDR값도 정렬
for load_name in load_name_list:   
    globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
    globals()['IDR_x_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
    
    globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_x_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
    globals()['IDR_x_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
    
    globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_max_{}_avg'.format(load_name)]['Name'], story_name[::-1])
    globals()['IDR_y_max_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
    
    globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'] = pd.Categorical(globals()['IDR_y_min_{}_avg'.format(load_name)]['Name'], story_name[::-1])
    globals()['IDR_y_min_{}_avg'.format(load_name)].sort_values('Name', inplace=True)
    
#%% IDR값(방향에 따른) 전체 평균

IDR_x_max_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                       (IDR_result_data['Direction'] == 'X') &\
                                       (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
IDR_x_min_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                       (IDR_result_data['Direction'] == 'X') &\
                                       (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
IDR_y_max_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                       (IDR_result_data['Direction'] == 'Y') &\
                                       (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
IDR_y_min_DE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('DE')) &\
                                       (IDR_result_data['Direction'] == 'Y') &\
                                       (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()

IDR_x_max_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                       (IDR_result_data['Direction'] == 'X') &\
                                       (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
IDR_x_min_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                       (IDR_result_data['Direction'] == 'X') &\
                                       (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()
IDR_y_max_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                       (IDR_result_data['Direction'] == 'Y') &\
                                       (IDR_result_data['Step Type'] == 'Max')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').max()
IDR_y_min_MCE_avg = IDR_result_data[(IDR_result_data['Load Case'].str.contains('MCE')) &\
                                       (IDR_result_data['Direction'] == 'Y') &\
                                       (IDR_result_data['Step Type'] == 'Min')].groupby(['Name', 'Position'])['Drift'].agg(**{'X Max avg':'mean'}).groupby('Name').min()



# 정렬된 Story에 따라 IDR값도 정렬
IDR_x_max_DE_avg = pd.merge(IDR_x_max_DE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_x_min_DE_avg = pd.merge(IDR_x_min_DE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_y_max_DE_avg = pd.merge(IDR_y_max_DE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_y_min_DE_avg = pd.merge(IDR_y_min_DE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_x_max_MCE_avg = pd.merge(IDR_x_max_MCE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_x_min_MCE_avg = pd.merge(IDR_x_min_MCE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_y_max_MCE_avg = pd.merge(IDR_y_max_MCE_avg, story_info, how='left', left_on='Name', right_on='Story Name')
IDR_y_min_MCE_avg = pd.merge(IDR_y_min_MCE_avg, story_info, how='left', left_on='Name', right_on='Story Name')

IDR_x_max_DE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_x_min_DE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_y_max_DE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_y_min_DE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_x_max_MCE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_x_min_MCE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_y_max_MCE_avg.sort_values(by='Height(mm)', inplace=True)
IDR_y_min_MCE_avg.sort_values(by='Height(mm)', inplace=True)

IDR_x_max_DE_avg.reset_index(inplace=True)
IDR_x_min_DE_avg.reset_index(inplace=True)
IDR_y_max_DE_avg.reset_index(inplace=True)
IDR_y_min_DE_avg.reset_index(inplace=True)
IDR_x_max_MCE_avg.reset_index(inplace=True)
IDR_x_min_MCE_avg.reset_index(inplace=True)
IDR_y_max_MCE_avg.reset_index(inplace=True)
IDR_y_min_MCE_avg.reset_index(inplace=True)

IDR_x_max_DE_avg = IDR_x_max_DE_avg.iloc[:,[3,1]]
IDR_x_min_DE_avg = IDR_x_min_DE_avg.iloc[:,[3,1]]
IDR_y_max_DE_avg = IDR_x_max_DE_avg.iloc[:,[3,1]]
IDR_y_min_DE_avg = IDR_x_min_DE_avg.iloc[:,[3,1]]
IDR_x_max_MCE_avg = IDR_x_max_MCE_avg.iloc[:,[3,1]]
IDR_x_min_MCE_avg = IDR_x_min_MCE_avg.iloc[:,[3,1]]
IDR_y_max_MCE_avg = IDR_x_max_MCE_avg.iloc[:,[3,1]]
IDR_y_min_MCE_avg = IDR_x_min_MCE_avg.iloc[:,[3,1]]  
    
#%% IDR값(위치에 따른)

# IDR_2_max = IDR_result_data[(IDR_result_data['Step Type'] == 'Max') & (IDR_result_data['Position'] == '2')]
# IDR_2_min = IDR_result_data[(IDR_result_data['Step Type'] == 'Min') & (IDR_result_data['Position'] == '2')]

# IDR_5_max = IDR_result_data[(IDR_result_data['Step Type'] == 'Max') & (IDR_result_data['Position'] == '5')]
# IDR_5_min = IDR_result_data[(IDR_result_data['Step Type'] == 'Min') & (IDR_result_data['Position'] == '5')]

# IDR_7_max = IDR_result_data[(IDR_result_data['Step Type'] == 'Max') & (IDR_result_data['Position'] == '7')]
# IDR_7_min = IDR_result_data[(IDR_result_data['Step Type'] == 'Min') & (IDR_result_data['Position'] == '7')]

# IDR_11_max = IDR_result_data[(IDR_result_data['Step Type'] == 'Max') & (IDR_result_data['Position'] == '11')]
# IDR_11_min = IDR_result_data[(IDR_result_data['Step Type'] == 'Min') & (IDR_result_data['Position'] == '11')]    

#%% 그래프
### 방향에 따른 그래프

### H1 DE 그래프
plt.figure(1, figsize=(4, 7), dpi=150)
plt.xlim(-0.025, 0.025)

# 지진파별 plot
for load_name in load_name_list:
    if 'DE' in load_name:
        plt.plot(globals()['IDR_x_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
        plt.plot(globals()['IDR_x_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], linewidth=0.7)
        
# 평균 plot      
plt.plot(IDR_x_max_DE_avg.iloc[:,-1], IDR_x_max_DE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
plt.plot(IDR_x_min_DE_avg.iloc[:,-1], IDR_x_min_DE_avg.iloc[:,0], color='k', linewidth=2)

# reference line 그려서 허용치 나타내기
plt.axvline(x=min_criteria_DE, color='r', linestyle='--', label='LS')
plt.axvline(x=max_criteria_DE, color='r', linestyle='--')

# 기타
plt.yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
plt.grid(linestyle='-.')
plt.xlabel('Interstory Drift Ratios(m/m)')
plt.ylabel('Story')
plt.legend(loc=4, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'IDR_H1_DE')

### H2 DE 그래프
plt.figure(2, figsize=(4, 7), dpi=150)
plt.xlim(-0.025, 0.025)

# 지진파별 plot
for load_name in load_name_list:
    if 'DE' in load_name:
        plt.plot(globals()['IDR_y_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
        plt.plot(globals()['IDR_y_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], linewidth=0.7)

# 평균 plot
plt.plot(IDR_y_max_DE_avg.iloc[:,-1], IDR_y_max_DE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
plt.plot(IDR_y_min_DE_avg.iloc[:,-1], IDR_y_min_DE_avg.iloc[:,0], color='k', linewidth=2)

# reference line 그려서 허용치 나타내기
plt.axvline(x=min_criteria_DE, color='r', linestyle='--', label='LS')
plt.axvline(x=max_criteria_DE, color='r', linestyle='--')

# 기타
plt.yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
plt.grid(linestyle='-.')
plt.xlabel('Interstory Drift Ratios(m/m)')
plt.ylabel('Story')
plt.legend(loc=4, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'IDR_H2_DE')

### H1 MCE 그래프
plt.figure(3, figsize=(4, 7), dpi=150)
plt.xlim(-0.025, 0.025)

# 지진파별 plot
for load_name in load_name_list:
    if 'MCE' in load_name:
        plt.plot(globals()['IDR_x_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
        plt.plot(globals()['IDR_x_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], linewidth=0.7)
        
# 평균 plot      
plt.plot(IDR_x_max_MCE_avg.iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
plt.plot(IDR_x_min_MCE_avg.iloc[:,-1], IDR_x_max_MCE_avg.iloc[:,0], color='k', linewidth=2)

# reference line 그려서 허용치 나타내기
plt.axvline(x=min_criteria_MCE, color='r', linestyle='--', label='CP')
plt.axvline(x=max_criteria_MCE, color='r', linestyle='--')

# 기타
plt.yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
plt.grid(linestyle='-.')
plt.xlabel('Interstory Drift Ratios(m/m)')
plt.ylabel('Story')
plt.legend(loc=4, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'IDR_H1_DE')

# H2 MCE 그래프
plt.figure(4, figsize=(4, 7), dpi=150)
plt.xlim(-0.025, 0.025)

# 지진파별 plot
for load_name in load_name_list:
    if 'MCE' in load_name:
        plt.plot(globals()['IDR_y_max_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], label='{}'.format(load_name), linewidth=0.7)
        plt.plot(globals()['IDR_y_min_{}_avg'.format(load_name)].iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], linewidth=0.7)

# 평균 plot
plt.plot(IDR_y_max_MCE_avg.iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], color='k', label='Average', linewidth=2)
plt.plot(IDR_y_min_MCE_avg.iloc[:,-1], IDR_y_max_MCE_avg.iloc[:,0], color='k', linewidth=2)

# reference line 그려서 허용치 나타내기
plt.axvline(x=min_criteria_MCE, color='r', linestyle='--', label='CP')
plt.axvline(x=max_criteria_MCE, color='r', linestyle='--')

# 기타
plt.yticks(story_name_window_reordered[::story_gap], story_name_window_reordered[::story_gap])
plt.grid(linestyle='-.')
plt.xlabel('Interstory Drift Ratios(m/m)')
plt.ylabel('Story')
plt.legend(loc=4, fontsize=8)
# plt.title('Shear Wall Rotation')

plt.tight_layout()
plt.savefig(data_path + '\\' + 'IDR_H2_DE')

plt.show()
