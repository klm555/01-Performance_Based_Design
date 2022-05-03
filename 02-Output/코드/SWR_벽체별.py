import pandas as pd
import os
import matplotlib.pyplot as plt

### Initial Setting
# 지진파 개수
DE_num = 14
MCE_num = 14

# 허용기준
min_criteria_DE = -0.002
max_criteria_DE = 0.002
min_criteria_MCE = -0.004/1.2
max_criteria_MCE = 0.004/1.2

### ECHO File/Gage data/Gage result/Node Coordinate data/Story Info 불러오기
data_path = r'C:\Users\khpark\Desktop\21-RM-513 광명 4R구역\퍼폼 모델링\110동\Output\data'

# ECHO File
ECHO_path = data_path + '\\' + 'ECHO_110.xlsx'
ECHO_data = pd.read_excel(ECHO_path)

# Story 정보 경로 설정
input_raw_xlsx_dir = r'C:\Users\khpark\Desktop\21-RM-513 광명 4R구역\퍼폼 모델링\110동\Output\data'
input_raw_xlsx = '110D_Input Sheets(Ver.2).xlsx'
story_info_xlsx_sheet = 'Story Info'

# Analysis Result 불러오기
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if 'Analysis Result' in file_name and '~$' not in file_name:
        to_load_list.append(file_name)

# Gage data
SWR_gage_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Gage Data - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 3, 7, 9, 11, 13]) # usecols로 원하는 열만 불러오기

# Gage result data
SWR_result_data = pd.DataFrame()
for i in to_load_list:
    SWR_result_data_temp = pd.read_excel(data_path + '\\' + i,
                               sheet_name='Gage Results - Wall Type', skiprows=[0, 2], header=0, usecols=[0, 3, 5, 7, 9, 11])
    SWR_result_data = SWR_result_data.append(SWR_result_data_temp)

# Node Coord data
Node_coord_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])

# Story Info data
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name=story_info_xlsx_sheet, keep_default_na=False)
story_height = story_info.loc[:, 'Height(mm)']
story_name = story_info.loc[:, 'Floor Name']


### Gage data에서 Element ID, I-Node ID 불러와서 v좌표 match하기
SWR_gage_data = SWR_gage_data[['Element ID', 'I-Node ID']]; gage_num = len(SWR_gage_data) # gage 개수 얻기
Node_v_coord_data = Node_coord_data[['Node ID', 'V']]

# I-Node의 v좌표 match해서 추가
SWR_gage_data = SWR_gage_data.join(Node_coord_data.set_index('Node ID')[['H1', 'H2', 'V']], on='I-Node ID')



### 2. Find start, end lines of nodes/ROWT/shear wall
## start line
for i in range(len(ECHO_data)):
    if ECHO_data.iloc[i, 1] == "*NODES":
        node_start_line = i+5
    elif ECHO_data.iloc[i, 4] == "ROTW":
        gage_start_line = i+9

wall_start_line = []
wall_end_line = []
for i in range(len(ECHO_data)):
    if ECHO_data.iloc[i, 1] == "*ELEMENTS" and ECHO_data.iloc[i+1, 4] == "WALL":
        wall_start_line.append(i+11)


## end line
for i in range(node_start_line, len(ECHO_data)):
    if pd.isna(ECHO_data.iloc[i, 1]) == True:
        node_end_line = i
        break

for i in range(gage_start_line, len(ECHO_data)):
    if pd.isna(ECHO_data.iloc[i, 4]) == True:
        gage_end_line = i
        break

for i in wall_start_line:
    for j in range(i, len(ECHO_data)):
        if pd.isna(ECHO_data.iloc[j, 1]) == True:
            wall_end_line.append(j)
            break                                                                                                        # wall line은 여러개



### 3. Define node_data/gage_data/wall_data frame
## node_data frame
node_data = ECHO_data.iloc[node_start_line:node_end_line, 1:5]; node_data.columns=['node_num', 'h1', 'h2', 'v']

## gage_data frame
gage_data = ECHO_data.iloc[gage_start_line:gage_end_line, 1:6]; gage_data.columns=['num', 'i', 'j', 'k', 'l']
gage_data = gage_data.join(node_data.set_index('node_num')['h1'], on='i'); gage_data.rename({'h1' : 'i_x'}, axis=1, inplace=True)     # python에서 vlookup 기능 사용하기, key 와 column이 서로 다른 문제때문에 이걸로 했음.
gage_data = gage_data.join(node_data.set_index('node_num')['h1'], on='j'); gage_data.columns.values[6] = 'j_x'
gage_data = gage_data.join(node_data.set_index('node_num')['h2'], on='i'); gage_data.columns.values[7] = 'i_y'           # 나중에 section 좌표와 비교해서 매칭할 때 필요
gage_data = gage_data.join(node_data.set_index('node_num')['v'], on='i'); gage_data.columns.values[8] = 'i_z'            # 나중에 층 구할 때 필요
gage_data['i_x-j_x'] = gage_data.apply(lambda x: x['i_x'] - x['j_x'], axis=1)

gage_data['direction'] = gage_data.apply(lambda x: 'h2'
                                         if (abs(x['i_x-j_x'])<=15) # 15mm 의 차이는 같은 것으로 간주
                                         else 'h1', axis=1)

## wall_data frame
wall_data = pd.DataFrame()
for i in range(len(wall_start_line)):                                                                                    # wall line은 여러개일 수 있으므로 for문을 통해 구성, temp를 append함으로써 한층씩 쌓아가기
    temp = ECHO_data.iloc[wall_start_line[i]:wall_end_line[i]:2, [2, 3, 8]]
    wall_data = wall_data.append(temp)
wall_data.columns=['i', 'j', 'wall_name']

wall_data = wall_data.join(node_data.set_index('node_num')['h1'], on='i'); wall_data.rename({'h1' : 'i_x'}, axis=1, inplace=True)
wall_data = wall_data.join(node_data.set_index('node_num')['h1'], on='j'); wall_data.columns.values[4] = 'j_x'
wall_data = wall_data.join(node_data.set_index('node_num')['v'], on='i'); wall_data.columns.values[5] = 'i_z'            # 나중에 층 구할 때 필요
wall_data['i_x-j_x'] = wall_data.apply(lambda x: x['i_x'] - x['j_x'], axis=1)

wall_data['direction'] = wall_data.apply(lambda x: 'h2'
                                         if (abs(x['i_x-j_x'])<=15) # 15mm 차이는 같은 것으로 간주
                                         else 'h1', axis=1)


###4. Match gage name with wall name
gage_data = gage_data.join(wall_data.set_index(['i', 'direction'])['wall_name'], on=['i', 'direction']); gage_data.columns.values[11] = 'gage_name_full'  # gage와 wall을 연결하는 작업, i를 기준으로
# gage orientation 이 다른 경우 대비하기 위해 추가
gage_data = gage_data.join(wall_data.set_index(['i', 'direction'])['wall_name'], on=['j', 'direction']); gage_data.columns.values[12] = 'gage_name_full_j'  # gage와 wall을 연결하는 작업, j를 기준으로

for i in range(len(gage_data)):
    if pd.isnull(gage_data.iloc[i, 11]):
        gage_data.iloc[i, 11] = gage_data.iloc[i, 12]

gage_data = gage_data.iloc[:, 0:12]


## Define gage_story_data
gage_story_data = pd.DataFrame()
for i in range(len(gage_data)):
    for j in range(len(story_info)):
        if gage_data.iloc[i, 8] == story_info.iloc[j, 2]:
            gage_story_data = gage_story_data.append(pd.Series(story_info.iloc[j, 1]), ignore_index=True)
            break

gage_data.reset_index(drop=True, inplace=True)
gage_story_data.reset_index(drop=True, inplace=True)
gage_data = pd.concat([gage_data, gage_story_data], axis=1)
gage_data.columns.values[12] = 'Story'

# 임시 방편으로 gage_data를 띄어놓음. 나중에 꼭 해결해야 할 key error
# 빠진 이름 있을 경우 아래 코드로 수동으로 이름 채워넣기
#gage_data = gage_data.fillna('CW6_')

gage_data['gage_name'] = gage_data['gage_name_full'].apply(lambda x: x.split('_')[0])


### gage_data와 SWR_result_data join 해서 벽체 이름 얻기
SWR_result_data = SWR_result_data.join(gage_data.set_index('num')['gage_name'], on=['Element ID'])



### SWR_total data 만들기(기존 코드에서 gage name 까지 추가)

SWR_max = SWR_result_data[(SWR_result_data['Step Type'] == 'Max') & (SWR_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
SWR_max_gagename = SWR_result_data[(SWR_result_data['Step Type'] == 'Max') & (SWR_result_data['Performance Level'] == 1)][['gage_name']].values # dataframe을 array로
SWR_max = SWR_max.reshape(gage_num, 28, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
SWR_max_gagename = SWR_max_gagename.reshape(gage_num, 28, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
SWR_max_gagename = pd.DataFrame(SWR_max_gagename) # array를 다시 dataframe으로


SWR_min = SWR_result_data[(SWR_result_data['Step Type'] == 'Min') & (SWR_result_data['Performance Level'] == 1)][['Rotation']].values
SWR_min_gagename = SWR_result_data[(SWR_result_data['Step Type'] == 'Min') & (SWR_result_data['Performance Level'] == 1)][['gage_name']].values
SWR_min = SWR_min.reshape(gage_num, 28, order='F')
SWR_min_gagename = SWR_min_gagename.reshape(gage_num, 28, order='F')
SWR_min = pd.DataFrame(SWR_min)
SWR_min_gagename = pd.DataFrame(SWR_min_gagename)

SWR_total = pd.concat([SWR_gage_data['V'], SWR_max_gagename.iloc[:,0], SWR_max, SWR_min], axis=1)

SWR_total.columns = ['Height', 'gage_name',
                     'DE11_max', 'DE12_max', 'DE21_max', 'DE22_max', 'DE31_max', 'DE32_max', 'DE41_max', 'DE42_max', 'DE51_max', 'DE52_max', 'DE61_max', 'DE62_max', 'DE71_max', 'DE72_max',
                     'MCE11_max', 'MCE12_max', 'MCE21_max', 'MCE22_max', 'MCE31_max', 'MCE32_max', 'MCE41_max', 'MCE42_max', 'MCE51_max', 'MCE52_max', 'MCE61_max', 'MCE62_max', 'MCE71_max', 'MCE72_max',
                     'DE11_min', 'DE12_min', 'DE21_min', 'DE22_min', 'DE31_min', 'DE32_min', 'DE41_min', 'DE42_min', 'DE51_min', 'DE52_min', 'DE61_min', 'DE62_min', 'DE71_min', 'DE72_min',
                     'MCE11_min', 'MCE12_min', 'MCE21_min', 'MCE22_min', 'MCE31_min', 'MCE32_min', 'MCE41_min', 'MCE42_min', 'MCE51_min', 'MCE52_min', 'MCE61_min', 'MCE62_min', 'MCE71_min', 'MCE72_min']

SWR_total = SWR_total.sort_values(['gage_name', 'Height']) # 그래프가 깔끔하게 나오기 위해 벽체마다 다시 height 기준으로 재정렬

### SWR_avg_data 만들기
DE_max_avg = SWR_total.iloc[:, 2:DE_num+2].mean(axis=1) # 2를 더해준 건 앞에 Height와 gage_name이 추가되었기 때문
MCE_max_avg = SWR_total.iloc[:, DE_num+2 : DE_num + MCE_num+2].mean(axis=1)
DE_min_avg = SWR_total.iloc[:, DE_num+MCE_num+2 : 2*DE_num+MCE_num+2].mean(axis=1)
MCE_min_avg = SWR_total.iloc[:, 2*DE_num+MCE_num+2 : 2*DE_num + 2*MCE_num+2].mean(axis=1)
SWR_avg_total = pd.concat([SWR_total['Height'], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
SWR_avg_total.columns = ['Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']




### Graph Plotting
# Legend에 쓸 지진파 이름 설정
DE_name = ['DE11', 'DE12', 'DE21', 'DE22', 'DE31', 'DE32', 'DE41', 'DE42', 'DE51', 'DE52', 'DE61', 'DE62', 'DE71', 'DE72']
MCE_name = ['MCE11', 'MCE12', 'MCE21', 'MCE22', 'MCE31', 'MCE32', 'MCE41', 'MCE42', 'MCE51', 'MCE52', 'MCE61', 'MCE62', 'MCE71', 'MCE72']


# 벽체가 바뀌는 지점 찾기
wall_name_change = []
for i in range(len(SWR_total)-1):
    if SWR_total.iloc[i, 1] != SWR_total.iloc[i + 1, 1]:
        wall_name_change.append(i)


### DE max 그래프
for i in range(3):
    if i==0:
        plt.figure(0, figsize=(16,7))
        plt.xlim(0, max_criteria_DE*1.1)
        plt.grid()
        plt.title('[%s] Shear Wall Rotation(DE)' % SWR_total.iloc[wall_name_change[i], 1], fontsize = 14, fontweight = 'bold')
        plt.xlabel('Strain', fontsize = 13, fontweight = 'bold')
        plt.ylabel('Floor', fontsize = 13, fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0: wall_name_change[i], 1], SWR_avg_total.iloc[0:wall_name_change[i], 0], marker='o', markerfacecolor = 'white', markersize = 5, color='k', linewidth = 4)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_total.iloc[0: wall_name_change[i],j+2], SWR_total.iloc[0:wall_name_change[i], 0], color='k',linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_DE, color = 'red', linestyle='--', linewidth = 2) # 기준선
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.9, 0.8))
    else:
        plt.figure(i, figsize=(16,7))
        plt.xlim(0, max_criteria_DE*1.1)
        plt.grid()
        plt.title('[%s] Shear Wall Rotation(DE)' % SWR_total.iloc[wall_name_change[i], 1], fontsize=14, fontweight='bold')
        plt.xlabel('Strain', fontsize=13, fontweight='bold')
        plt.ylabel('Floor', fontsize=13, fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 1], SWR_total.iloc[wall_name_change[i-1] + 1:wall_name_change[i], 0], marker='o', markerfacecolor = 'white', markersize = 5, color='k', linewidth = 4)  # avg plot line
        for j in range(DE_num):
            plt.plot(SWR_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2], SWR_total.iloc[wall_name_change[i-1]+1:wall_name_change[i], 0], color='k',linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_DE, color='red', linestyle='--', linewidth=2)  # 기준선
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + DE_name + ['LS'], bbox_to_anchor = (0.9, 0.8))



### MCE max 그래프
for i in range(3):
    if i==0:
        plt.figure(0, figsize=(16,7))
        plt.xlim(0, max_criteria_MCE*1.1)
        plt.grid()
        plt.title('[%s] Shear Wall Rotation(MCE)' % SWR_total.iloc[wall_name_change[i], 1], fontsize = 14, fontweight = 'bold')
        plt.xlabel('Strain', fontsize = 13, fontweight = 'bold')
        plt.ylabel('Floor', fontsize = 13, fontweight = 'bold')
        plt.plot(SWR_avg_total.iloc[0: wall_name_change[i], 3], SWR_avg_total.iloc[0:wall_name_change[i], 0], marker='o', markerfacecolor = 'white', markersize = 5, color='k', linewidth = 4)  # avg plot line
        for j in range(MCE_num):
            plt.plot(SWR_total.iloc[0: wall_name_change[i],j+2+DE_num], SWR_total.iloc[0:wall_name_change[i], 0], color='k',linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_MCE, color = 'red', linestyle='--', linewidth = 2) # 기준선
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + MCE_name + ['CP'], bbox_to_anchor = (0.9, 0.8))
    else:
        plt.figure(i, figsize=(16,7))
        plt.xlim(0, max_criteria_MCE*1.1)
        plt.grid()
        plt.title('[%s] Shear Wall Rotation(MCE)' % SWR_total.iloc[wall_name_change[i], 1], fontsize=14, fontweight='bold')
        plt.xlabel('Strain', fontsize=13, fontweight='bold')
        plt.ylabel('Floor', fontsize=13, fontweight='bold')
        plt.plot(SWR_avg_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], 3], SWR_total.iloc[wall_name_change[i-1] + 1:wall_name_change[i], 0], marker='o', markerfacecolor = 'white', markersize = 5, color='k', linewidth = 4)  # avg plot line
        for j in range(MCE_num):
            plt.plot(SWR_total.iloc[wall_name_change[i-1]+1: wall_name_change[i], j+2+DE_num], SWR_total.iloc[wall_name_change[i-1]+1:wall_name_change[i], 0], color='k',linestyle = 'dashed', linewidth = 1, alpha = 0.4)
        plt.axvline(max_criteria_MCE, color='red', linestyle='--', linewidth=2)  # 기준선
        plt.yticks(story_info['Height(mm)'][::-4], story_name[::-4])
        plt.legend(['Average'] + MCE_name + ['CP'], bbox_to_anchor = (0.9, 0.8))




