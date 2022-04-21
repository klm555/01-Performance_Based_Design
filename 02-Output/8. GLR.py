### Gravity Load Ratio - IO/LS/CP
# 층별 중력 하중 분담률 계산
# 필요파일 : ECHO.xlsx, Result Analysis, Input Sheet
# 폴더구성 : data라는 이름의 폴더 안에 ECHO, Result Analysis, Input Sheet 엑셀 파일들을 넣는다.
# Output 파일은 data에 저장된다
# Running time : 약 3 min

import os
import pandas as pd
from collections import deque
from io import StringIO
import matplotlib.pyplot as plt



### 0. Initial Setting
# Data path
data_path = r'D:\이형우\내진성능평가\광명 4R\해석 결과\110' # data 폴더 경로

input_raw_xlsx_dir = r'D:\이형우\내진성능평가\광명 4R\110'
input_raw_xlsx = 'Input Sheets(110).xlsx'

# 지진파 개수
DE_num = 14
MCE_num = 14

# IO/LS/CP/Collapse limit value
# wall rotation의 내진성능설계 평가 기준 IO, LS, CP값 입력
IO_lim = 0.001
LS_lim = 0.002
CP_lim = 0.004/1.2


### 1. Data preparation(ECHO/Analysis Result/Section/Input Sheets(Story Info))
## ECHO File
ECHO_path = data_path + '\\' + 'ECHO.xlsx'
ECHO_data = pd.read_excel(ECHO_path)

## Analysis Result
to_load_list = []
file_names = os.listdir(data_path)
for file_name in file_names:
    if 'Analysis Result' in file_name:
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

SWR_result_data = SWR_result_data.sort_values(by=['Load Case', 'Element ID', 'Step Type'])

# Node Coord data
Node_coord_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Node Coordinate Data', skiprows=[0, 2], header=0, usecols=[1, 2, 3, 4])


# Section data
section_data = pd.read_excel(data_path + '\\' + to_load_list[0],
                               sheet_name='Structure Section Forces', skiprows=[0, 2], header=0, usecols=[0, 1, 3, 5, 8]) # 먼저 Load Case 열만 불러온 후 section 개수 구한다음 전체 불러오기

section_data = section_data.sort_values(by=['Load Case', 'StrucSec ID', 'Step Type'])

for i in range(len(section_data)):
    if section_data.iloc[i, 2] != section_data.iloc[i+1, 2]:
        break
section_num = int(i+1)

section_data = section_data.iloc[0:section_num, :]
section_data = section_data[section_data['Step Type'] == 'Min'] # 어차피 max 나 min이나 똑같음


# Story Info data
story_info = pd.read_excel(input_raw_xlsx_dir + '\\' + input_raw_xlsx, sheet_name='Story Info', keep_default_na=False)
story_height = story_info.loc[:, 'Height(mm)']
story_name = story_info.loc[:, 'Floor Name']


### Gage data에서 Element ID, I-Node ID 불러와서 v좌표 match하기
SWR_gage_data = SWR_gage_data[['Element ID', 'I-Node ID']]; gage_num = len(SWR_gage_data) # gage 개수 얻기
Node_v_coord_data = Node_coord_data[['Node ID', 'V']]

# I-Node의 v좌표 match해서 추가
SWR_gage_data = SWR_gage_data.join(Node_coord_data.set_index('Node ID')['V'], on='I-Node ID')



### SWR_total data 만들기

SWR_max = SWR_result_data[(SWR_result_data['Step Type'] == 'Max') & (SWR_result_data['Performance Level'] == 1)][['Rotation']].values # dataframe을 array로
SWR_max = SWR_max.reshape(gage_num, DE_num + MCE_num, order='F') # order = 'C' 인 경우 row 우선 변경, order = 'F'인 경우 column 우선 변경
SWR_max = pd.DataFrame(SWR_max) # array를 다시 dataframe으로
SWR_min = SWR_result_data[(SWR_result_data['Step Type'] == 'Min') & (SWR_result_data['Performance Level'] == 1)][['Rotation']].values
SWR_min = SWR_min.reshape(gage_num, DE_num + MCE_num, order='F')
SWR_min = pd.DataFrame(SWR_min)
SWR_total = pd.concat([SWR_max, SWR_min], axis=1) # DE11_max~MCE72_max, DE11_min~MCE72_min 각각 28개씩

### SWR_avg_data 만들기
DE_max_avg = SWR_total.iloc[:, 0:DE_num].mean(axis=1)
MCE_max_avg = SWR_total.iloc[:, DE_num:DE_num + MCE_num].mean(axis=1)
DE_min_avg = SWR_total.iloc[:, DE_num + MCE_num:2*DE_num + MCE_num].mean(axis=1)
MCE_min_avg = SWR_total.iloc[:, 2*DE_num + MCE_num:2*DE_num + 2*MCE_num].mean(axis=1)
SWR_avg_total = pd.concat([SWR_gage_data['V'], DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
SWR_avg_total.columns = ['Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

SWR_total = pd.concat([DE_max_avg, DE_min_avg, MCE_max_avg, MCE_min_avg], axis=1)
SWR_total.reset_index(drop=True, inplace=True)


# total data
SWR_total = pd.concat([SWR_gage_data['V'], SWR_total], axis=1)
SWR_total.columns = ['Height', 'DE_max_avg', 'DE_min_avg', 'MCE_max_avg', 'MCE_min_avg']

# Determine I0/LS/CP
result_1 = []

for i in range(len(SWR_total)):
    if SWR_total.iloc[i, 1]<IO_lim and abs(SWR_total.iloc[i, 2])<IO_lim:
        result_1.append('IO')
    elif SWR_total.iloc[i, 1]<LS_lim and abs(SWR_total.iloc[i, 2])<LS_lim:
        result_1.append('LS')
    elif SWR_total.iloc[i, 1]<CP_lim and abs(SWR_total.iloc[i, 2])<CP_lim:
        result_1.append('CP')
    else:
        result_1.append('Collapse')

SWR_total['DE_result'] = result_1

result_2 = []

for i in range(len(SWR_total)):
    if SWR_total.iloc[i, 3]<IO_lim and abs(SWR_total.iloc[i, 4])<IO_lim:
        result_2.append('IO')
    elif SWR_total.iloc[i, 3]<LS_lim and abs(SWR_total.iloc[i, 4])<LS_lim:
        result_2.append('LS')
    elif SWR_total.iloc[i, 3]<CP_lim and abs(SWR_total.iloc[i, 4])<CP_lim:
        result_2.append('CP')
    else:
        result_2.append('Collapse')

SWR_total['MCE_result'] = result_2


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
gage_data = gage_data.join(wall_data.set_index(['i', 'direction'])['wall_name'], on=['j', 'direction'], rsuffix='_right'); gage_data.columns.values[12] = 'gage_name_full_j'  # gage와 wall을 연결하는 작업, j를 기준으로

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

## Matching Section with Gage
section_name_data = section_data.iloc[:,0]
section_name_head = section_name_data.apply(lambda x : x.split('_')[0]); section_name_head.reset_index(drop=True, inplace=True)
section_name_tail = section_name_data.apply(lambda x : x.split('_')[-1]); section_name_tail.reset_index(drop=True, inplace=True)
section_GV_data = abs(section_data.iloc[:,4])
section_name_data.reset_index(drop=True, inplace=True);
section_GV_data.reset_index(drop=True, inplace=True);
section_GV_data = pd.concat([section_name_data, section_name_head, section_name_tail, section_GV_data], axis = 1, ignore_index=True)
section_GV_data.columns = ['Section Name', 'Name', 'Story', 'GV']


gage_data['DE_result'] = result_1
gage_data['MCE_result'] = result_2
gage_GV_data = gage_data.join(section_GV_data.set_index(['Name', 'Story'])['GV'], on=['gage_name', 'Story'])
gage_GV_data = gage_GV_data.iloc[:, [11, 12, 13, 14, 15, 16]]
gage_GV_data = gage_GV_data.drop_duplicates(subset=['GV'])
pd.to_numeric(gage_GV_data.iloc[:,5]) # gage_GV_data의 GV값을 numeric으로 바꾸기


###5. Performance Table
PF_table = story_info['Floor Name']
PF_table = PF_table.to_frame()
PF_table_DE = PF_table
PF_table_MCE = PF_table


GV_sum = 0
GV_total = [] # 층별 중력하중의 총합
for i in range(len(PF_table)): # RF 빼고
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0]:
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_total.append(GV_sum)
    GV_sum = 0

GV_IO_DE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 3] == 'IO':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_IO_DE.append(GV_sum)
    GV_sum = 0


GV_LS_DE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 3] == 'LS':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_LS_DE.append(GV_sum)
    GV_sum = 0

GV_CP_DE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 3] == 'CP':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_CP_DE.append(GV_sum)
    GV_sum = 0

GV_collapse_DE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 3] == 'Collapse':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_collapse_DE.append(GV_sum)
    GV_sum = 0

GV_IO_MCE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 4] == 'IO':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_IO_MCE.append(GV_sum)
    GV_sum = 0

GV_LS_MCE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 4] == 'LS':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_LS_MCE.append(GV_sum)
    GV_sum = 0

GV_CP_MCE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 4] == 'CP':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_CP_MCE.append(GV_sum)
    GV_sum = 0

GV_collapse_MCE = []
for i in range(len(PF_table)):
    for j in range(len(gage_GV_data)):
        if gage_GV_data.iloc[j, 1] == PF_table.iloc[i, 0] and gage_GV_data.iloc[j, 4] == 'Collapse':
            GV_sum = GV_sum + gage_GV_data.iloc[j, 5]
    GV_collapse_MCE.append(GV_sum)
    GV_sum = 0



# Performance table of DE
PF_table_DE = pd.concat([PF_table_DE, pd.Series(GV_total), pd.Series(GV_IO_DE)], axis = 1, ignore_index=True)
PF_table_DE['IO(%)'] = PF_table_DE[2]/PF_table_DE[1]*100 # IO 비율
PF_table_DE = pd.concat([PF_table_DE, pd.Series(GV_LS_DE)], axis = 1, ignore_index=True)
PF_table_DE['LS(%)'] = PF_table_DE[4]/PF_table_DE[1]*100 # LS 비율
PF_table_DE = pd.concat([PF_table_DE, pd.Series(GV_CP_DE)], axis = 1, ignore_index=True)
PF_table_DE['CP(%)'] = PF_table_DE[6]/PF_table_DE[1]*100 # CP 비율
PF_table_DE = pd.concat([PF_table_DE, pd.Series(GV_collapse_DE)], axis = 1, ignore_index=True)
PF_table_DE['Collapse(%)'] = PF_table_DE[8]/PF_table_DE[1]*100 # Collapse 비율
PF_table_DE = PF_table_DE.T.reset_index(drop=True).T # column index reset
PF_table_DE = PF_table_DE.fillna(0)
PF_table_DE.columns = ['Story', 'Total', 'IO', 'IO(%)', 'LS', 'LS(%)', 'CP', 'CP(%)', 'Collapse', 'Collapse(%)']


# Performance table of MCE
PF_table_MCE = pd.concat([PF_table_MCE, pd.Series(GV_total), pd.Series(GV_IO_MCE)], axis = 1, ignore_index=True)
PF_table_MCE['IO(%)'] = PF_table_MCE[2]/PF_table_MCE[1]*100
PF_table_MCE = pd.concat([PF_table_MCE, pd.Series(GV_LS_MCE)], axis = 1, ignore_index=True)
PF_table_MCE['LS(%)'] = PF_table_MCE[4]/PF_table_MCE[1]*100
PF_table_MCE = pd.concat([PF_table_MCE, pd.Series(GV_CP_MCE)], axis = 1, ignore_index=True)
PF_table_MCE['CP(%)'] = PF_table_MCE[6]/PF_table_MCE[1]*100
PF_table_MCE = pd.concat([PF_table_MCE, pd.Series(GV_collapse_MCE)], axis = 1, ignore_index=True)
PF_table_MCE['Collapse(%)'] = PF_table_MCE[8]/PF_table_MCE[1]*100
PF_table_MCE = PF_table_MCE.T.reset_index(drop=True).T
PF_table_MCE = PF_table_MCE.fillna(0)
PF_table_MCE.columns = ['Story', 'Total', 'IO', 'IO(%)', 'LS', 'LS(%)', 'CP', 'CP(%)', 'Collapse', 'Collapse(%)']



# Determine final Performance
final_Perf_DE = []
for i in range(len(PF_table)):
    if PF_table_DE.iloc[i, 2] >= 80:
        final_Perf_DE.append('IO')
    elif PF_table_DE.iloc[i, 2] + PF_table_DE.iloc[i, 4] >= 80:
        final_Perf_DE.append('LS')
    elif PF_table_DE.iloc[i, 2] + PF_table_DE.iloc[i, 4] + PF_table_DE.iloc[i, 6] >= 80:
        final_Perf_DE.append('CP')
    else:
        final_Perf_DE.append('Collapse')

final_Perf_MCE = []
for i in range(len(PF_table)):
    if PF_table_MCE.iloc[i, 2] >= 80:
        final_Perf_MCE.append('IO')
    elif PF_table_MCE.iloc[i, 2] + PF_table_MCE.iloc[i, 4] >= 80:
        final_Perf_MCE.append('LS')
    elif PF_table_MCE.iloc[i, 2] + PF_table_MCE.iloc[i, 4] + PF_table_MCE.iloc[i, 6] >= 80:
        final_Perf_MCE.append('CP')
    else:
        final_Perf_MCE.append('Collapse')

PF_table_DE['Perf.'] = final_Perf_DE
PF_table_MCE['Perf.'] = final_Perf_MCE

### 6. Export data
# Export performance table as .xlsx file
PF_table_DE.to_excel(data_path + '\\' + 'Performance Table_DE.xlsx')
PF_table_MCE.to_excel(data_path + '\\' + 'Performance Table_MCE.xlsx')

