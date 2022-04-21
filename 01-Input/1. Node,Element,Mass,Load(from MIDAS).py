"""
###  update(2021.10.22)  ###
1. 경로 설정 편하게 다 앞으로 몰아놨음.
2. Mass의 node도 자동으로 추가되게 해놨음.
"""

import pandas as pd

# Input 경로 설정
node_raw_xlsx_dir = r'C:\Users\hwlee\Desktop\Python\내진성능설계'
node_raw_xlsx = 'Data Conversion_Shear Wall Type_Ver.1.0_220224.xlsx'

node_raw_xlsx_sheet = 'Nodes'
nodal_load_raw_xlsx_sheet = 'Nodal Loads'
mass_raw_xlsx_sheet = 'Story Mass'
element_raw_xlsx_sheet = 'Elements'

# Output 경로 설정
output_csv_dir = node_raw_xlsx_dir # 또는 '경로'

node_DL_merged_csv = 'DL.csv'
node_LL_merged_csv = 'LL.csv'
mass_csv = 'Mass.csv'
node_csv = 'Nodes.csv'
beam_csv = 'Beam.csv'
wall_csv = 'Wall.csv'
plate_csv = 'Plate.csv'

#%% Nodal Load 뽑기

# Node 정보 load
node = pd.read_excel(node_raw_xlsx_dir + '\\' + node_raw_xlsx, sheet_name = node_raw_xlsx_sheet, skiprows = 3, index_col = 0)  # Node 열을 인덱스로 지정
node.columns = ['X(mm)', 'Y(mm)', 'Z(mm)']

# Nodal Load 정보 load
nodal_load = pd.read_excel(node_raw_xlsx_dir+'\\'+node_raw_xlsx, sheet_name = nodal_load_raw_xlsx_sheet, skiprows = 3, index_col = 0)
nodal_load.columns = ['Loadcase', 'FX(kN)', 'FY(kN)', 'FZ(kN)', 'MX(kN-mm)', 'MY(kN-mm)', 'MZ(kN-mm)']

# Nodal Load를 DL과 LL로 분리
DL = nodal_load.loc[lambda x: nodal_load['Loadcase'].str.contains('DL'), :]  # lambda로 만든 함수로 Loadcase가 DL인 행만 slicing
LL = nodal_load.loc[lambda x: nodal_load['Loadcase'].str.contains('LL'), :] # 마찬가지. Loadcase의 첫 두글자가 LL인 행만 slicing

DL2 = DL.drop('Loadcase', axis=1)  # axis=1(열), axis=0(행)
LL2 = LL.drop('Loadcase', axis=1)  # 필요없어진 Loadcase 열은 drop으로 떨굼

# Node와 Nodal Load를 element number 기준으로 병합
node_DL_merged = pd.merge(node, DL2, left_index=True, right_index=True)  # node 좌표와 하중을 결합하여 dataframe 만들기, merge : 공통된 index를 기준으로 합침
node_LL_merged = pd.merge(node, LL2, left_index=True, right_index=True)  # left_index, right_index는 뭔지 기억은 안나는데 오류고치기위해서 더함

# DL, LL 결과값을 csv로 변환
node_DL_merged.to_csv(output_csv_dir+'\\'+node_DL_merged_csv, mode='w', index=False)  # to_csv 사용. index=False로 index 열은 떨굼
node_LL_merged.to_csv(output_csv_dir+'\\'+node_LL_merged_csv, mode='w', index=False)

#%% Mass 뽑기

# Mass 정보 load
mass = pd.read_excel(node_raw_xlsx_dir+'\\'+node_raw_xlsx, sheet_name = mass_raw_xlsx_sheet, skiprows = 3)
mass.columns = ['Story', 'Level(mm)', 'Trans Mass X-dir(kN/g)', 'Trans Mass Y-dir(kN/g)', 'Rotat Mass(kN/g-mm^2)', 'X(mm)', 'Y(mm)']

# 필요없는 열 제거
mass2 = mass.drop('Story', axis=1)

# 열 재배치
mass2 = mass2[['X(mm)', 'Y(mm)', 'Level(mm)', 'Trans Mass X-dir(kN/g)', 'Trans Mass Y-dir(kN/g)', 'Rotat Mass(kN/g-mm^2)']]

# 형태 맞추기 위해 열 추가
mass2.insert(5, 'Trans Mass Z-dir(kN/g)', 0)  # insert로 5번째 열의 위치에 column 삽입
mass2.insert(6, 'Rotat Mass X-dir(kN/g mm^2)', 0)
mass2.insert(7, 'Rotat Mass Y-dir(kN/g mm^2)', 0)

# Mass 결과값을 csv로 변환
mass2.to_csv(output_csv_dir+'\\'+mass_csv, mode='w', index=False)

#%% Node 뽑기

# Mass의 좌표 추가
node_mass_considered = node.append(mass2.iloc[:,[0,1,2]])

# Node 결과값을 csv로 변환
node_mass_considered.to_csv(output_csv_dir+'\\'+node_csv, mode='w', index=False)

#%% Beam Element 뽑기

# Element 정보 load
element = pd.read_excel(node_raw_xlsx_dir+'\\'+node_raw_xlsx, sheet_name = element_raw_xlsx_sheet, skiprows = 3)
element.columns = ['Element', 'Type', 'Wall Type', 'Sub Type', 'Wall ID', 'Material', 'Property', 'B-Angle', 'Node1', 'Node2', 'Node3', 'Node4']

# Beam Element만 추출(slicing)
beam = element.loc[lambda x: element['Type'] == 'BEAM', :]

# 필요한 열만 추출(drop하기에는 drop할 열이 너무 많아서...)
beam_node_1 = beam.loc[:, 'Node1']
beam_node_2 = beam.loc[:, 'Node2']

beam_node_1.name = 'Node'  # Merge(같은 열을 기준으로 두 dataframe 결합)를 사용하기 위해 index를 Node로 바꾸기
beam_node_2.name = 'Node'

node.index.name = 'Node'
node.reset_index(inplace=True) # Index로 지정되어있던 Node 번호를 다시 reset

# Merge로 Node 번호에 맞는 좌표를 결합
beam_node_1_coord = pd.merge(beam_node_1, node, how='left')  # how='left' : 두 데이터프레임 중 왼쪽 데이터프레임은 그냥 두고 오른쪽 데이터프레임값을 대응시킴
beam_node_2_coord = pd.merge(beam_node_2, node, how='left')

# Node1, Node2의 좌표를 모두 결합시켜 출력
beam_node_1_coord = beam_node_1_coord.drop('Node', axis=1)
beam_node_2_coord = beam_node_2_coord.drop('Node', axis=1)

beam_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']  # 결합 때 이름이 중복되면 안되서 이름 바꿔줌
beam_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']

beam_node_coord = pd.concat([beam_node_1_coord, beam_node_2_coord], axis=1)

# Beam Element 결과값을 csv로 변환
beam_node_coord.to_csv(output_csv_dir+'\\'+beam_csv, mode='w', index=False) 

#%% Wall Element 뽑기

# Wall Element만 추출(slicing)
wall = element.loc[lambda x: element['Type'] == 'WALL', :]

# 필요한 열만 추출
wall_node_1 = wall.loc[:, 'Node1']
wall_node_2 = wall.loc[:, 'Node2']
wall_node_3 = wall.loc[:, 'Node3']
wall_node_4 = wall.loc[:, 'Node4']

wall_node_1.name = 'Node'
wall_node_2.name = 'Node'
wall_node_3.name = 'Node'
wall_node_4.name = 'Node'

# Merge로 Node 번호에 맞는 좌표를 결합
wall_node_1_coord = pd.merge(wall_node_1, node, how='left')
wall_node_2_coord = pd.merge(wall_node_2, node, how='left')
wall_node_3_coord = pd.merge(wall_node_3, node, how='left')
wall_node_4_coord = pd.merge(wall_node_4, node, how='left')

# Node1, Node2, Node3, Node4의 좌표를 모두 결합시켜 출력
wall_node_1_coord = wall_node_1_coord.drop('Node', axis=1)
wall_node_2_coord = wall_node_2_coord.drop('Node', axis=1)
wall_node_3_coord = wall_node_3_coord.drop('Node', axis=1)
wall_node_4_coord = wall_node_4_coord.drop('Node', axis=1)

wall_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']
wall_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
wall_node_3_coord.columns = ['X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
wall_node_4_coord.columns = ['X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']

# wall_node_coord_list = [wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord]
wall_node_coord = pd.concat([wall_node_1_coord, wall_node_2_coord, wall_node_3_coord, wall_node_4_coord], axis=1)

# Wall Element 결과값을 csv로 변환
wall_node_coord.to_csv(output_csv_dir+'\\'+wall_csv, mode='w', index=False) 

#%% Plate Element 뽑기

# Plate Element만 추출(slicing)
if 'PLATE' in element['Type']:
    plate = element.loc[lambda x: element['Type'] == 'PLATE', :]
    
    # 필요한 열만 추출
    plate_node_1 = plate.loc[:, 'Node1']
    plate_node_2 = plate.loc[:, 'Node2']
    plate_node_3 = plate.loc[:, 'Node3']
    plate_node_4 = plate.loc[:, 'Node4']
    
    plate_node_1.name = 'Node'
    plate_node_2.name = 'Node'
    plate_node_3.name = 'Node'
    plate_node_4.name = 'Node'
    
    # Merge로 Node 번호에 맞는 좌표를 결합
    plate_node_1_coord = pd.merge(plate_node_1, node, how='left')
    plate_node_2_coord = pd.merge(plate_node_2, node, how='left')
    plate_node_3_coord = pd.merge(plate_node_3, node, how='left')
    plate_node_4_coord = pd.merge(plate_node_4, node, how='left')
    
    # Node1, Node2, Node3, Node4의 좌표를 모두 결합시켜 출력
    plate_node_1_coord = plate_node_1_coord.drop('Node', axis=1)
    plate_node_2_coord = plate_node_2_coord.drop('Node', axis=1)
    plate_node_3_coord = plate_node_3_coord.drop('Node', axis=1)
    plate_node_4_coord = plate_node_4_coord.drop('Node', axis=1)
    
    plate_node_1_coord.columns = ['X_1(mm)', 'Y_1(mm)', 'Z_1(mm)']
    plate_node_2_coord.columns = ['X_2(mm)', 'Y_2(mm)', 'Z_2(mm)']
    plate_node_3_coord.columns = ['X_3(mm)', 'Y_3(mm)', 'Z_3(mm)']
    plate_node_4_coord.columns = ['X_4(mm)', 'Y_4(mm)', 'Z_4(mm)']
    
    # plate_node_coord_list = [plate_node_1_coord, plate_node_2_coord, plate_node_3_coord, plate_node_4_coord]
    plate_node_coord = pd.concat([plate_node_1_coord, plate_node_2_coord, plate_node_3_coord, plate_node_4_coord], axis=1)
    
    # plate Element 결과값을 csv로 변환
    plate_node_coord.to_csv(output_csv_dir+'\\'+plate_csv, mode='w', index=False)
    
else: pass