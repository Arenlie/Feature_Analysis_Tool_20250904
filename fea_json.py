import copy
import os

import pandas as pd
import json

from feature_values import *


def feature_json_all(input_path, output_path):
    sheets = pd.read_excel(input_path, sheet_name=None)

    # 输出所有工作表的名称和内容
    for sheet_name, df in sheets.items():
        # 创建文件夹（如果不存在的话）
        folder_path = os.path.join(output_path, sheet_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        try:
            feature_json(df, folder_path)
        except Exception as e:
            return f"Json文件输出失败，请检查DW-导入表文件格式"


def feature_json(input_data, output_path):
    bearing_data = pd.read_csv("后台文件/t_bearing_head.csv", encoding='GBK')
    my_null = '/'
    mac_address = input_data.iloc[0][2]
    board_type = ''
    card_index = ''
    card_judge = ''
    json_output_channel_settings = []
    json_output_channel_Feature = {}
    json_output_feature = []
    error_list = {}
    json_channel_setting_tmp = {}
    point_type_list = ['加速度', '速度', '径向振动位移', '轴向位移', '转速', '过程变量', '普通电压', '温度', '位移']
    # point_type_list = ['转速', '径向振动位移', '加速度', '速度', '声音', '轴向位移', '动态电压', '温度', '过程变量', '普通电压']
    json_file_path = '后台文件/tmp_ChannelSettings.json'
    with open(json_file_path, 'r', encoding='utf-8') as file:
        json_channel_settings = json.load(file)

    def none_judge(excel_cell, m_null=my_null):
        if not pd.isna(excel_cell) and excel_cell != m_null:
            return True
        else:
            return False

    for df_index, df_row in input_data.iterrows():
        row_error_list = []
        # 板卡是否启用
        if pd.isna(df_row['板卡是否启用']) and card_judge != '':
            row_card_judge = card_judge
        else:
            row_card_judge = df_row['板卡是否启用']
            card_judge = df_row['板卡是否启用']
        if row_card_judge == '否':
            continue
        if none_judge(df_row['板卡类型']) and not none_judge(df_row['测点（通道）类型']):
            row_error_list.append(f"NoneCardTypeError:card_type for {df_row['板卡编号']} is empty")
        elif not none_judge(df_row['测点（通道）类型']):
            row_error_list.append(f"NonePointTypeError:point_type for {df_row['板卡编号']} is empty")
        row_index = df_index + 2
        # 板卡类型
        if pd.isna(df_row['板卡类型']) and board_type != '':
            row_board_type = board_type
        else:
            row_board_type = df_row['板卡类型']
            board_type = df_row['板卡类型']

        # 测点类型
        point_type = df_row['测点（通道）类型']
        # 设备名称
        equipment_name = df_row['设备名称']
        if pd.isna(equipment_name):
            row_error_list.append(f'NoneEquipmentNameError:equipment_name at {row_index} row is empty')
        # 通道编号
        try:
            channel_index = df_row['通道编号'][-1]
        except TypeError:
            channel_index = 'x'
            row_error_list.append(f'WrongChannelIndexError:channel_index at {row_index} row is wrong')
        # 板卡编号
        if pd.isna(df_row['板卡编号']) and card_index != '':
            row_card_index = card_index
        else:
            try:
                row_card_index = df_row['板卡编号'][-2:]
            except TypeError:
                row_error_list.append(f'WrongCardIndexError:card_index at {row_index} row is wrong')
                row_card_index = 'x'
            card_index = row_card_index

        if row_board_type == '高速卡':
            if point_type not in ['加速度', '转速']:
                row_error_list.append(f'WrongPointTypeError:point_type at {row_index} row is wrong')
            if df_row['键相类型'] == '虚拟键相' and point_type != '转速':
                try:
                    _ = int(df_row['工作转速'])
                except ValueError:
                    row_error_list.append(f'WrongVirtualRevError:virtual_rev at {row_index} row is wrong')
            elif df_row['键相类型'] == '外部键相' and point_type != '转速':
                try:
                    _ = mac_address + df_row['工作转速'][1:3] + df_row['工作转速'][-1]
                except TypeError:
                    row_error_list.append(f'WrongRealRevError:real_rev at {row_index} row is wrong')
        if row_board_type == '低速卡':
            if point_type not in ['普通电压', '温度', '过程变量']:
                row_error_list.append(f'WrongPointTypeError:point_type at {row_index} row is wrong')

        if row_card_index != 'x' and channel_index != 'x':
            channel_id = mac_address + row_card_index + channel_index
        else:
            channel_id = row_index
        if row_error_list:
            error_list[channel_id] = row_error_list

    if error_list:
        file = open(f'{output_path}/error_list.json', 'w', encoding='utf-8')
        file.write(json.dumps(error_list, ensure_ascii=False))
        file.close()
        return 'error'

    json_feature_path = '后台文件/tmp_features.json'
    with open(json_feature_path, 'r', encoding='utf-8') as file:
        json_feature = json.load(file)

    for df_index, df_row in input_data.iterrows():
        # 板卡类型
        if pd.isna(df_row[6]) and board_type != '':
            row_board_type = board_type
        else:
            row_board_type = df_row[6]
            board_type = df_row[6]

        # 测点类型
        point_type = df_row['测点（通道）类型']
        # 设备名称
        equipment_name = df_row['设备名称']
        # 通道编号
        channel_index = df_row['通道编号'][-1]
        # 测点点位名称
        channel_name = df_row['测点（点位）名称']
        # 板卡编号
        if pd.isna(df_row['板卡编号']) and card_index != '':
            row_card_index = card_index
        else:
            row_card_index = df_row['板卡编号'][-2:]
            card_index = row_card_index

        # 创建row_json_cs
        point_type_index = point_type_list.index(point_type)
        # print("point_type_index", point_type_index)
        row_json_cs = copy.deepcopy(json_channel_settings[0])
        # 指定测点名称
        if not none_judge(channel_name):
            row_json_cs['channelName'] = f'第{int(row_card_index)}板卡第{int(channel_index)}通道'
        else:
            row_json_cs['channelName'] = channel_name
        row_json_cs['sensorModuleNo'] = int(row_card_index)
        row_json_cs['sensorNo'] = int(channel_index)

        # 板卡是否启用
        if pd.isna(df_row['板卡是否启用']) and card_judge != '':
            row_card_judge = card_judge
        else:
            row_card_judge = df_row['板卡是否启用']
            card_judge = df_row['板卡是否启用']
        # 板卡未启用时的处理
        if row_card_judge == '否':
            if row_board_type == '高速卡':
                json_channel_setting_tmp = copy.deepcopy(json_channel_settings[0])
            else:
                json_channel_setting_tmp = copy.deepcopy(json_channel_settings[6])
            json_channel_setting_tmp['channelId'] = mac_address + row_card_index + channel_index
            json_channel_setting_tmp['isEnable'] = 0
            json_channel_setting_tmp['isWork'] = 0
            if not none_judge(channel_name):
                json_channel_setting_tmp['channelName'] = f'第{int(row_card_index)}板卡第{int(channel_index)}通道'
            else:
                json_channel_setting_tmp['channelName'] = channel_name
            json_channel_setting_tmp['sensorModuleNo'] = int(row_card_index)
            json_channel_setting_tmp['sensorNo'] = int(channel_index)
            json_output_channel_settings.append(json_channel_setting_tmp)
            continue

        row_json_cs['dataWatchNo'] = mac_address
        row_json_cs['equipmentName'] = equipment_name
        row_json_cs['channelId'] = mac_address + row_card_index + channel_index
        # 低速卡
        if row_board_type == '低速卡' or point_type == '转速':
            json_output_channel_settings.append(row_json_cs)
            continue
        # channel_Feature
        json_output_channel_Feature[row_json_cs['channelId']] = 1
        # 高速卡
        # row_json_cs['samplingRate'] = 32768
        # 轴向位移
        row_json_feature = copy.deepcopy(json_feature[point_type_index])
        row_json_feature['channel_id'] = row_json_cs['channelId']
        if point_type == '轴向位移':
            json_output_feature.append(row_json_feature)
            json_output_channel_settings.append(row_json_cs)
            continue

        if df_row['键相类型'] == '虚拟键相':
            row_json_cs['analog_rpm'] = int(df_row['工作转速'])
            row_json_feature['analog_rpm'] = int(df_row['工作转速'])
        elif df_row['键相类型'] == '外部键相':
            del row_json_cs['analog_rpm']
            rev_channel_id = mac_address + df_row['工作转速'][1:3] + df_row['工作转速'][-1]
            row_json_cs['speedRefChannelId'] = rev_channel_id
        json_output_channel_settings.append(row_json_cs)

        # json_feature
        N, nc, n, f0, m, Bearing_designation, Manufacturer, Z, vane, G_vane = df_row[13:]
        if none_judge(N):
            row_json_feature['base_freq_domain'] = [kBaseFreqDomain1X, kBaseFreqDomain2X, kBaseFreqDomain3X,
                                                    kBaseFreqDomain4X, kBaseFreqDomain5X, kBaseFreqDomain1D2X,
                                                    kBaseFreqDomain1D3X, kBaseFreqDomain1D4X, kBaseFreqDomain1D5X]
        if none_judge(f0):
            row_json_feature['motor_fault'] = {
                'feature_list': [kMotorFaultPowerFreq2X, kMotorFaultPowerFreq4X, kMotorFaultPowerFreq6X,
                                 kMotorFaultPowerFreq8X, kMotorFaultPowerFreq10X], 'power_freq': df_row['电源频率']}
        if none_judge(f0) and none_judge(nc) and none_judge(n):
            row_json_feature['motor_fault']['feature_list'].extend(
                [kMotorFaultAirGapHCR1X, kMotorFaultAirGapHCR2X, kMotorFaultAirGapHCR3X, kMotorFaultAirGapHCR4X,
                 kMotorFaultAirGapHCR5X])
            row_json_feature['motor_fault']['rated_rpm'] = nc
            row_json_feature['motor_fault']['sync_rpm'] = n
            if none_judge(N):
                row_json_feature['motor_fault']['feature_list'].extend(
                    [kMotorFaultFractureHCR1X, kMotorFaultFractureHCR2X, kMotorFaultFractureHCR3X,
                     kMotorFaultFractureHCR4X,
                     kMotorFaultFractureHCR5X])
            if none_judge(m):
                row_json_feature['motor_fault']['feature_list'].extend(
                    [kMotorFaultLoosenHCR1X, kMotorFaultLoosenHCR2X, kMotorFaultLoosenHCR3X, kMotorFaultLoosenHCR4X,
                     kMotorFaultLoosenHCR5X])
                row_json_feature['motor_fault']['rotor_number'] = m
        if none_judge(N) and none_judge(Z):
            row_json_feature['gear_fault'] = {'feature_list': list(range(kGearMeshFreqAmp1X, kGearFaultHRSLOWER5X + 1)),
                                              'teeth_number': Z}
        if none_judge(N) and (none_judge(vane) or none_judge(G_vane)):
            row_json_feature['other_fault'] = {'feature_list': []}
            if none_judge(vane):
                row_json_feature['other_fault']['feature_list'].extend(
                    [kImpellerBladePF1X, kImpellerBladePF2X, kImpellerBladePF3X, kImpellerBladePF4X, kImpellerBladePF5X,
                     kRotatingStallImpellerHRS])
                row_json_feature['other_fault']['impeller_blade_num'] = vane
            if none_judge(G_vane):
                row_json_feature['other_fault']['feature_list'].extend(
                    [kDerivedBladePF1X, kDerivedBladePF2X, kDerivedBladePF3X, kDerivedBladePF4X, kDerivedBladePF5X,
                     kRotatingStallDerivedHRS])
                row_json_feature['other_fault']['derived_blade_num'] = G_vane
        if none_judge(N) and Bearing_designation == "滑动轴承":
            # noinspection PyTypedDict
            row_json_feature['rolling_bear'] = {"feature_list": [kOilWhirlHRS]}
        elif none_judge(N) and none_judge(Bearing_designation) and none_judge(Manufacturer):
            Bearing_designation = str(Bearing_designation).split(".")[0]
            bearing_one = bearing_data.loc[
                (bearing_data['NAME=型号'] == Bearing_designation) & (bearing_data['MANUFACTURE'] == Manufacturer)]
            if not bearing_one.empty:
                row_json_feature['rolling_bear'] = {'bear_id': int(bearing_one.iloc[0][0]),
                                                    'feature_list': list(range(kRollingBearingInnerRing1X,
                                                                               kRollingBearingRollingElement5X + 1))}
        json_output_feature.append(row_json_feature)

    # if input_data['通道编号'].astype(str).str.contains("CH08").any():
    #     for i in range(len(input_data), 16):
    #         card_index = i // 8 + 1
    #         channel_index = i % 8 + 1
    #         json_channel_setting_tmp['channelId'] = mac_address + str(f'0{card_index}') + str(channel_index)
    #         json_output_channel_settings.append(copy.deepcopy(json_channel_setting_tmp))
    # else:
    #     for i in range(len(input_data), 32):
    #         card_index = i // 4 + 1
    #         channel_index = i % 4 + 1
    #         json_channel_setting_tmp['channelId'] = mac_address + str(f'0{card_index}') + str(channel_index)
    #         json_output_channel_settings.append(copy.deepcopy(json_channel_setting_tmp))

    file = open(f'{output_path}/ChannelSettings.json', 'w', encoding='utf-8')
    file.write(json.dumps(json_output_channel_settings, ensure_ascii=False))
    file.close()

    if json_output_channel_Feature:
        file = open(f'{output_path}/FeatureCalc.json', 'w', encoding='utf-8')
        file.write(json.dumps(json_output_channel_Feature, ensure_ascii=False))
        file.close()

    if json_output_feature:
        file = open(f'{output_path}/Features.json', 'w', encoding='utf-8')
        file.write(json.dumps(json_output_feature, ensure_ascii=False))
        file.close()


if __name__ == '__main__':
    # from dataToDWTable import dataToDWTable
    # dataToDWTable(r"C:\Users\Administrator\Desktop\测试\data_all(2700).xlsx",
    #               r"C:\Users\Administrator\Desktop\测试\DW2700导入表.xlsx")
    if_err = feature_json_all(r"C:\Users\Administrator\Desktop\测试\DW2300导入表.xlsx",
                              r"C:\Users\Administrator\Desktop\测试\json_2300")
    if if_err:
        print(if_err)
    else:
        print(r"Json文件保存到：\nC:\Users\Administrator\Desktop\测试\json_2300")
