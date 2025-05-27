from dateutil import parser
from datetime import datetime
import jionlp as jio
import re

def parse_id_card(id_number):
    try:
        current_year = datetime.now().year
        parse_result = jio.parse_id_card(str(id_number))
        if parse_result and 'birth_year' in parse_result:
            birth_year = parse_result['birth_year']
            age = current_year - int(birth_year)
            return {'年龄': age, '性别': parse_result['gender']}
    except:
        return {'年龄': "不详", '性别': "不详"}
    return {'年龄': "不详", '性别': "不详"}

def get_decedents_age_cause_view(df):
    # 筛选出伤害程度为"死亡"的数据，获取事故编号集合
    death_cases = df[df['伤害程度'] == '死亡']
    dset = set(death_cases['事故编号'].unique())
    
    if len(dset) == 0:
        print("没有找到死亡案例。")
        return {}, 0

    # 在原始数据中筛选出事故编号在dset中的所有记录
    related_df = df[df['事故编号'].isin(dset)].drop_duplicates(subset=['身份证明号码'], keep='first')
    
    # 计算年龄和性别
    age_sex = related_df['身份证明号码'].apply(parse_id_card)
    related_df['年龄'] = age_sex.apply(lambda x: x['年龄'])
    related_df['性别'] = age_sex.apply(lambda x: x['性别'])
    # 创建以事故编号为key的字典
    res_dict = {}
    for accident_id in dset:
        res_dict[accident_id] = related_df[related_df['事故编号'] == accident_id]

    return res_dict, len(dset)

def get_accidents_decedents_detail_view(dec_base_dict, dec_acc_count, dataframe):
    dec_base_details = ""
    dec_time_details = ""
    dec_road_details = ""
    dec_age_details = ""
    dec_cause_details = ""
    for accident_id, cases in dec_base_dict.items():
        dset = cases[cases['伤害程度']== '死亡'].reset_index(drop=True)
        if len(dset) == 1:
            for idx, row in dset.iterrows():
                location = row['事故地点']
                parsed_location = jio.parse_location(location)
                
                # 使用正则表达式提取道路的信息
                match = re.search(r'^(.*?[路道口])', parsed_location['detail'])
                if match:
                    road = match.group(0).replace(parsed_location['city'], '')
                    dec_road_details += f"{road}、"
                
                time_obj = parser.parse(str(row['事故发生时间']))
                dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{road}，因{row['违法行为']}违法行为，造成1人死亡。"
                dec_time_details += f"{time_obj.strftime('%H时%M分')}、"
        else:
            behaviors = ""
            for idx, row in dset.iterrows():
                behaviors += f"{row['违法行为']}、"
            location = row['事故地点']
            parsed_location = jio.parse_location(location)
            # 使用正则表达式提取道路的信息
            match = re.search(r'^(.*?[路道口])', parsed_location['detail'])
            if match:
                road = match.group(0).replace(parsed_location['city'], '')
                dec_road_details += f"{road}、"

            time_obj = parser.parse(str(row['事故发生时间']))
            dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{road}，因{behaviors[:-1]}违法行为，造成{len(dset)}人死亡。"
            dec_time_details += f"{time_obj.strftime('%H时%M分')}、"

    dec_time_details = dec_time_details[:-1]
    dec_road_details = dec_road_details[:-1]
    # dec_team_df = get_accidents_decedents_team_view(dataframe)

    dec_age_dict, _= get_decedents_age_cause_view(dataframe)
    if dec_acc_count == 1:
        for accident_id, cases in dec_age_dict.items():
            dec_age_details = f"事故双方为"
            dec_cause_details = f"事故原因为"
            prename = ""
            for idx, row in cases.iterrows():
                if row['伤害程度'] == '死亡':
                    prename = f"死者"
                else:
                    prename = f"当事人"   
                dec_age_details += f"{prename}{row['姓名']}，{row['性别']}，{row['年龄']}岁。"
                dec_cause_details += f"{row['违法行为'][:-1]},"
    else:
        count = 1
        for accident_id, cases in dec_age_dict.items():
            dec_age_details += f"事故{count}双方为"
            dec_cause_details += f"事故{count}原因为"
            prename = ""
            for idx, row in cases.iterrows():
                if row['伤害程度'] == '死亡':
                    prename = f"死者"
                else:
                    prename = f"当事人"
                dec_age_details += f"{prename}{row['姓名']}，{row['性别']}，{row['年龄']}岁。"
                dec_cause_details += f"{row['违法行为'][:-1]},"
            count += 1
    
    return dec_base_details[:-1], dec_time_details, dec_road_details, dec_age_details[:-1], dec_cause_details[:-1]