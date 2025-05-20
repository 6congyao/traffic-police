import pandas as pd
import jionlp as jio
from datetime import datetime

def analyze_decedents_age_cause_view():
    try:
        # 读取Excel文件
        df = pd.read_excel('test.xlsx')
        
        # 筛选出伤害程度为"死亡"的数据，获取事故编号集合
        death_cases = df[df['伤害程度'] == '死亡']
        dset = set(death_cases['事故编号'].unique())
        
        if len(dset) == 0:
            print("没有找到死亡案例。")
            return
        
        # 在原始数据中筛选出事故编号在dset中的所有记录
        related_cases = df[df['事故编号'].isin(dset)].drop_duplicates(subset=['身份证明号码'], keep='first')
        
        # 计算年龄和性别
        age_sex = related_cases['身份证明号码'].apply(parse_id_card)
        related_cases['年龄'] = age_sex.apply(lambda x: x['年龄'])
        related_cases['性别'] = age_sex.apply(lambda x: x['性别'])

        # 创建以事故编号为key的字典
        cases_dict = {}
        for accident_id in dset:
            cases_dict[accident_id] = related_cases[related_cases['事故编号'] == accident_id]
        
        # print("\n与死亡案例相关的所有记录：")
        # print(f"找到 {len(dset)} 个事故，共 {len(related_cases)} 条记录")
        # print("=" * 100)
        
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 30)
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        
        # 为每个事故编号输出表格形式的数据
        for accident_id, cases in cases_dict.items():
            print(f"\n事故编号: {accident_id} (涉及 {len(cases)} 人)")
            print("-" * 120)
            # 重置索引并删除多余的事故编号列
            display_df = cases.reset_index(drop=True)
            display_df = display_df.drop(columns=['事故编号'])
            
            # 调整列宽和显示格式
            print(display_df.to_string(
                index=True,
                justify='left',
                col_space={
                    '姓名': 8,
                    '身份证明号码': 20,
                    '年龄': 6,
                    '性别': 6,
                    '事故责任': 8,
                    '号牌号码': 10,
                    '号牌种类': 10,
                    '车辆使用性质': 12,
                    '事故发生时间': 20,
                    '事故地点': 30,
                    '所属中队': 15,
                    '违法行为': 30,
                    '来源': 8,
                    '伤害程度': 8
                }
            ))
            print("=" * 120)
    
    except FileNotFoundError:
        print("错误：未找到Excel文件 'test.xlsx'")
    except Exception as e:
        print(f"错误：处理数据时出现异常 - {str(e)}")

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

if __name__ == '__main__':
    analyze_decedents_age_cause_view()
