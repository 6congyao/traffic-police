import pandas as pd

def analyze_accidents_decedents_base_view():
    """
    分析死亡数据统计：
    1. 按身份证明号码去重统计死亡人数
    2. 按事故编号去重统计死亡事故数
    """
    try:
        # 读取Excel文件
        df = pd.read_excel('testcases/test.xlsx')
        
        # 过滤出死亡数据，忽略NaN值
        death_df = df[df['伤害程度'] == '死亡'].dropna(subset=['伤害程度'])
        
        # 按身份证明号码去重统计
        unique_id_count = len(death_df['身份证明号码'].dropna().unique())
        
        # 按事故编号去重统计
        unique_accident_count = len(death_df['事故编号'].dropna().unique())
        
        # 输出结果
        print(f"\n统计结果:")
        print(f"按身份证明号码统计的死亡人数: {unique_id_count}")
        print(f"按事故编号统计的死亡事故数: {unique_accident_count}")
        
    except FileNotFoundError:
        print("错误：未找到Excel文件")
    except Exception as e:
        print(f"错误：{str(e)}")

def get_accidents_decedents_base_view(df):
    # 筛选出伤害程度为"死亡"的数据，获取事故编号集合
    decedents_df = df[df['伤害程度'] == '死亡']
    daset = set(decedents_df['事故编号'].unique())

    if len(daset) == 0:
        return {}, 0
    
    # 在原始数据中筛选出事故编号在dset中的所有记录
    related_df = df[df['事故编号'].isin(daset)].drop_duplicates(subset=['身份证明号码'], keep='first')

    # 创建以事故编号为key的字典
    related_dict = {}
    for accident_id in daset:
        related_dict[accident_id] = related_df[related_df['事故编号'] == accident_id]
    
    # print("=" * 100)
    # print("\n与死亡案例相关的所有记录：")
    # print(f"找到 {len(daset)} 个事故，共 {len(related_df)} 条记录")

    # 为每个事故编号输出表格形式的数据
    # for accident_id, cases in related_dict.items():
    #     print(f"\n事故编号: {accident_id} (涉及 {len(cases)} 人)")
    #     print("-" * 120)
    #     # 重置索引并删除多余的事故编号列
    #     display_df = cases.reset_index(drop=True)
    #     display_df = display_df.drop(columns=['事故编号'])
        
    #     # 调整列宽和显示格式
    #     print(display_df.to_string(
    #         index=True,
    #         justify='left',
    #         col_space={
    #             '姓名': 8,
    #             '身份证明号码': 20,
    #             '事故责任': 8,
    #             '号牌号码': 10,
    #             '号牌种类': 10,
    #             '车辆使用性质': 12,
    #             '事故发生时间': 20,
    #             '事故地点': 30,
    #             '所属中队': 15,
    #             '违法行为': 30,
    #             '来源': 8,
    #             '伤害程度': 8
    #         }
    #     ))
    #     print("=" * 120)

    return related_dict, len(daset)

if __name__ == "__main__":
    analyze_accidents_decedents_base_view()
