import pandas as pd
from views import accident_statistics_overall_view as overall_view
from views import accident_statistics_decedents_base_view as dec_base_view
from views import accident_statistics_decedents_team_view as dec_team_view
from views import accident_statistics_decedents_age_cause_view as dec_age_view
from views import accident_statistics_casualties_base_view as cas_base_view
from views import accident_statistics_casualties_team_view as cas_team_view
from views import accident_statistics_casualties_time_view as cas_time_view
from utils import word_template as wt
from dateutil import parser

def read_excel(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        print(f"正在读取文件: {file_path}")
        print(f"文件读取成功，总行数: {len(df)}")
            
        # 检查必要的列是否存在
        required_columns = ['事故编号', '来源', '身份证明号码', '伤害程度']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要的列: {col}")
        
        return df
                
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{file_path}'", flush=True)
    except pd.errors.EmptyDataError:
        print(f"错误: Excel文件 '{file_path}' 是空的", flush=True)
    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)

if __name__ == '__main__':
    dataframe = read_excel('test.xlsx')

    acc_res = overall_view.get_accidents_overall_view(dataframe)
    id_res = overall_view.get_accidents_casualties_overall_view(dataframe)

    dec_base_res, dec_acc_count = dec_base_view.get_accidents_decedents_base_view(dataframe)

    dec_base_details = ""
    dec_time_details = ""
    dec_age_details = ""
    dec_cause_details = ""
    if dec_acc_count > 0:
        # 亡人事故基本情况
        for accident_id, cases in dec_base_res.items():
            dset = cases[cases['伤害程度']== '死亡'].reset_index(drop=True)
            if len(dset) == 1:
                for idx, row in dset.iterrows():
                    time_obj = parser.parse(str(row['事故发生时间']))
                    dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{row['事故地点']}，因{row['违法行为']}违法行为，造成1人死亡。"
                    dec_time_details += f"{time_obj.strftime('%H时%M分')}、"
            else:
                behaviors = ""
                for idx, row in dset.iterrows():
                    behaviors += f"{row['违法行为']}、"
                time_obj = parser.parse(str(row['事故发生时间']))
                dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{row['事故地点']}，因{behaviors[:-1]}违法行为，造成{len(dset)}人死亡。"
                dec_time_details += f"{time_obj.strftime('%H时%M分')}、"
    
        dec_time_details = dec_time_details[:-1]
        dec_team_df = dec_team_view.get_accidents_decedents_team_view(dataframe)

        dec_age_dict, _= dec_age_view.get_decedents_age_cause_view(dataframe)
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
        
    cas_base_res_a, _ = cas_base_view.get_casualties_base_view(dataframe)
    total_c_a = cas_base_res_a['一般事故'] + cas_base_res_a['简易事故']

    cas_team_df = cas_team_view.get_casualties_team_view(dataframe)
    cas_time_table_df, cas_time_chart_df, cas_time_top3_list = cas_time_view.get_casualties_time_view(dataframe)

    replacements = {
        '{$total_a}': acc_res['一般事故'] + acc_res['简易事故'],
        '{$total_c}': id_res[('一般事故', '轻伤')] + id_res[('简易事故', '轻伤')],
        '{$total_d}': id_res[('一般事故', '死亡')],
        '{$simple_a}': acc_res['简易事故'],
        '{$simple_c}': id_res[('简易事故', '轻伤')],
        '{$general_a}': acc_res['一般事故'],
        '{$general_c}': id_res[('一般事故', '轻伤')],
        '{$general_d}': id_res[('一般事故', '死亡')],
        '{$dec_a}': dec_acc_count,
        '{$dec_detail}': dec_base_details,
        '{$dec_time}': dec_time_details,
        '{$dec_age}': dec_age_details,
        '{$dec_cause}': dec_cause_details[:-1]+'。',
        '{$total_c_a}': total_c_a,
        '{$total_c_c}': id_res[('一般事故', '轻伤')] + id_res[('简易事故', '轻伤')],
        '{$simple_c_a}': cas_base_res_a['简易事故'],
        '{$simple_c_c}': id_res[('简易事故', '轻伤')],
        '{$occupy_sa}': round((cas_base_res_a['简易事故'] / total_c_a) * 100, 1),
        '{$occupy_sc}': round((id_res[('简易事故', '轻伤')] / (id_res[('一般事故', '轻伤')] + id_res[('简易事故', '轻伤')])) * 100, 1),
        '{$general_c_a}': cas_base_res_a['一般事故'],
        '{$general_c_c}': id_res[('一般事故', '轻伤')],
        '{$top1_cas_time}': cas_time_top3_list[0]['时段'],
        '{$top2_cas_time}': cas_time_top3_list[1]['时段'],
        '{$top3_cas_time}': cas_time_top3_list[2]['时段'],
        '{$occupy_c_t1}': cas_time_top3_list[0]['占比'],
        '{$occupy_c_t2}': cas_time_top3_list[1]['占比'],
        '{$occupy_c_t3}': cas_time_top3_list[2]['占比']
    }

    table_list = {
        0: dec_team_df,
        1: cas_team_df,
        3: cas_time_table_df
    }

    charts_list = {
        '{$chart_cas_time}': {
            'data': cas_time_chart_df,
            'x_column': '时段',
            'y_column': '受伤人数',
            'title': '本月伤人事故时段分布',
            'xlabel': '时间段',
            'ylabel': '受伤人数'
        }
    }

    wt.replace_template_variables('templates/monthly_report.docx', 'outputs/monthly_report_filled.docx', replacements, table_list, charts_list)


