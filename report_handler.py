import pandas as pd
from views import accident_statistics_overall_view as overall_view
from views import accident_statistics_decedents_base_view as dec_base_view
from views import accident_statistics_decedents_detail_view as dec_detail_view
from views import accident_statistics_casualties_base_view as cas_base_view
from views import accident_statistics_casualties_team_view as cas_team_view
from views import accident_statistics_casualties_road_view as cas_road_view
from views import accident_statistics_casualties_time_view as cas_time_view
from views import accident_statistics_casualties_age_view as cas_age_view
from views import accident_statistics_casualties_cause_view as cas_cause_view

from utils import word_template as wt
from dateutil import parser

def read_excel(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        
        print(f"正在读取文件: {file_path}")
        print(f"文件读取成功，总行数: {len(df)}")
            
        # 检查必要的列是否存在
        required_columns = ['事故编号', '来源', '身份证明号码', '伤害程度', '事故发生时间']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要的列: {col}")
            
        
        # 提取“事故发生时间”列，忽略空值
        time_column = df['事故发生时间'].dropna()

        # 计算起止区间
        start_time = time_column.min()
        end_time = time_column.max()

        return df, f"{parser.parse(str(start_time)).strftime('%Y年%m月%d日')}", f"{parser.parse(str(end_time)).strftime('%Y年%m月%d日')}"
        
                
    except FileNotFoundError:
        print(f"错误: 找不到文件 '{file_path}'", flush=True)
        exit(1)
    except pd.errors.EmptyDataError:
        print(f"错误: Excel文件 '{file_path}' 是空的", flush=True)
        exit(1)
    except Exception as e:
        print(f"处理数据时发生错误: {str(e)}", flush=True)
        import traceback
        print("详细错误信息:", flush=True)
        print(traceback.format_exc(), flush=True)
        exit(1)

def generate_report(input_file):
    dataframe, str_start_ts, str_end_ts = read_excel(input_file)
    template_path = 'templates/full_report.docx'

    acc_res = overall_view.get_accidents_overall_view(dataframe)
    id_res = overall_view.get_accidents_casualties_overall_view(dataframe)

    dec_base_res, dec_acc_count = dec_base_view.get_accidents_decedents_base_view(dataframe)

    replacements_dec = None
    if dec_acc_count > 0:
        dec_base_details, dec_time_details, dec_road_details, dec_team_details, dec_age_details, dec_cause_details = dec_detail_view.get_accidents_decedents_detail_view(dec_base_res, dec_acc_count, dataframe)
        replacements_dec = {
            '{$dec_a}': dec_acc_count,
            '{$dec_detail}': dec_base_details,
            '{$dec_team}': dec_team_details,
            '{$dec_road}': dec_road_details,
            '{$dec_time}': dec_time_details,
            '{$dec_age}': dec_age_details,
            '{$dec_cause}': dec_cause_details
        }
    else:
        template_path = 'templates/simple_report.docx'
        
    cas_base_res_a, _ = cas_base_view.get_casualties_base_view(dataframe)
    total_c_a = cas_base_res_a['一般事故'] + cas_base_res_a['简易事故']

    cas_team_df = cas_team_view.get_casualties_team_view(dataframe)
    cas_road_table_df = cas_road_view.get_casualties_road_view(dataframe)
    cas_time_table_df, cas_time_chart_df, cas_time_top3_list = cas_time_view.get_casualties_time_view(dataframe)
    cas_age_chart_df, cas_age_top3_list = cas_age_view.get_casualties_age_view(dataframe)
    cas_cause_top3, cas_cause_filtered_top3 = cas_cause_view.get_casualties_cause_view(dataframe)

    replacements = {
        '{$total_time_slot}': str_start_ts + ' - ' + str_end_ts,
        '{$total_ts}': str_start_ts + ' - ' + str_end_ts,
        '{$total_a}': acc_res['一般事故'] + acc_res['简易事故'],
        '{$total_c}': id_res.get(('一般事故', '轻伤'), 0) + id_res.get(('简易事故', '轻伤'), 0),
        '{$total_d}': id_res.get(('一般事故', '死亡'), 0),
        '{$simple_a}': acc_res['简易事故'],
        '{$simple_c}': id_res.get(('简易事故', '轻伤'), 0),
        '{$general_a}': acc_res['一般事故'],
        '{$general_c}': id_res.get(('一般事故', '轻伤'), 0),
        '{$general_d}': id_res.get(('一般事故', '死亡'), 0),
        '{$total_c_a}': total_c_a,
        '{$total_c_c}': id_res.get(('一般事故', '轻伤'), 0) + id_res.get(('简易事故', '轻伤'), 0),
        '{$simple_c_a}': cas_base_res_a['简易事故'],
        '{$simple_c_c}': id_res.get(('简易事故', '轻伤'), 0),
        '{$occupy_sa}': round((cas_base_res_a['简易事故'] / total_c_a) * 100, 1),
        '{$occupy_sc}': round((id_res.get(('简易事故', '轻伤'), 0) / (id_res.get(('一般事故', '轻伤'), 0) + id_res.get(('简易事故', '轻伤'), 0))) * 100, 1) if id_res.get(('简易事故', '轻伤'), 0) != 0 else 0.0,
        '{$general_c_a}': cas_base_res_a['一般事故'],
        '{$general_c_c}': id_res.get(('一般事故', '轻伤'), 0),
        '{$top1_cas_time}': cas_time_top3_list[0]['时段'],
        '{$top2_cas_time}': cas_time_top3_list[1]['时段'],
        '{$top3_cas_time}': cas_time_top3_list[2]['时段'],
        '{$occupy_c_t1}': cas_time_top3_list[0]['占比'],
        '{$occupy_c_t2}': cas_time_top3_list[1]['占比'],
        '{$occupy_c_t3}': cas_time_top3_list[2]['占比'],
        '{$total_c_age}': id_res.get(('一般事故', '轻伤'), 0) + id_res.get(('简易事故', '轻伤'), 0),
        '{$top1_cas_age}': cas_age_top3_list[0]['年龄段'],
        '{$top2_cas_age}': cas_age_top3_list[1]['年龄段'],
        '{$top3_cas_age}': cas_age_top3_list[2]['年龄段'],
        '{$top1_cas_count}': cas_age_top3_list[0]['受伤人数'],
        '{$top2_cas_count}': cas_age_top3_list[1]['受伤人数'],
        '{$top3_cas_count}': cas_age_top3_list[2]['受伤人数'],
        '{$occupy_c_a1}': cas_age_top3_list[0]['占比'],
        '{$occupy_c_a2}': cas_age_top3_list[1]['占比'],
        '{$occupy_c_a3}': cas_age_top3_list[2]['占比'],
        '{$top1_cause}': cas_cause_top3[0]['违法行为'],
        '{$top2_cause}': cas_cause_top3[1]['违法行为'],
        '{$top3_cause}': cas_cause_top3[2]['违法行为'],
        '{$top1_cause_count}': cas_cause_top3[0]['数量'],
        '{$top2_cause_count}': cas_cause_top3[1]['数量'],
        '{$top3_cause_count}': cas_cause_top3[2]['数量'],
        '{$top1_occupy_cause}': cas_cause_top3[0]['占比'],
        '{$top2_occupy_cause}': cas_cause_top3[1]['占比'],
        '{$top3_occupy_cause}': cas_cause_top3[2]['占比'],
        '{$top1_cas_cause}': cas_cause_filtered_top3[0]['违法行为'],
        '{$top2_cas_cause}': cas_cause_filtered_top3[1]['违法行为'],
        '{$top3_cas_cause}': cas_cause_filtered_top3[2]['违法行为']
    }

    if replacements_dec is not None:
        replacements.update(replacements_dec)

    table_list = {
        # 0: dec_team_df,
        0: cas_team_df,
        1: cas_road_table_df,
        2: cas_time_table_df
    }

    charts_list = {
        '{$chart_cas_time}': {
            'data': cas_time_chart_df,
            'x_column': '时段',
            'y_column': '受伤人数',
            'title': '本月伤人事故时段分布',
            'xlabel': '时间段',
            'ylabel': '受伤人数'
        },
        '{$chart_cas_age}': {
            'data': cas_age_chart_df,
            'x_column': '年龄段',
            'y_column': '受伤人数',
            'title': '本月伤人事故年龄段分布',
            'xlabel': '年龄段',
            'ylabel': '受伤人数'
        }
    }

    return wt.replace_template_variables(template_path, 'outputs/pre_generated_report.docx', replacements, table_list, charts_list)

def reinforce_report(input_file, contents_a, contents_s):
    replacements = {
        '{$ai_analysis}': contents_a,
        '{$ai_suggestion}': contents_s
    }

    return wt.replace_template_variables(input_file, 'outputs/generated_report.docx', replacements, None, None)
