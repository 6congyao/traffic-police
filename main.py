import pandas as pd
from views import accident_statistics_overall_view as overall_view
from views import accident_statistics_decedents_base_view as dec_base_view
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
    if dec_acc_count > 0:
        # 亡人事故基本情况
        for accident_id, cases in dec_base_res.items():
            dset = cases[cases['伤害程度']== '死亡'].reset_index(drop=True)
            if len(dset) == 1:
                for idx, row in dset.iterrows():
                    time_obj = parser.parse(str(row['事故发生时间']))
                    dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{row['事故地点']}，因{row['违法行为']}违法行为，造成1人死亡。"
            else:
                behaviors = ""
                for idx, row in dset.iterrows():
                    behaviors += f"{row['违法行为']}、"
                time_obj = parser.parse(str(row['事故发生时间']))
                dec_base_details += f"{time_obj.strftime('%m月%d日%H时%M分')}, 在{row['事故地点']}，因{behaviors[:-1]}违法行为，造成{len(dset)}人死亡。"
            
            print(f"\n事故编号: {accident_id} (涉及 {len(cases)} 人)")
            print("-" * 100)
    
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
        '{$dec_detail}': dec_base_details

    }

    wt.replace_template_variables('templates/monthly_report.docx', 'outputs/monthly_report_filled.docx', replacements)


