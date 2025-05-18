from docx import Document

def replace_template_variables(template_path, output_path, replacements):
    """
    替换Word模板中的变量
    :param template_path: 模板文件路径
    :param output_path: 输出文件路径
    :param replacements: 要替换的变量字典
    """
    doc = Document(template_path)
    remaining_keys = set(replacements.keys())
    
    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        if not remaining_keys:  # 如果所有变量都已替换完成，直接退出
            break
            
        for key in list(remaining_keys):
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, str(replacements[key]))
                remaining_keys.remove(key)
    
    # 如果还有未替换的变量，继续在表格中查找
    if remaining_keys:
        for table in doc.tables:
            if not remaining_keys:  # 如果所有变量都已替换完成，直接退出
                break
                
            for row in table.rows:
                if not remaining_keys:
                    break
                    
                for cell in row.cells:
                    if not remaining_keys:
                        break
                        
                    for paragraph in cell.paragraphs:
                        if not remaining_keys:
                            break
                            
                        for key in list(remaining_keys):
                            if key in paragraph.text:
                                for run in paragraph.runs:
                                    run.text = run.text.replace(key, str(replacements[key]))
                                remaining_keys.remove(key)
    
    # 保存修改后的文档
    doc.save(output_path)

# 示例使用
if __name__ == '__main__':
    template_path = 'templates/monthly_report.docx'
    output_path = 'outputs/monthly_report_filled.docx'
    
    replacements = {
        '{$total_a}': '3',
        '{$total_c}': '123',
        '{$total_d}': '456'
    }
    
    replace_template_variables(template_path, output_path, replacements)
