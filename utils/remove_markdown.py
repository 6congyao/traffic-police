def clean_markdown_file(file_path):
    """
    处理 Markdown 文件，去除语义字符和特定标签
    
    Args:
        file_path (str): Markdown 文件路径
    """
    try:
        # 读取文件内容
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. 删除 <think> 标签及其内容
        content = re.sub(r'<think>[\s\S]*?</think>', '', content)
        
        # 2. 去除 Markdown 语义字符
        # 移除注释
        content = re.sub(r'<!--.*?-->', '', content)
        # 移除标题标记 (#)
        content = re.sub(r'^#+\s*', '', content, flags=re.MULTILINE)
        # 移除粗体和斜体标记 (* _)
        content = re.sub(r'[*_]{1,2}(.*?)[*_]{1,2}', r'\1', content)
        # 移除链接 [text](url)
        content = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', content)
        # 移除代码块标记
        content = re.sub(r'```[\s\S]*?```', '', content)
        content = re.sub(r'`([^`]+)`', r'\1', content)
        # 移除所有无序列表标记（包括缩进的情况）
        content = re.sub(r'^\s*[-\*\+]\s+', '', content, flags=re.MULTILINE)
        # 移除水平分隔线
        content = re.sub(r'^-{3,}$|^_{3,}$|^\*{3,}$', '', content, flags=re.MULTILINE)
        
        # 3. 查找第一个数字列表项的位置
        match = re.search(r'（一）', content)
        if match:
            content = content[match.start():]
        
        # 4. 清理多余的空行，但保留基本换行结构
        # 将多个空行替换为两个换行
        content = re.sub(r'\n\s*\n', '\n', content)
        # 去除行首尾的空白字符
        content = '\n'.join(line.strip() for line in content.splitlines())
        
        # 5. 保存处理后的文本
        output_path = file_path.rsplit('.', 1)[0] + '_cleaned.txt'
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content.strip())
            
        print(f"处理完成，已保存到: {output_path}")
        
    except Exception as e:
        print(f"处理文件时发生错误: {str(e)}")

if __name__ == '__main__':
    import re
    # 处理指定的 Markdown 文件
    clean_markdown_file('testcases/example3.md')