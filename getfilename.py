import os
import re


def get_file_extension(file_path):
    _, file_extension = os.path.splitext(file_path)
    return file_extension


def extract_content(md_file):
    result = []
    with open(md_file, 'r', encoding='utf-8') as file:
        text = file.read()
        # 使用正则表达式匹配括号及其前面的两个汉字
        pattern = r'(\b[\u4e00-\u9fa5]{2})\s*\((\d{4})\)'
        matches = re.findall(pattern, text)

        for match in matches:
            # 构建完整的字符串形式
            content = match[0] + '(' + match[1] + ')'
            result.append(content)
    return result


def save_to_txt(content_list, txt_file):
    if os.path.exists(txt_file):
        os.remove(txt_file)
    with open(txt_file, 'w', encoding='utf-8') as f:
        for content in content_list:
            f.write(content + '\n')


# Markdown 文件路径
md_file = './getmmdname.md'
# 输出 TXT 文件路径
txt_file = 'output.txt'


# 提取内容
extracted_content = extract_content(md_file)
save_to_txt(extracted_content, txt_file)

print('check end ~')
