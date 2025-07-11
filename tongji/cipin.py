import pandas as pd
import jieba
from collections import Counter
import sys

# 设置默认编码为 utf-8（防止中文乱码）
sys.setdefaultencoding = 'utf-8'

# 常见无意义停用词列表（可根据实际内容扩展）
STOPWORDS = {
    '的', '了', '是', '我', '我们', '在', '有', '和', '就', '这', '那',
    '都', '也', '上', '下', '着', '啊', '哦', '吧', '嘛', '呢', '呀',
    '他们', '她们', '它们', '你', '您', '他', '她', '它', '一个', '一些', '但是'
}


def custom_tokenize(text, custom_keywords):
    """
    使用 jieba 分词，并保留自定义关键词不被拆分。
    同时过滤掉长度 < 2 的词和停用词。
    """
    # 将自定义关键词加入分词词典
    for word in custom_keywords:
        jieba.add_word(word)

    # 使用全模式分词（更少拆分），也可以使用 lcut 精确模式
    words = jieba.lcut(text)

    # 过滤：只保留 >=2 字、不在停用词中的词
    filtered_words = [
        word for word in words
        if len(word) >= 2 and word not in STOPWORDS
    ]
    return filtered_words


def count_keywords_in_excel(file_path, custom_keywords, output_path="output_keyword_freq.xlsx"):
    """
    统计 Excel 第一个 Sheet 第一列中的关键词频率。

    参数：
    - file_path: 输入的 .xlsx 文件路径
    - custom_keywords: 自定义不拆分的关键词列表
    - output_path: 输出文件路径
    """
    # 读取 Excel 文件
    df = pd.read_excel(file_path, sheet_name=0)  # 只读取第一个 sheet

    # 获取第一列的内容并转为字符串列表
    texts = df.iloc[:, 0].astype(str).tolist()

    # 合并所有文本
    all_text = '\n'.join(texts)

    # 分词并统计词频
    words = custom_tokenize(all_text, custom_keywords)
    word_counts = Counter(words)

    # 排序：按词频降序
    sorted_word_counts = sorted(word_counts.items(), key=lambda x: x[1], reverse=True)

    # 构建 DataFrame
    result_df = pd.DataFrame(sorted_word_counts, columns=["关键词", "词频"])

    # 写入 Excel
    result_df.to_excel(output_path, index=False, engine='openpyxl')

    print(f"词频统计完成，结果已保存至 {output_path}")


# 示例调用
if __name__ == "__main__":
    # input_file = "/Users/guangxin/PycharmProjects/pythonProject1/tongji/yijipinglun.xlsx"  # 替换为你的文件路径

    input_file = "/Users/guangxin/PycharmProjects/pythonProject1/tongji/西湖_帖子_词频.xlsx"
    custom_keywords = ["人工智能", "大数据", "机器学习", "深度学习", "自然语言处理"]  # 自定义关键词
    count_keywords_in_excel(input_file, custom_keywords)
