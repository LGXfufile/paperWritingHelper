import pandas as pd
import jieba
from gensim import corpora, models
import matplotlib.pyplot as plt
import pyLDAvis.gensim_models as gensimvis
import pyLDAvis
import webbrowser
from wordcloud import WordCloud
import time
import warnings
import json
import numpy as np
import os

# 忽略警告
warnings.filterwarnings("ignore", category=RuntimeWarning)

# 设置中文字体以避免词云乱码
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False  # 显示中文标签


# 自定义 JSON Encoder，处理 numpy 数据类型
class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.int64, np.int32)):
            return int(obj)
        elif isinstance(obj, (np.float64, np.float32)):
            return float(obj)
        elif isinstance(obj, np.bool_):
            return bool(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, complex):
            return str(obj)
        return super().default(obj)


# 1. 读取Excel文件第一个sheet的第一列数据
def read_excel_comments(file_path):
    df = pd.read_excel(file_path, sheet_name=0)  # 第一个sheet
    comments = df.iloc[:, 0].dropna().astype(str).tolist()  # 第一列，去除空值并转为字符串列表
    return comments


# 2. 文本预处理（适用于中文评论）
def preprocess(texts):
    # 加载中文停用词表（你可以替换为你自己的停用词文件路径）
    stop_words = set()
    with open("chinese_stopwords.txt", "r", encoding="utf-8") as f:
        for line in f:
            stop_words.add(line.strip())

    processed_texts = []
    for doc in texts:
        tokens = jieba.lcut(doc.lower())  # 中文分词
        words = [word for word in tokens if word.strip() and word not in stop_words and len(word.strip()) > 1]
        processed_texts.append(words)
    return processed_texts


# 3. 构建LDA模型
def run_lda(processed_texts, num_topics):
    dictionary = corpora.Dictionary(processed_texts)
    corpus = [dictionary.doc2bow(text) for text in processed_texts]

    lda_model = models.LdaModel(
        corpus=corpus,
        id2word=dictionary,
        num_topics=num_topics,
        random_state=42,
        update_every=1,
        chunksize=10,
        passes=5,
        alpha='auto',
        per_word_topics=True
    )

    return lda_model, corpus, dictionary


# 4. 打印主题关键词
def print_topics(lda_model, num_words=10):
    topics = lda_model.print_topics(num_words=num_words)
    for idx, topic in enumerate(topics):
        print(f"\nTopic {idx}:")
        print(topic)


# 5. 寻找最佳主题数
def find_optimal_number_of_topics(processed_texts, start=2, end=9):  # 优化为 2~9
    coherence_values = []
    model_list = []

    for num_topics in range(start, end + 1):
        print(f"⏳ 正在训练模型，主题数: {num_topics}")
        start_time = time.time()

        try:
            model, corpus, dictionary = run_lda(processed_texts, num_topics)
            perplexity = -model.log_perplexity(corpus)
            coherence_values.append(perplexity)
            model_list.append((model, corpus, dictionary))
        except Exception as e:
            print(f"⚠️ 主题数 {num_topics} 训练失败: {e}")
            continue

        elapsed = time.time() - start_time
        print(f"✅ 主题数 {num_topics} 训练完成，耗时: {elapsed:.2f}s")

    x = range(start, end + 1)
    plt.figure()
    plt.plot(x, coherence_values, marker='o')
    plt.xlabel("Num Topics")
    plt.ylabel("-Log Perplexity")
    plt.title("Perplexity vs Number of Topics")
    plt.grid(True)
    plt.savefig("topic_perplexity_curve.png")  # 自动保存图像
    plt.close()

    optimal_num_topics = x[coherence_values.index(max(coherence_values))]
    best_model_idx = optimal_num_topics - start
    print(f"✅ 推荐的最佳主题数为: {optimal_num_topics}")
    return model_list[best_model_idx][0], model_list[best_model_idx][1], model_list[best_model_idx][
        2], optimal_num_topics


# 6. 可视化LDA结果
def visualize_lda(lda_model, corpus, dictionary, output_file="lda_visual.html"):
    print("🌐 正在准备可视化数据...")
    vis_data = gensimvis.prepare(lda_model, corpus, dictionary, mds="mmds")
    print("🌐 正在保存可视化结果...")

    html = pyLDAvis.prepared_data_to_html(vis_data)
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"🌐 可视化结果已保存到: {output_file}")
    # webbrowser.open(output_file)


# 7. 生成词云（支持中文）
def generate_word_clouds(lda_model, num_topics=9, max_words=40, font_path="/Users/guangxin/PycharmProjects/pythonProject1/tongji/SourceHanSans.ttf"):
    # 如果使用非系统字体，请指定字体文件路径，例如：font_path="fonts/simhei.ttf"
    for i in range(num_topics):
        print(f"🎨 正在生成主题 {i} 的词云...")
        topic_terms = lda_model.show_topic(i, topn=max_words)
        word_freq = {term: freq for term, freq in topic_terms}
        wordcloud = WordCloud(width=800, height=400,
                              background_color='white',
                              font_path=font_path,  # 指定中文字体
                              max_font_size=100,
                              max_words=100,
                              ).generate_from_frequencies(word_freq)
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis("off")
        plt.title(f'主题 {i}')
        plt.savefig(f'topic_{i}_wordcloud.png')  # 保存图片
        plt.close()


# 主函数
def main():
    file_path = "/Users/guangxin/PycharmProjects/pythonProject1/tongji/outputfile/outputResult.xlsx"  # 替换为你自己的路径
    print("🚀 正在读取Excel文件...")
    comments = read_excel_comments(file_path)

    print("🧹 正在预处理文本数据...")
    processed = preprocess(comments)

    print("🔍 正在寻找最佳主题数...")
    best_model, corpus, dictionary, optimal_num_topics = find_optimal_number_of_topics(processed, start=2, end=9)

    print("\n📊 提取出的最佳主题及关键词如下：")
    print_topics(best_model, num_words=10)

    print("🌐 正在生成可视化页面...")
    visualize_lda(best_model, corpus, dictionary)

    print("☁️ 正在生成词云...")
    generate_word_clouds(best_model, num_topics=optimal_num_topics, max_words=40)

    print("✅ 全部任务已完成！")


if __name__ == "__main__":
    main()
