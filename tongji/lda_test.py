import pandas as pd
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import string
import gensim
from gensim import corpora
import matplotlib.pyplot as plt
import pyLDAvis.gensim_models as gensimvis
import pyLDAvis
import webbrowser

# 下载必要的nltk资源包（第一次运行时需要）
nltk.download([
    'punkt',
    'punkt_tab',  # 解决你的报错问题
    'stopwords'
])


# 1. 读取Excel文件第一个sheet的第一列数据
def read_excel_comments(file_path):
    df = pd.read_excel(file_path, sheet_name=0)  # 第一个sheet
    comments = df.iloc[:, 0].dropna().astype(str).tolist()  # 第一列，去除空值并转为字符串列表
    return comments


# 2. 文本预处理（适用于英文评论）
def preprocess(texts):
    stop_words = set(stopwords.words('english'))
    processed_texts = []
    for doc in texts:
        tokens = word_tokenize(doc.lower())  # 分词 + 小写
        words = [word for word in tokens if word.isalpha()]  # 去除非字母字符
        words = [w for w in words if w not in stop_words and len(w) > 2]  # 去停用词和太短的词
        processed_texts.append(words)
    return processed_texts


# 3. 构建LDA模型
def run_lda(processed_texts, num_topics):
    dictionary = corpora.Dictionary(processed_texts)
    corpus = [dictionary.doc2bow(text) for text in processed_texts]

    lda_model = gensim.models.LdaModel(
        corpus=corpus,
        id2word=dictionary,
        num_topics=num_topics,
        random_state=42,
        update_every=1,
        chunksize=10,
        passes=10,
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


# 5. 计算困惑度并寻找最佳主题数
def find_optimal_number_of_topics(processed_texts, start=2, end=10):
    coherence_values = []
    model_list = []
    for num_topics in range(start, end + 1):
        print(f"正在训练模型，主题数: {num_topics}")
        model, corpus, dictionary = run_lda(processed_texts, num_topics)
        model_list.append(model)
        coherence_values.append(-model.log_perplexity(corpus))  # 使用-log(perplexity)作为coherence值

    # 绘制困惑度图
    x = range(start, end + 1)
    plt.plot(x, coherence_values)
    plt.xlabel("Num Topics")
    plt.ylabel("-Log Perplexity")  # 负对数困惑度
    plt.legend(("Perplexity"), loc='best')
    plt.show()

    # 找到最佳主题数
    optimal_num_topics = coherence_values.index(max(coherence_values)) + start
    print(f"根据困惑度，推荐的主题数量为: {optimal_num_topics}")
    return model_list[optimal_num_topics - start], optimal_num_topics


# 6. 可视化LDA结果
def visualize_lda(lda_model, corpus, dictionary, output_file="lda_visual.html"):
    vis_data = gensimvis.prepare(lda_model, corpus, dictionary)
    pyLDAvis.save_html(vis_data, output_file)
    print(f"可视化结果已保存到: {output_file}")
    webbrowser.open(output_file)  # 自动打开浏览器查看


# 主函数
def main():
    file_path = "/Users/guangxin/PycharmProjects/pythonProject1/tongji/outputfile/outputResult.xlsx"  # 替换为你自己的路径
    print("正在读取Excel文件...")
    comments = read_excel_comments(file_path)

    print("正在预处理文本数据...")
    processed = preprocess(comments)

    print("正在寻找最佳主题数...")
    best_model, optimal_num_topics = find_optimal_number_of_topics(processed, start=2, end=10)

    print("\n提取出的最佳主题及关键词如下：")
    print_topics(best_model, num_words=10)

    print("正在生成可视化页面...")
    dictionary = corpora.Dictionary(processed)
    corpus = [dictionary.doc2bow(text) for text in processed]
    visualize_lda(best_model, corpus, dictionary)


if __name__ == "__main__":
    main()
