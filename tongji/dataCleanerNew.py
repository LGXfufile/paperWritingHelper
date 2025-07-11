#   请在terminal中运行此脚本, pip install -r requirements.txt

import traceback

import pandas as pd
import re
from bs4 import BeautifulSoup
import os
from openai import OpenAI

import time  # 新增导入时间模块

# 初始化DeepSeek客户端
client = OpenAI(api_key="sk-b56d299a263d4570a59580b1082a262e", base_url="https://api.deepseek.com")

# 创建全局变量ds_count，记录处理的数据
global ds_count
ds_count = 0
# 是否开启DeepSeek,True为开启，False为关闭
ds_is_open = False

need_delete = []

# 无效词列表
sensitive_keywords = {
    '酷', '好帅', '赛博', '姐姐', '帅哥', '姐妹', '厉害', '宝宝', '联系', '手术', '松鼠桂鱼', '松鼠',
    '姐', '刘海', '刘亦菲', '买', '学生', '鞋', '老婆', '老公', '地铁', '支付宝', '卧槽', '美女', '超模', '感谢',
    '男朋友', '主播', '茶叶', '难吃', '好好吃', '特别好吃', '老鼠', '饭店', '餐厅', '女朋友', '你好', '醋鱼', '职场',
    '滤镜', '不同价格', '包装', '别吃', '不好吃', '啊啊啊啊啊啊', '价格', 'v信', '微信', '这一包', '这一点',
    '不太好吃', '什么软件', '礼盒', '自助餐', '吃不了', '倒掉', '姨妈巾', '这家店', '衣品',
    '快递', '特别好吃', '非常好喝', '糖醋排骨', '鱼是鱼', '佑圣观路', '电话', '垃圾食品', '杨梅',
    '笑死', '笑哭', '淘宝', '你俩', '礼貌拿图', '礼貌取图', '乌鲁木齐',
    '贵不贵', '网红', '又腥又酸', '又酸', '浪费', '鱼香肉丝', '这是哪家', '昨夜雨疏风骤', '想要'
}

CITY_KEYWORDS = {
    # 浙江省
    "建德", "寿昌", "海盐", "宁波", "衢州", "温州", "东阳", "淳安",

    # 广东省
    "惠州", "潮州", "雷州", "阳江",

    # 湖南省
    "衡阳", "耒阳", "津市", "郴州",

    # 四川省
    "自贡", "成都", "乐山", "宜宾",

    # 福建省
    "福州", "厦门", "泉州",

    # 江西省
    "南昌", "都昌", "上饶",

    # 河北省
    "石家庄", "保定",

    # 安徽省
    "阜阳", "颍州西湖",

    # 江苏省
    "扬州", "瘦西湖",

    # 其他省份
    "桂林", "天门", "武汉", "许昌", "西安", "济南", "大明湖", "大理", "沈阳",

    # 海外
    "山梨县", "河内",
}


def clean_comment(comment):
    if not isinstance(comment, str):
        return ""

    soup = BeautifulSoup(comment, "lxml")
    comment = soup.get_text()

    comment = re.sub(r'https?://\S+|www\.\S+', '', comment)
    comment = re.sub(r'@\S+', '', comment)
    comment = re.sub(r'([^\w\s])\1{2,}', r'\1', comment)
    comment = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s,.!?~、，。！？]', '', comment)

    return comment.strip()


def is_hangzhou_westlake_related(comment):
    global ds_count
    ds_count += 1
    start_timestamp = int(pd.Timestamp.now().timestamp())

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system",
             "content": "你是一位旅游评论分析助手，专门负责判断小红书上的评论是否与杭州西湖的游览体验相关。"},
            {"role": "user", "content": (
                "请根据以下标准判断一条来自小红书的评论是否与杭州西湖的游览体验相关：\n\n"
                "1. 如果评论直接描述了西湖的风景、文化、历史、活动、游览感受等内容，请返回 'true'。\n"
                "2. 如果评论涉及到西湖的游览路线、时间安排、交通接驳、停车、门票、开放时间等实用信息，请返回 'true'。\n"
                "3. 如果评论主要谈论的是个人外貌、美食推荐、非西湖相关的日常聊天、天气抱怨、不涉及西湖的地铁公交信息等，则返回 'false'。\n"
                "4. 如果评论内容混合了西湖相关信息和其他无关内容，请尝试提取核心意图，若主要关注点在西湖游览体验上，请返回 'true'。\n"
                "5. 特别注意：如果评论中的‘公交’‘地铁’‘打车’等词汇是为了了解如何更好地游览西湖或其周边景点，请视为相关内容并返回 'true'。\n\n"
                "请只输出一个单词：'true' 或 'false'，不要添加其他解释或格式。\n\n"
                "现在请判断以下评论是否与西湖游览体验相关：\n"
                f"{comment}"
            )}
        ],
        stream=False
    )

    end_timestamp = int(pd.Timestamp.now().timestamp())
    print(
        f'正在处理第{ds_count}条数据，评论内容：{comment[:30]}...，deepseek处理结果：{response.choices[0].message.content},耗时：' + str(
            end_timestamp - start_timestamp) + '秒')

    answer = response.choices[0].message.content.strip().lower()
    return answer == 'true'


def contains_ad(comment):
    if '优惠' in comment:
        return True
    if re.search(r'私信.*?([a-zA-Z0-9_\-]{5,}|[0-9]{6,})', comment):
        return True
    return False


def is_meaningful_chinese(text, min_chinese_chars=5):
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', text)
    if len(chinese_chars) < min_chinese_chars:
        return False

    common_gibberish = ['哈哈', '呵呵', '嘻嘻', '哦哦', '嗯嗯', '嘿嘿', '鹅鹅']
    for word in common_gibberish:
        if text.startswith(word * 2):  # 如 "哈哈哈哈哈哈"
            return False

    if len(text) > 3 * len(chinese_chars):
        return False

    return True


def has_repeated_pattern(text, repeat_threshold=3):
    words = re.split(r'\s+', text)
    from collections import Counter
    counter = Counter(words)
    for word, count in counter.items():
        if count >= repeat_threshold and len(word) >= 2:
            return True
    return False


def contains_city_keyword(comment):
    for city in CITY_KEYWORDS:
        if city in comment:
            return True
    return False


def load_exclusion_keywords(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        keywords = [line.strip() for line in f.readlines()]
    return set(keywords)


def filter_comments(input_file, output_file, exclusion_file=None):
    if os.path.isfile(output_file):
        os.remove(output_file)

    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    df = pd.read_excel(input_file, sheet_name=0)
    comments = df.iloc[:, 0].dropna().astype(str)

    filtered_comments = []
    exclusion_keywords = set()
    if exclusion_file and os.path.isfile(exclusion_file):
        exclusion_keywords = load_exclusion_keywords(exclusion_file)

    for comment in comments:
        cleaned = clean_comment(comment)

        if cleaned in filtered_comments:
            continue

        if not is_meaningful_chinese(cleaned):
            continue

        if has_repeated_pattern(cleaned):
            continue

        if exclusion_keywords:
            for word in exclusion_keywords:
                cleaned = re.sub(r'\b' + re.escape(word) + r'\b', '', cleaned)
            cleaned = re.sub(r'\s+', ' ', cleaned).strip()

        if len(cleaned) < 5 and '西湖' not in cleaned:
            continue

        if contains_city_keyword(cleaned):
            continue

        if any(kw in cleaned for kw in sensitive_keywords):
            continue

        if contains_ad(cleaned):
            continue

        if ds_is_open:
            if not is_hangzhou_westlake_related(cleaned):
                continue

        filtered_comments.append(cleaned)
        print(f"保留评论：{cleaned}")

    result_df = pd.DataFrame(filtered_comments, columns=["Cleaned Comments"])
    result_df.to_excel(output_file, index=False)
    print(f"处理完成，共保留 {len(filtered_comments)} 条有效评论，已保存至 {output_file}")


def main():
    start_time = time.time()
    current_dir = os.path.dirname(os.path.abspath(__file__))

    input_folder = os.path.join(current_dir, '..', 'inputFolder')
    output_folder = os.path.join(current_dir, '..', 'outputfile')
    exclude_path = os.path.join(current_dir, 'deletewords.txt')

    # 创建输入输出文件夹（如果不存在）
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print("📁 输入文件夹已自动创建，请将待处理的 .xlsx 文件放入以下目录后重新运行程序：")
        print(os.path.abspath(input_folder))
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 检查输入文件夹中是否有 .xlsx 文件
    xlsx_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]
    if not xlsx_files:
        print("⚠️ 输入文件夹中没有可处理的 .xlsx 文件。")
        print("📂 请将待处理的 .xlsx 文件放入以下目录：")
        print(os.path.abspath(input_folder))
        return

    print(f"✅ 程序开始时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time))}")
    print(f"🔍 正在处理文件夹: {os.path.abspath(input_folder)}")

    total_files = 0
    success_files = 0

    for filename in xlsx_files:
        total_files += 1
        input_path = os.path.join(input_folder, filename)
        output_filename = f"outputResult_{filename}"
        output_path = os.path.join(output_folder, output_filename)

        try:
            print(f"\n📦 正在处理文件: {filename}")
            filter_comments(input_path, output_path, exclude_path)
            success_files += 1
            print(f"✅ 文件 {filename} 处理完成。")
        except Exception as e:
            print(f"❌ 文件 {filename} 处理失败！错误信息：{str(e)}")
            traceback.print_exc()  # 打印完整错误堆栈

    end_time = time.time()
    elapsed_time = end_time - start_time

    print("\n" + "=" * 60)
    print(f"🏁 程序结束时间: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
    print(f"⏱️ 总共耗时: {elapsed_time:.2f} 秒")
    print(f"📄 共处理文件数量: {total_files}")
    print(f"✅ 成功处理文件数量: {success_files}")
    print(f"❌ 失败处理文件数量: {total_files - success_files}")
    print(f"📂 输出目录: {os.path.abspath(output_folder)}")
    print("=" * 60)


if __name__ == '__main__':
    main()
