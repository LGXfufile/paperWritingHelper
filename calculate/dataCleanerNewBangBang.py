import os
import time
from datetime import datetime
from openai import OpenAI
import pandas as pd
from openpyxl import Workbook

# 配置参数
SAVE_EVERY_N_LINES = 5  # 每处理多少行保存一次
ds_is_open = True  # 是否启用DeepSeek接口

# 初始化DeepSeek客户端
# 清洗输出文件格式： 旅游要素	评论词	情感态度	评论索引
client = OpenAI(api_key="sk-b56d299a263d4570a59580b1082a262e", base_url="https://api.deepseek.com")

# 全局变量用于中断处理
wb_global = None
ws_global = None
current_file_global = None


def call_deepseek_api(comment_index, comment):
    """
    调用DeepSeek API分析评论，提取旅游要素、评论词和情感态度。
    返回列表：[{"element": "", "comment": "", "sentiment": ""}]
    """
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "你是一位旅游评论分析助手，专门负责从评论中抽取旅游要素（如景点、住宿、餐饮等）、"
                        "评论词，并判断作者对每个名词的情感态度（正面、负面、中性）。\n\n"
                        "要求：\n"
                        "1. 旅游要素，不超过4个字，例如景点，文物，参观学习，时间，配置设施，讲解人员，景区运营这种都属于旅游要素。\n"
                        "2. 评论词要简短，直接来自原文或精准概括，不超过6个字。\n"
                        "3. 若无明确评论词，请填写“无”。\n"
                        "4. 情感态度只能是“正面”、“负面”或“中性”。\n"
                        "5. 返回格式为每行为一组三列内容，逗号分隔：\n"
                        "   旅游要素,评论词,情感态度\n"
                        "示例：\n"
                        "景点,景色非常美,正面\n"
                        "住宿,房间很干净,正面"
                    )
                },
                {"role": "user", "content": f"请分析以下评论：\n{comment}"}
            ],
            stream=False
        )

        results_text = response.choices[0].message.content.strip()
        if not results_text:
            return []

        parsed_results = []
        for line in results_text.split('\n'):
            parts = [p.strip() for p in line.split(',', 2)]
            if len(parts) == 3:
                element, comment_word, sentiment = parts
                parsed_results.append({
                    "element": element,
                    "comment": comment_word or "无",
                    "sentiment": sentiment
                })
        return parsed_results

    except Exception as e:
        print(f"⚠️ API调用失败 (评论 {comment_index + 1}): {e}")
        return []


def save_and_reset_workbook(wb, ws, output_folder, base_filename, file_counter):
    """保存当前工作簿并重置"""
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    output_filename = f"{base_filename}_part{file_counter}_{timestamp}.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    try:
        wb.save(output_path)
        print(f"✅ 已保存部分结果到: {output_path}")

        # 创建新的工作簿并添加表头
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.append(["旅游要素", "评论词", "情感态度", "评论索引"])
        return new_wb, new_ws, output_filename

    except Exception as e:
        print(f"❌ 保存文件失败: {e}")
        return wb, ws, None


def signal_handler(sig, frame):
    """捕获Ctrl+C，进行优雅退出"""
    global wb_global, ws_global, current_file_global
    print("\n🛑 用户中断程序，正在保存当前数据...")
    if wb_global and ws_global and current_file_global:
        try:
            wb_global.save(current_file_global)
            print(f"✅ 当前缓存数据已保存至: {current_file_global}")
        except Exception as e:
            print(f"❌ 保存缓存数据失败: {e}")
    exit(0)


def main():
    import signal
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    start_time = time.time()
    print(f"🕒 程序启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(current_dir, '..', 'inputFolder')
    output_folder = os.path.join(current_dir, '..', 'outputfile')

    # 创建输入输出文件夹（如果不存在）
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print("📁 输入文件夹已自动创建，请将待处理的 .xlsx 文件放入以下目录后重新运行程序：")
        print(os.path.abspath(input_folder))
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print("📁 输出文件夹已创建。")

    print("\n🔍 开始处理Excel文件...\n")

    xlsx_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx")]

    if not xlsx_files:
        print("❌ 输入文件夹中没有可处理的 .xlsx 文件。")
        return

    global wb_global, ws_global, current_file_global

    for filename in xlsx_files:
        input_path = os.path.join(input_folder, filename)
        print(f"📄 正在处理文件: {filename}")

        try:
            df = pd.read_excel(input_path, sheet_name=0)
            comments = df.iloc[:, 0].dropna().astype(str)
        except Exception as e:
            print(f"❌ 无法读取文件 {filename}: {e}")
            continue

        base_filename = os.path.splitext(os.path.basename(filename))[0]

        # 初始化第一个工作簿
        wb = Workbook()
        ws = wb.active
        ws.append(["旅游要素", "评论词", "情感态度", "评论索引"])

        file_counter = 1
        current_line_count = 0
        wb_global, ws_global = wb, ws
        current_file_global = ""

        for idx, comment in enumerate(comments):
            start = time.time()
            print(f"💬 处理第 {idx + 1} 条评论...")

            if ds_is_open:
                results = call_deepseek_api(idx, comment)
            else:
                results = []

            if not results:
                ws.append(["无", "无", "无", str(idx + 1)])
                current_line_count += 1
            else:
                for result in results:
                    row = [
                        result["element"],
                        result["comment"],
                        result["sentiment"],
                        str(idx + 1)
                    ]
                    ws.append(row)
                    current_line_count += 1

            end = time.time()
            print(f"⏱️ 第 {idx + 1} 条评论处理完成，耗时 {end - start:.2f} 秒")

            if (idx + 1) % SAVE_EVERY_N_LINES == 0:
                wb, ws, new_file = save_and_reset_workbook(wb, ws, output_folder, base_filename, file_counter)
                current_file_global = new_file
                file_counter += 1

        # 保存剩余未保存的数据
        if current_line_count > 0:
            wb, ws, new_file = save_and_reset_workbook(wb, ws, output_folder, base_filename, file_counter)
            current_file_global = new_file

    end_time = time.time()
    print(f"\n🏁 程序执行结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📊 结果已保存至: {os.path.abspath(output_folder)}")
    print(f"⏱️ 总耗时：{end_time - start_time:.2f} 秒")


if __name__ == '__main__':
    print(f"✅ 程序每处理{SAVE_EVERY_N_LINES}行，就生成一个新的文件，并保存到指定的输出目录下.")
    main()
