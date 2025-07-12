import os
import time
from datetime import datetime
import pandas as pd


# 合并excel文件所有列数据
# 全局变量用于中断处理
save_on_exit = True


def signal_handler(sig, frame):
    """捕获Ctrl+C，进行优雅退出"""
    print("\n🛑 用户中断程序，提前结束。")
    exit(0)


def merge_excel_all_columns(input_folder, output_folder, output_filename="merged_output"):
    """
    合并指定文件夹下所有Excel文件第一个Sheet页的所有列数据。
    :param input_folder: 输入文件夹路径
    :param output_folder: 输出文件夹路径
    :param output_filename: 输出文件名
    """
    print("🔍 开始合并Excel文件的所有列数据...")

    # 创建输出文件夹（如果不存在）
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 获取所有 .xlsx 文件
    xlsx_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx") and not f.startswith("~$")]

    if not xlsx_files:
        print("❌ 没有可合并的Excel文件。")
        return

    all_data = []

    for filename in xlsx_files:
        file_path = os.path.join(input_folder, filename)
        print(f"📄 正在读取文件: {filename}")
        try:
            # 只读取第一个sheet
            df = pd.read_excel(file_path, sheet_name=0, header=None)
            # 转换为列表，便于后续合并
            all_data.append(df.values.tolist())
        except Exception as e:
            print(f"⚠️ 跳过文件 {filename}: {e}")

    if not all_data:
        print("❌ 没有提取到任何有效数据。")
        return

    # 找出最大列数
    max_columns = max(len(row) for data_block in all_data for row in data_block)

    # 合并所有数据
    merged_rows = []
    for data_block in all_data:
        for row in data_block:
            padded_row = row + [""] * (max_columns - len(row))  # 补齐空列
            merged_rows.append(padded_row)

    # 构建 DataFrame
    columns = [f"列{i + 1}" for i in range(max_columns)]
    result_df = pd.DataFrame(merged_rows, columns=columns)

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    output_filename = f"{output_filename}_{timestamp}.xlsx"

    output_path = os.path.join(output_folder, output_filename)

    try:
        result_df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"✅ 数据已成功保存至: {output_path}")
    except Exception as e:
        print(f"❌ 保存文件失败: {e}")


def main():
    import signal
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    start_time = time.time()
    print(f"🕒 程序启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    current_dir = os.path.dirname(os.path.abspath(__file__))
    input_folder = os.path.join(current_dir, '..', 'toBeMergedFolder')
    output_folder = os.path.join(current_dir, '..', 'outputfile')

    # 创建输入输出文件夹（如果不存在）
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print("📁 输入文件夹已自动创建，请将待处理的 .xlsx 文件放入以下目录后重新运行程序：")
        print(os.path.abspath(input_folder))
        return

    merge_excel_all_columns(input_folder, output_folder)

    end_time = time.time()
    print(f"\n🏁 程序执行结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ 总耗时：{end_time - start_time:.2f} 秒")


if __name__ == '__main__':
    main()
