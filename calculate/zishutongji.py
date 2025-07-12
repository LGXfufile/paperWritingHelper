import openpyxl


def count_chinese_characters_in_excel(file_path):
    try:
        # 加载 Excel 文件
        workbook = openpyxl.load_workbook(file_path)

        # 获取第一个 sheet
        sheet = workbook.worksheets[0]

        total_char_count = 0

        # 遍历第一列的所有行
        for row in sheet.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            if cell.value and isinstance(cell.value, str):
                # 统计每个单元格中的字符数（包括中文、标点、空格等）
                total_char_count += len(cell.value)
        return total_char_count
    except Exception as e:
        print(f"读取文件出错: {e}")
        return None


# 使用示例
if __name__ == "__main__":
    file_path = "/Users/guangxin/PycharmProjects/pythonProject1/calculate/yijipinglun.xlsx"  # 替换为你的文件路径
    char_count = count_chinese_characters_in_excel(file_path)
    if char_count is not None:
        print(f"总字数（含标点、空格等）: {char_count}")
