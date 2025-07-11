import os


def get_file_extension(file_path):
    _, file_extension = os.path.splitext(file_path)
    return file_extension


# 示例用法
file_path = './getmmdname.md'
extension = get_file_extension(file_path)
print(f"The file extension is: {extension}")
