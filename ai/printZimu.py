
from string import ascii_lowercase

def generate_letter_combinations():
    """生成从a-z, aa-zz的所有字母组合"""
    result = []
    # 先添加单个字母
    result.extend(list(ascii_lowercase))
    # 再添加双字母组合
    for first in ascii_lowercase:
        for second in ascii_lowercase:
            result.append(f"{first}{second}")
    return result

# 示例调用
combinations = generate_letter_combinations()
print(combinations)

print(f"len(combinations)={len(combinations)}")