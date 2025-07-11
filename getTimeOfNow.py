from datetime import datetime

# 获取当前的日期和时间
now = datetime.now()

# 打印当前的日期和时间
# print("Current date and time:", now)

# 如果需要，你可以格式化日期和时间
formatted_now = now.strftime("%Y-%m-%d %H:%M:%S")
print("Formatted date and time:", formatted_now)
