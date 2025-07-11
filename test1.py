import socket

def get_local_ip():
    # 创建一个UDP套接字
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
        try:
            # 连接到一个公共的DNS服务器，这里使用8.8.8.8，实际并不发送数据
            s.connect(('8.8.8.8', 80))
            # 获取连接信息中的本地IP地址
            local_ip = s.getsockname()[0]
            return local_ip
        except Exception as e:
            print(f"Error getting IP address: {e}")
            return None

# 调用函数并打印结果
local_ip_address = get_local_ip()
if local_ip_address:
    print(f"Local IP Address: {local_ip_address}")
else:
    print("Could not retrieve the local IP address.")
