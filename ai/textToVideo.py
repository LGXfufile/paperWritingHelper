import os
import time
# 通过 pip install 'volcengine-python-sdk[ark]' 安装方舟SDK
from volcenginesdkarkruntime import Ark

apikey="299aefde-c9ea-4f3c-8e23-c3aba030d327"

modelIdList = ['doubao-seedance-1-0-lite-t2v-250428']
# 请确保您已将 API Key 存储在环境变量 ARK_API_KEY 中
# 初始化Ark客户端，从环境变量中读取您的API Key
client = Ark(
    # 此为默认路径，您可根据业务所在地域进行配置
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    # 从环境变量中获取您的 API Key。此为默认方式，您可根据需要进行修改
    api_key=apikey,
)

if __name__ == "__main__":
    print("----- create request -----")
    create_result = client.content_generation.tasks.create(
        # 替换 <Model> 为模型的Model ID
        model="doubao-seedance-1-0-lite-t2v-250428",
        content=[
            {
                # 文本提示词
                "type": "text",
                "text": "A close-up selfie of two people.  The man, positioned on the left side of the image, is of Asian descent and appears to be in his 20s or 30s. He has short dark gray hair and is wearing rectangular glasses. His facial expression is slightly exaggerated, with a puzzled or funny look. He is wearing a casual, light-colored collared shirt. The woman, on the right, is also of Asian descent and appears to be of similar age. She has long, dark hair and is wearing large, round sunglasses with gold frames. Her lips are puckered in a playful manner. Both subjects are looking directly at the camera. The background is slightly out-of-focus, showing a partly visible outdoor scene with a pale blue sky. The lighting is bright, natural daylight, and the colors are muted, with a focus on the subjects. The image is a candid, informal selfie, highlighting the subjects' facial expressions and a casual interaction. The perspective is from a slightly low angle, looking up at the individuals. The composition emphasizes the subjects and their close proximity. The atmosphere is playful and friendly."
            },
        ]
    )
    print(create_result)

    # 轮询查询部分
    print("----- pooling task status -----")
    task_id = create_result.id
    while True:
        get_result = client.content_generation.tasks.get(task_id=task_id)
        status = get_result.status
        if status == "succeeded":
            print("----- task succeeded -----")
            print(get_result)
            break
        elif status == "failed":
            print("----- task failed -----")
            print(f"Error: {get_result.error}")
            break
        else:
            print(f"Current status: {status}, Retrying after 10 seconds...")
            time.sleep(10)