# -*- coding: utf-8 -*-
"""
@Time: 2024/3/8 14:10
@Auth: MCZ
@File: concurrent_test.py
@IDE: PyCharm
@Motto: ABC(Always Be Coding)
"""
import time
import concurrent.futures
from openai import OpenAI

openai_api_key = "EMPTY"
openai_api_base = "xxxx"

client = OpenAI(
    api_key=openai_api_key,
    base_url=openai_api_base,
)

# 定义要发送到 API 的提示或文本
prompt = "从前有一天，"


# 定义发送请求到 OpenAI API 的函数
def send_request(prompt):
    chat_completion = client.chat.completions.create(
        messages=[{
            "role": "system",
            "content": ""
        }, {
            "role": "user",
            "content": prompt
        }],
        model="qwen-14b",
        stream=False,
    )
    return chat_completion.choices[0].message.content


# 定义并发请求数量
num_requests = 100

# 测量发送请求的并发时间
start_time = time.time()

with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = [executor.submit(send_request, prompt) for _ in range(num_requests)]
    results = [future.result() for future in futures]

end_time = time.time()

# 计算总耗时
total_time = end_time - start_time

# 打印结果和并发信息
print(f"总共用时 {num_requests} 个请求：{total_time} 秒")
print("结果:")
for i, result in enumerate(results):
    print(f"请求 {i + 1}: {result}")
