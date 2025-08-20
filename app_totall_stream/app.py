# -*- coding: utf-8 -*-

import yaml
from pathlib import Path
import os
import requests
from flask import Flask, request, jsonify, Response
import logging
from datetime import datetime
import json
from werkzeug.utils import secure_filename
import tempfile
from typing import Generator, List, Dict

# 导入 openai 库用于流式调用
from openai import OpenAI

# 假设 read_ppt.py 与 app.py 在同一目录下
from ppt2context_total import extract_structured_text_from_pptx2
from read_ppt import extract_structured_text_from_pptx

# --- 全局配置 ---
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B-Instruct-2507")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-rirdsmibduedlwfaipdntinnqafzygcjgugtujnieuixggeq")
# !! 注意: 流式API的URL可能与非流式不同，请确保这是支持stream的endpoint
# SiliconFlow 的 URL 通常是 /v1/，而不是 /v1/chat/completions
MODEL_API_URL = os.getenv("MODEL_API_URL", "https://api.siliconflow.cn/v1/").strip()
TIMEOUT = int(os.getenv("TIMEOUT", "300"))
OUTPUT_DIR = Path("./output") # 建议使用相对路径或可配置路径
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- 初始化 ---
app = Flask(__name__)
if __name__ != '__main__':
    gunicorn_logger = logging.getLogger('gunicorn.error')
    app.logger.handlers = gunicorn_logger.handlers
    app.logger.setLevel(gunicorn_logger.level)
else:
    # 为本地运行设置一个基础的logger
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')


class PromptManager:
    def __init__(self, file_path="prompts.yaml"):
        try:
            self.prompts = yaml.safe_load(Path(file_path).read_text(encoding='utf-8'))
        except FileNotFoundError:
            app.logger.error(f"错误：找不到prompts文件 '{file_path}'。")
            self.prompts = {}
    
    def get_prompt(self, key):
        return self.prompts.get(key, f"PROMPT_KEY_{key}_NOT_FOUND")

# 确保prompts.yaml路径正确
# pm = PromptManager("/data/shanxi_lsa_conda/app_totall/prompts.yaml")
pm = PromptManager("prompts.yaml") # 假设在同级目录
EXAMPLE = pm.get_prompt("EXAMPLE")
LLM_KEY_VALUE_MAP = pm.get_prompt("LLM_KEY-VALUE_MAP")
SHANGHAI_SYSTEM_PROMPT = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT_STREAM")
LLM_KEY_VALUE_MAP_2 = pm.get_prompt("LLM_KEY-VALUE_MAP_2")
SHANGHAI_SYSTEM_PROMPT_2 = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT_2")

# ==============================================================================
# vvvvvvvvvvvv   全新的流式处理核心函数   vvvvvvvvvvvvvv
# ==============================================================================

def stream_and_parse_model_api(payload: Dict) -> Generator[Dict, None, None]:
    """
    向模型API发送单个流式请求，并实时解析返回的JSON对象。

    :param payload: 发送给模型API的请求体，包含model, messages等。
    :return: 一个生成器，逐个产出解析出的JSON对象(字典)。
    """
    client = OpenAI(api_key=ACCESS_TOKEN, base_url=MODEL_API_URL, timeout=TIMEOUT)
    request_id = payload.pop("request_id", "Unknown Task")

    app.logger.info(f"[{request_id}] Starting stream request to model API...")
    
    try:
        response_stream = client.chat.completions.create(
            model=payload['model'],
            messages=payload['messages'],
            stream=True
        )

        buffer = ""
        brace_level = 0
        in_string = False

        for chunk in response_stream:
            content = chunk.choices[0].delta.content if chunk.choices and chunk.choices[0].delta else ""
            if not content:
                continue

            for char in content:
                buffer += char
                if char == '"':
                    # 简单处理字符串内的引号，避免错误分割
                    # 注意：这无法处理转义的引号 '\"'，但对于大多数LLM输出足够了
                    in_string = not in_string
                
                if not in_string:
                    if char == '{':
                        brace_level += 1
                    elif char == '}':
                        brace_level -= 1
                
                # 当找到一个完整的顶层JSON对象时 (brace_level回到0)
                if brace_level == 0 and buffer.strip().startswith('{') and buffer.strip().endswith('}'):
                    try:
                        # 尝试解析这个对象
                        parsed_obj = json.loads(buffer)
                        yield parsed_obj  # 成功解析，产出这个对象
                        buffer = ""  # 清空缓冲区，准备接收下一个对象
                    except json.JSONDecodeError:
                        # 如果解析失败，说明对象还不完整，继续在缓冲区中累积
                        pass
        
        app.logger.info(f"[{request_id}] Stream finished.")

    except Exception as e:
        app.logger.error(f"!!! [{request_id}] An error occurred during streaming: {e}", exc_info=True)
        # 在流中产生一个错误对象，让客户端知道出错了
        yield {"error": f"Error in {request_id}", "details": str(e)}

# ==============================================================================
# ^^^^^^^^^^^^^^^^   以上是全新的流式处理核心函数   ^^^^^^^^^^^^^^^^
# ==============================================================================


@app.route('/process', methods=['POST'])
def process_documents():
    app.logger.info("-" * 50)
    app.logger.info(f"Received new stream request on /process from {request.remote_addr}")
    
    if 'pptx_file' not in request.files:
        return jsonify({"error": "请求中必须包含名为 'pptx_file' 的文件"}), 400

    pptx_file = request.files['pptx_file']
    if pptx_file.filename == '':
        return jsonify({"error": "上传的文件名不能为空"}), 400

    # 使用临时文件处理上传，更安全
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
        pptx_file.save(temp_file.name)
        temp_pptx_path = temp_file.name

    app.logger.info(f"Temporarily saved PPTX to: {temp_pptx_path}")

    try:
        # 1. 从PPTX提取文本
        app.logger.info("Extracting text from PPTX...")
        extracted_text = extract_structured_text_from_pptx(temp_pptx_path)
        extracted_text2 = extract_structured_text_from_pptx2(temp_pptx_path)
        app.logger.info("Text extraction complete.")

        # 2. 准备两个任务的Payload
        key_value_str = json.dumps(extracted_text, indent=2, ensure_ascii=False)
        key_value_str2 = json.dumps(extracted_text2, indent=2, ensure_ascii=False)
        
        shanghai_user_prompt = f"""
请根据你在系统指令中被设定的角色和规则，处理以下信息并返回JSON结果。

### 1. Key对应关系表 (English Key -> Chinese Key)
---
{LLM_KEY_VALUE_MAP}

### 2. 源数据文档 (Source Document to search in)
---
{key_value_str}

### 3. 输出样例 (Output Example)
---
{EXAMPLE}

请开始抽取。
"""
        shanghai_user_prompt_2 = f"""
请根据你在系统指令中被设定的角色和规则，处理以下信息并返回JSON结果。

### 1. Key对应关系表 (English Key -> Chinese Key)
---
{LLM_KEY_VALUE_MAP_2}

### 2. 源数据文档 (Source Document to search in)
---
{key_value_str2}

### 3. 输出样例 (Output Example)
---
{EXAMPLE}

请开始抽取。
"""
        payloads = [
            {
                "model": MODEL_NAME,
                "messages": [{"role": "system", "content": SHANGHAI_SYSTEM_PROMPT}, {"role": "user", "content": shanghai_user_prompt}],
                "request_id": "Task 1"
            },
            {
                "model": MODEL_NAME,
                "messages": [{"role": "system", "content": SHANGHAI_SYSTEM_PROMPT_2}, {"role": "user", "content": shanghai_user_prompt_2}],
                "request_id": "Task 2"
            }
        ]

        # 3. 创建一个生成器函数来顺序处理流
        def stream_response_generator():
            full_response_for_saving = []

            def _process_and_format_stream(stream_generator):
                """
                这个内部生成器处理所有事情：打印、收集、格式化和 yield。
                """
                for item in stream_generator:
                    print("Received item:", item) # 实时打印到服务器控制台
                    full_response_for_saving.append(item) # 收集以备最终保存
                    yield json.dumps(item, ensure_ascii=False) + '\n' # 格式化并产出

            try:
                # 现在主逻辑变得异常简单和清晰！
                for payload in payloads:
                    yield from _process_and_format_stream(stream_and_parse_model_api(payload))

            finally:
                # ... finally 块中的逻辑保持不变 ...
                app.logger.info("Stream ended. Saving full result to file.")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"result_{timestamp}.json"
                output_filepath = OUTPUT_DIR / output_filename
                try:
                    with open(output_filepath, 'w', encoding='utf-8') as f:
                        json.dump(full_response_for_saving, f, ensure_ascii=False, indent=4)
                    app.logger.info(f"Full result successfully saved to {output_filepath}")
                except Exception as save_e:
                    app.logger.error(f"!!! Failed to save full result to file: {save_e}", exc_info=True)
                
                # 清理临时文件
                os.unlink(temp_pptx_path)
                app.logger.info(f"Cleaned up temporary file: {temp_pptx_path}")

        # 4. 返回流式响应
        return Response(stream_response_generator(), mimetype='application/json-lines')

    except Exception as e:
        app.logger.error(f"!!! An unexpected error occurred in /process endpoint: {e}", exc_info=True)
        # 如果在流开始前就出错，则返回一个标准的JSON错误
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8899, debug=True)