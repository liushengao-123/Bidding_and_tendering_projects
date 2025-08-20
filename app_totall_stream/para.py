# -*- coding: utf-8 -*-

import yaml
from pathlib import Path
import os
import requests
from flask import Flask, request, jsonify
import logging
from datetime import datetime
import json
from werkzeug.utils import secure_filename
import tempfile
# 导入并发处理库
import concurrent.futures

# 假设 read_ppt.py 和 prompts.yaml 与 app.py 在同一目录下
from ppt2context_total import extract_structured_text_from_pptx2
from read_ppt import extract_structured_text_from_pptx
# --- 全局配置 ---
#ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-mjgunayzdoeaycpmhedpkvppmvhgspatesqafaiuelsomwkr")
#MODEL_API_URL = os.getenv("MODEL_API_URL", "http://140.210.92.250:25081/v1/chat/completions")
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B-Instruct-2507")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-rirdsmibduedlwfaipdntinnqafzygcjgugtujnieuixggeq")
MODEL_API_URL = os.getenv("MODEL_API_URL", "https://api.siliconflow.cn/v1/chat/completions")
#MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B-Thinking-2507")

TIMEOUT = int(os.getenv("TIMEOUT", "300000"))
DEFAULT_RESPONSE = "对不起，服务暂时不可用。请稍后再试。"
OUTPUT_DIR = Path("/app/output")

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

class PromptManager:
    # 请确保这里的路径是正确的
    def __init__(self, file_path="prompts.yaml"): # 假设prompts.yaml在同级目录
        try:
            self.prompts = yaml.safe_load(Path(file_path).read_text(encoding='utf-8'))
        except FileNotFoundError:
            print(f"错误：找不到prompts文件 '{file_path}'。请确保文件路径正确。")
            self.prompts = {}
    
    def get_prompt(self, key):
        return self.prompts.get(key, f"PROMPT_KEY_{key}_NOT_FOUND")

# --- 初始化 ---
app = Flask(__name__)
if __name__ != '__main__':
    gunicorn_logger = logging.getLogger('gunicorn.error')
    app.logger.handlers = gunicorn_logger.handlers
    app.logger.setLevel(gunicorn_logger.level)

pm = PromptManager("E:\project\优化版本1\\app_totall\prompts.yaml")
EXAMPLE = pm.get_prompt("EXAMPLE")
NOTE = pm.get_prompt("NOTE")
LLM_KEY_VALUE_MAP = pm.get_prompt("LLM_KEY-VALUE_MAP")
SHANGHAI_SYSTEM_PROMPT = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT")

LLM_KEY_VALUE_MAP_2 = pm.get_prompt("LLM_KEY-VALUE_MAP_2")
SHANGHAI_SYSTEM_PROMPT_2 = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT_2")


# ==============================================================================
# vvvvvvvvvvvvvvvv   这里是主要修改的函数   vvvvvvvvvvvvvvvvvv
# ==============================================================================

def call_model_api(original_question,original_question2 ,doc_context):
    """
    并发调用大语言模型API的核心函数，合并结果后返回。
    """
    # 将原始数据转换为格式化的JSON字符串，用于prompt
    key_value_str = json.dumps(original_question, indent=2, ensure_ascii=False)
    key_value_str2 = json.dumps(original_question2, indent=2, ensure_ascii=False)
    # 准备通用的User Prompt
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

    # 准备两个不同的请求Payload
    payloads = [
        {
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": f"{SHANGHAI_SYSTEM_PROMPT}"},
                {"role": "user", "content": f'{shanghai_user_prompt}'}
            ],
            "request_id": "Task 1" # 添加一个ID用于日志跟踪
        },
        {
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": f"{SHANGHAI_SYSTEM_PROMPT_2}"},
                # 注意：这里我们复用相同的user prompt，因为源数据是一样的
                {"role": "user", "content": f'{shanghai_user_prompt_2}'} 
            ],
            "request_id": "Task 2" # 添加一个ID用于日志跟踪
        }
    ]

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }

    # 定义一个发送单个请求的函数
    def send_request(payload):
        request_id = payload.pop("request_id", "Unknown Task") # 弹出自定义ID，不发送给API
        app.logger.info(f"[{request_id}] Sending POST request to model API...")
        
        try:
            response = requests.post(MODEL_API_URL, json=payload, headers=headers, timeout=TIMEOUT)
            app.logger.info(f"[{request_id}] Received response. Status Code: {response.status_code}")
            response.raise_for_status()
            
            response_data = response.json()
            content = response_data.get('choices', [{}])[0].get('message', {}).get('content', DEFAULT_RESPONSE)
            app.logger.info(f"[{request_id}] Successfully parsed model response.")
            return content
        except requests.exceptions.Timeout:
            app.logger.error(f"!!! [{request_id}] Request timed out.", exc_info=True)
            return json.dumps({"error": f"{request_id} timed out"})
        except requests.exceptions.RequestException as e:
            app.logger.error(f"!!! [{request_id}] A network error occurred: {e}", exc_info=True)
            return json.dumps({"error": f"Network error in {request_id}: {e}"})
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            app.logger.error(f"!!! [{request_id}] Error processing response: {e}", exc_info=True)
            return json.dumps({"error": f"Error parsing response from {request_id}: {e}"})

    app.logger.info("="*20 + " Starting concurrent API calls " + "="*20)
    
    # 使用线程池并发执行两个请求
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        # executor.map会按顺序返回结果
        results = list(executor.map(send_request, payloads))

    app.logger.info("="*20 + " All concurrent API calls finished " + "="*20)

    # # 合并结果
    # merged_data = {}
    # for i, result_str in enumerate(results):
    #     app.logger.info(f"Result from Task {i+1}:\n{result_str[:500]}...") # 打印每个任务返回结果的前500字符
    #     try:
    #         # 尝试将结果字符串解析为JSON对象（字典）
    #         data = json.loads(result_str)
    #         if isinstance(data, dict):
    #             # 如果是字典，则更新到合并字典中
    #             merged_data.update(data)
    #         else:
    #             # 如果不是字典（例如是列表或字符串），则存入一个特定的键
    #             merged_data[f"task_{i+1}_result"] = data
    #     except json.JSONDecodeError:
    #         # 如果解析失败，说明返回的不是有效的JSON，作为原始字符串存入
    #         app.logger.warning(f"Task {i+1} result is not valid JSON. Storing as raw string.")
    #         merged_data[f"task_{i+1}_raw_result"] = result_str

    #final_json_output = json.dumps(merged_data, ensure_ascii=False, indent=4)

    # =======================================================
    # vvvvvvv  新的、用于合并列表的逻辑  vvvvvvv
    # =======================================================
    merged_list = [] # 初始化一个空列表，而不是字典
    for i, result_str in enumerate(results):
        app.logger.info(f"Result from Task {i+1}:\n{result_str[:500]}...") # 打印每个任务返回结果的前500字符
        try:
            # 尝试将结果字符串解析为JSON对象
            data = json.loads(result_str)
            
            # 检查解析出的数据是否是列表类型
            if isinstance(data, list):
                # 如果是列表，则将其所有元素添加到我们的主列表中
                merged_list.extend(data)
            else:
                # 如果返回的不是一个列表（例如，可能是一个包含错误的字典），
                # 记录一个警告，这样它就不会破坏最终的输出结构。
                app.logger.warning(
                    f"Task {i+1} result was a valid JSON but not a list, so it will be skipped in the final merge. "
                    f"Result type: {type(data)}, Content (truncated): {str(data)[:200]}"
                )
        except json.JSONDecodeError:
            # 如果解析失败，说明返回的不是有效的JSON，记录警告并跳过。
            app.logger.warning(
                f"Task {i+1} result is not valid JSON and will be skipped. "
                f"Raw content (truncated): {result_str[:200]}"
            )
            
    # 将合并后的Python列表转换回格式化的JSON字符串
    # 最终的输出将是一个JSON数组，就像您的示例一样
    final_json_output = json.dumps(merged_list, ensure_ascii=False, indent=4)
    # =======================================================
    # ^^^^^^^  新的、用于合并列表的逻辑  ^^^^^^^
    # =======================================================

    # 将合并后的Python字典转换回格式化的JSON字符串
   
    
    app.logger.info("Successfully merged results.")
    return final_json_output

# ==============================================================================
# ^^^^^^^^^^^^^^^^   以上是主要修改的函数   ^^^^^^^^^^^^^^^^^^
# ==============================================================================


@app.route('/process', methods=['POST'])
def process_documents():
    app.logger.info("-" * 50)
    app.logger.info(f"Received new request on /process from {request.remote_addr}")
    
    if 'pptx_file' not in request.files: # 您原代码检查了pptx和docx，这里根据您的函数调用只保留pptx
        app.logger.warning("Request is missing 'pptx_file' in form-data.")
        return jsonify({"error": "请求中必须包含名为 'pptx_file' 的文件"}), 400

    pptx_file = request.files['pptx_file']

    if pptx_file.filename == '':
        app.logger.warning("The uploaded file has no name.")
        return jsonify({"error": "上传的文件名不能为空"}), 400

    pptx_filename = secure_filename(pptx_file.filename)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pptx_path = Path(temp_dir) / pptx_filename
        pptx_file.save(temp_pptx_path)
        
        app.logger.info(f"Temporarily saved PPTX to: {temp_pptx_path}")

        try:
            app.logger.info(f"Extracting text from PPTX: {temp_pptx_path}")
            extracted_text = extract_structured_text_from_pptx(temp_pptx_path)
            extracted_text2= extract_structured_text_from_pptx2(temp_pptx_path)
            formatted_json = json.dumps(extracted_text, ensure_ascii=False, indent=4)
            app.logger.info("Extracted text from PPTX successfully.")
            # print(formatted_json) # 在生产环境中可以注释掉这个print

            # 调用经过改造的、支持并发的API函数
            # doc_context 变量在您的代码中未被使用，因此传入空字符串
            result_content = call_model_api(extracted_text,extracted_text2, "")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"result_{timestamp}.json" # 建议保存为.json
            output_filepath = OUTPUT_DIR / output_filename
            
            with open(output_filepath, 'w', encoding='utf-8') as f:
                f.write(result_content)
            
            app.logger.info(f"Result saved to {output_filepath}")

            # 返回时，为了方便前端处理，可以直接返回解析后的JSON对象
            return jsonify({
                "code": 200,
                "message": "Processing successful",
                "output_file_container_path": str(output_filepath),
                "data": json.loads(result_content) # 返回JSON对象而不是字符串
            }), 200

        except Exception as e:
            app.logger.error(f"!!! An unexpected error occurred in /process endpoint: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8894, debug=True)