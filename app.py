# -*- coding: utf-8 -*-

import yaml
from pathlib import Path
import os
import requests
from flask import Flask, request, jsonify # <--- 【修正1】从这里移除了 logging
import logging # <--- 【修正2】单独、正确地导入标准库 logging
from datetime import datetime
import json
from werkzeug.utils import secure_filename # 导入 secure_filename
import tempfile # 用于创建临时文件
# 假设 read_ppt.py 和 prompts.yaml 与 app.py 在同一目录下
from read_ppt import extract_structured_text_from_pptx

# --- 全局配置 ---
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-uUTtsplQO7yzLVQH40682353C3B44a9bB417045f9321B563")
MODEL_API_URL = os.getenv("MODEL_API_URL", "http://140.210.92.250:25081/v1/chat/completions")
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen3-30B-A3B-Instruct-2507")
# MODEL_API_URL = os.getenv("MODEL_API_URL", "http://172.16.81.180:3000/v1/chat/completions")
# MODEL_NAME = os.getenv("MODEL_NAME", "DeepSeek")
TIMEOUT = int(os.getenv("TIMEOUT", "300000"))
DEFAULT_RESPONSE = "对不起，服务暂时不可用。请稍后再试。"
OUTPUT_DIR = Path("/app/output")

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

class PromptManager:
    def __init__(self, file_path="E:\project\优化版本1\prompts.yaml"):
        self.prompts = yaml.safe_load(Path(file_path).read_text(encoding='utf-8'))
    
    def get_prompt(self, key):
        return self.prompts.get(key, f"PROMPT_KEY_{key}_NOT_FOUND")

# --- 初始化 ---
app = Flask(__name__)
# 配置Gunicorn的日志，使其能正常输出
if __name__ != '__main__':
    gunicorn_logger = logging.getLogger('gunicorn.error')
    app.logger.handlers = gunicorn_logger.handlers
    app.logger.setLevel(gunicorn_logger.level)

pm = PromptManager()
EXAMPLE = pm.get_prompt("EXAMPLE")
NOTE = pm.get_prompt("NOTE")
LLM_KEY_VALUE_MAP = pm.get_prompt("LLM_KEY-VALUE_MAP")
SHANGHAI_SYSTEM_PROMPT = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT")



# call_model_api 和 process_documents 函数保持不变...
# (这里省略了这两个函数的代码，因为它们没有变化，请保留您文件中的这部分)
def call_model_api(original_question, doc_context):
    original_question = json.dumps(original_question, indent=2, ensure_ascii=False)
    KEY_VALUE=original_question#中文的key和中文的value集合,ppt中抽取
    """
    调用大语言模型API的核心函数
    """
    SHANGHAI_USER_PROMPT = f"""
请根据你在系统指令中被设定的角色和规则，处理以下信息并返回JSON结果。

### 1. Key对应关系表 (English Key -> Chinese Key)
---
{LLM_KEY_VALUE_MAP}

### 2. 源数据文档 (Source Document to search in)
---
{KEY_VALUE}

### 3. 输出样例 (Output Example)
---
{EXAMPLE}

请开始抽取。
"""

    print(KEY_VALUE)




    # --- 【调试日志1】 打印将要发送的请求信息 ---
    app.logger.info("="*20 + " Preparing to call Model API " + "="*20)
    #app.logger.info(f"Target URL: {MODEL_API_URL}")
    #app.logger.info(f"Model Name: {MODEL_NAME}")
    app.logger.info(f"Timeout set to: {TIMEOUT} seconds")
    app.logger.info(f" {original_question} ")
    app.logger.info(f" ############################################### ")
    app.logger.info(f" {KEY_VALUE} ")
    app.logger.info(f" ############################################### ")
    request_payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system",
             "content": f"{SHANGHAI_SYSTEM_PROMPT}"},
            {"role": "user",
             "content": f'{SHANGHAI_USER_PROMPT}'}
        ],
        
        
    }
    
    # --- 【调试日志2】 打印完整的请求体(Payload)，这是排查问题的关键 ---
    # 使用json.dumps美化输出，方便查看
    # try:
    #     # 因为原始文本可能非常长，截断一下以防日志爆炸
    #     log_payload = request_payload.copy()
    #     log_payload['messages'][1]['content'] = log_payload['messages'][1]['content'][:500] + '...' # 截断user content
    #     payload_str = json.dumps(log_payload, indent=4, ensure_ascii=False)
    #     app.logger.info(f"Request Payload (truncated for log):\n{payload_str}")
    # except Exception as e:
    #     app.logger.error(f"Failed to serialize payload for logging: {e}")

    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    app.logger.info(f"Request Headers: {headers}")
    
    try:
        # --- 【调试日志3】 标记请求即将发出 ---
        app.logger.info("Sending POST request to model API...")
        print(request_payload)
        response = requests.post(MODEL_API_URL, json=request_payload, headers=headers, timeout=TIMEOUT)
        
        # --- 【调试日志4】 标记收到响应，并打印状态码和部分内容 ---
        app.logger.info(f"Received response. Status Code: {response.status_code}")
        # 打印响应内容的前500个字符，防止日志过长
        app.logger.info(f"Response Body (first 500 chars): {response.text[:500]}")
        
        response.raise_for_status()
        
        response_data = response.json()
        best_guess = response_data.get('choices', [{}])[0].get('message', {}).get('content', DEFAULT_RESPONSE)
        app.logger.info("Successfully parsed model response.")
        return best_guess

    except requests.exceptions.Timeout:
        # --- 【调试日志5】 专门捕获并记录超时错误 ---
        app.logger.error("!!! Request timed out. The model API server did not respond in time.", exc_info=True)
        raise ConnectionError(f"无法连接到模型API: 请求在{TIMEOUT}秒后超时。")
    except requests.exceptions.RequestException as e:
        # --- 【调试日志6】 捕获其他所有网络相关的错误 ---
        app.logger.error(f"!!! A network error occurred while calling model API: {e}", exc_info=True)
        raise ConnectionError(f"无法连接到模型API: {e}")
    except (KeyError, IndexError, json.JSONDecodeError) as e:
        app.logger.error(f"!!! Error processing or parsing the model's response: {e}", exc_info=True)
        raise ValueError(f"解析模型响应时出错: {e}")


@app.route('/process', methods=['POST'])
def process_documents():
    app.logger.info("-" * 50)
    app.logger.info(f"Received new request on /process from {request.remote_addr}")
    
    # 1. 检查文件是否在请求中
    if 'pptx_file' not in request.files or 'docx_file' not in request.files:
        app.logger.warning("Request is missing 'pptx_file' or 'docx_file' in form-data.")
        return jsonify({"error": "请求中必须包含名为 'pptx_file' 和 'docx_file' 的文件"}), 400

    pptx_file = request.files['pptx_file']
    docx_file = request.files['docx_file']

    # 2. 检查文件名是否为空
    if pptx_file.filename == '' or docx_file.filename == '':
        app.logger.warning("One of the uploaded files has no name.")
        return jsonify({"error": "上传的文件名不能为空"}), 400

    # 为了安全，使用安全的文件名
    pptx_filename = secure_filename(pptx_file.filename)
    docx_filename = secure_filename(docx_file.filename)

    # 3. 【重要】因为你的处理函数 extract_structured_text_from_pptx 和 read_docx 接收的是文件路径
    # 而不是文件流，所以我们需要先将上传的文件保存到服务器的一个临时位置。
    # 创建一个临时目录来存放上传的文件
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_pptx_path = Path(temp_dir) / pptx_filename
        temp_docx_path = Path(temp_dir) / docx_filename

        pptx_file.save(temp_pptx_path)
        docx_file.save(temp_docx_path)
        
        app.logger.info(f"Temporarily saved PPTX to: {temp_pptx_path}")
        app.logger.info(f"Temporarily saved DOCX to: {temp_docx_path}")

        try:
            # 使用临时文件路径调用你的现有函数
            app.logger.info(f"Reading PPTX from: {temp_pptx_path}")
            extracted_text = extract_structured_text_from_pptx(temp_pptx_path)
            # new_data = [
            #             {"交货期": "从采购方案-合同主要条款中抽取"},
            #             {"交货地点": "从采购方案-合同主要条款中抽取"}
            #             ]
            # extracted_text["采购方案"].extend(new_data)

            

            formatted_json = json.dumps(extracted_text, ensure_ascii=False, indent=4)
            print(formatted_json)

            # 可选：保存到文件
            with open('project_details.json', 'w', encoding='utf-8') as f:
                json.dump(extracted_text, f, ensure_ascii=False, indent=4)

            app.logger.info(f"Reading DOCX from: {temp_docx_path}")
            #doc_context = read_docx(temp_docx_path)
            doc_context=""
            result_content = call_model_api(extracted_text, doc_context)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"result_{timestamp}.txt"
            output_filepath = OUTPUT_DIR / output_filename
            
            with open(output_filepath, 'w', encoding='utf-8') as f:
                f.write(result_content)
            
            app.logger.info(f"Result saved to {output_filepath}")

            return jsonify({
                "code": 200,
                "message": "Processing successful",
                "output_file_container_path": str(output_filepath),
                "data": result_content
            }), 200

        except Exception as e:
            app.logger.error(f"!!! An unexpected error occurred in /process endpoint: {e}", exc_info=True)
            return jsonify({"error": str(e)}), 500
        # 临时文件和目录在此 'with' 代码块结束时会自动被删除

if __name__ == '__main__':
    # 使用 gunicorn 启动时不会执行这里，这里仅用于本地调试
    app.run(host='0.0.0.0', port=8893, debug=True)

















