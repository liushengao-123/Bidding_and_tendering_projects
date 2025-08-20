# -*- coding: utf-8 -*-

import yaml
from pathlib import Path
import os
from flask import Flask, request, jsonify, Response
import logging
from datetime import datetime
import json
from werkzeug.utils import secure_filename
import tempfile
from typing import Generator, Dict

from openai import OpenAI

# 假设 read_ppt.py 文件存在且函数正确
# 如果这些文件不存在，请先注释掉这两行，并用虚拟数据代替
from ppt2context_total import extract_structured_text_from_pptx2
from read_ppt import extract_structured_text_from_pptx

# --- 全局配置 ---
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B-Instruct-2507")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-rirdsmibduedlwfaipdntinnqafzygcjgugtujnieuixggeq")
MODEL_API_URL = os.getenv("MODEL_API_URL", "https://api.siliconflow.cn/v1/").strip()
TIMEOUT = int(os.getenv("TIMEOUT", "300"))
OUTPUT_DIR = Path("./output_debug")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# --- 初始化 ---
app = Flask(__name__)
# 设置日志级别为DEBUG，以捕获所有信息
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
app.logger.setLevel(logging.DEBUG)


class PromptManager:
    def __init__(self, file_path="prompts.yaml"):
        try:
            self.prompts = yaml.safe_load(Path(file_path).read_text(encoding='utf-8'))
        except FileNotFoundError:
            app.logger.error(f"错误：找不到prompts文件 '{file_path}'。")
            self.prompts = {}
    def get_prompt(self, key):
        return self.prompts.get(key, f"PROMPT_KEY_{key}_NOT_FOUND")

pm = PromptManager("prompts.yaml")
EXAMPLE = pm.get_prompt("EXAMPLE")
LLM_KEY_VALUE_MAP = pm.get_prompt("LLM_KEY-VALUE_MAP")
SHANGHAI_SYSTEM_PROMPT = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT_STREAM")
LLM_KEY_VALUE_MAP_2 = pm.get_prompt("LLM_KEY-VALUE_MAP_2")
SHANGHAI_SYSTEM_PROMPT_2 = pm.get_prompt("SHANGHAI_SYSTEM_PROMPT_2")

# ==============================================================================
# vvvvvvvvvvvv   【调试专用】极简流式函数   vvvvvvvvvvvvvv
# ==============================================================================
def simple_model_streamer(payload: Dict) -> Generator[str, None, None]:
    """
    一个极度简化的调试函数，只负责调用API并打印原始返回，不做任何解析。
    """
    client = OpenAI(api_key=ACCESS_TOKEN, base_url=MODEL_API_URL, timeout=TIMEOUT)
    request_id = payload.get("request_id", "Unknown Task") # 使用 .get() 更安全
    
    app.logger.debug(f"[{request_id}] Preparing to send payload to API.")
    # 打印完整的 payload (除了 messages 的 content，因为它可能太长)
    loggable_payload = {k: v for k, v in payload.items() if k != 'messages'}
    loggable_payload['messages_count'] = len(payload.get('messages', []))
    app.logger.debug(f"[{request_id}] Payload (sanitized): {json.dumps(loggable_payload, ensure_ascii=False)}")
    
    try:
        app.logger.info(f"[{request_id}] ==> Sending request NOW...")
        response_stream = client.chat.completions.create(
            model=payload['model'],
            messages=payload['messages'],
            stream=True
        )
        
        chunk_received = False
        for chunk in response_stream:
            chunk_received = True
            content = chunk.choices[0].delta.content if chunk.choices and chunk.choices[0].delta.content else ""
            
            # 【核心调试点】直接打印和 yield 原始 content
            if content:
                app.logger.info(f"[{request_id}] << RAW CONTENT CHUNK RECEIVED: '{repr(content)}'")
                yield content
        
        if not chunk_received:
            app.logger.warning(f"[{request_id}] !!! Stream completed BUT NO CHUNKS were received from the API.")
        
        app.logger.info(f"[{request_id}] <== Stream finished.")

    except Exception as e:
        app.logger.error(f"!!! [{request_id}] FATAL ERROR during API call: {e}", exc_info=True)
        yield f'{{"error": "API call failed for {request_id}", "details": "{str(e)}"}}'

# ==============================================================================
# ^^^^^^^^^^^^^^^^   以上是调试专用函数   ^^^^^^^^^^^^^^^^
# ==============================================================================

@app.route('/process', methods=['POST'])
def process_documents():
    app.logger.info("-" * 50)
    app.logger.info(f"Received new stream request on /process from {request.remote_addr}")
    
    if 'pptx_file' not in request.files:
        return jsonify({"error": "Missing 'pptx_file'"}), 400

    pptx_file = request.files['pptx_file']
    if pptx_file.filename == '':
        return jsonify({"error": "Empty filename"}), 400

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_file:
        pptx_file.save(temp_file.name)
        temp_pptx_path = temp_file.name
    
    app.logger.info(f"PPTX saved to: {temp_pptx_path}")

    try:
        app.logger.info("Extracting text from PPTX...")
        extracted_text = extract_structured_text_from_pptx(temp_pptx_path)
        # extracted_text2 = extract_structured_text_from_pptx2(temp_pptx_path) # 暂时注释掉第二个，简化调试
        app.logger.info("Text extraction complete.")

        key_value_str = json.dumps(extracted_text, indent=2, ensure_ascii=False)
        
        # 完整的 User Prompt 构造
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
        # 【调试重点】暂时只使用一个 payload
        payloads = [
            {
                "model": MODEL_NAME,
                "messages": [{"role": "system", "content": SHANGHAI_SYSTEM_PROMPT}, {"role": "user", "content": shanghai_user_prompt}],
                "request_id": "Task 1 (DEBUG)"
            }
        ]
        
        # 打印完整的System Prompt，确保它被正确加载
        app.logger.debug("="*20 + " SYSTEM PROMPT FOR TASK 1 " + "="*20)
        app.logger.debug(SHANGHAI_SYSTEM_PROMPT)
        app.logger.debug("="*20 + " END OF SYSTEM PROMPT " + "="*20)

        def stream_response_generator():
            try:
                # 只处理第一个 payload
                payload_to_send = payloads[0]
                # 直接调用极简的调试函数
                yield from simple_model_streamer(payload_to_send)
            finally:
                app.logger.info("Stream generator finished. Cleaning up temp file.")
                os.unlink(temp_pptx_path)

        # 返回纯文本流，方便 curl 查看原始输出
        return Response(stream_response_generator(), mimetype='text/plain; charset=utf-8')

    except Exception as e:
        app.logger.error(f"!!! An unexpected error in /process endpoint: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8899, debug=True)