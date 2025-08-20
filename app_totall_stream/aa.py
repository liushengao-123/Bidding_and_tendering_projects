import logging
import os
import sys
import time
from logging.handlers import TimedRotatingFileHandler
from typing import Generator, List, Dict, Any

from flask import Flask, request, Response, jsonify
from openai import OpenAI, APIConnectionError, RateLimitError, APIError

# --- 1. Flask 应用和日志初始化 ---
app = Flask(__name__)

def setup_logger(name: str) -> logging.Logger:
    """配置一个同时输出到控制台和文件的日志记录器"""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger
    
    logger.setLevel(logging.INFO)
    file_handler = TimedRotatingFileHandler("stream_client.log", when="midnight", interval=1, backupCount=7, encoding='utf-8')
    stream_handler = logging.StreamHandler()
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    stream_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    return logger

logger = setup_logger("deepseek_flask_app")


# --- 2. 配置参数 ---
MODEL_NAME = os.getenv("MODEL_NAME", "Qwen/Qwen3-30B-A3B-Instruct-2507")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN", "sk-rirdsmibduedlwfaipdntinnqafzygcjgugtujnieuixggeq")
MODEL_API_URL = os.getenv("MODEL_API_URL", "https://api.siliconflow.cn/v1/").strip()
TIMEOUT = int(os.getenv("TIMEOUT", "300"))
DEFAULT_RESPONSE = "对不起，服务暂时不可用。请稍后再试。"


# --- 3. 核心流式查询函数 (保持不变) ---
def query_deepseek_stream(
    api_key: str,
    messages: List[Dict[str, str]],
    model: str,
    base_url: str,
    max_retries: int = 3,
    timeout: int = 300
) -> Generator[str, None, None]:
    """流式查询模型的生成器函数"""
    try:
        client = OpenAI(api_key=api_key, base_url=base_url, timeout=timeout)
    except Exception as e:
        logger.error(f"创建 OpenAI 客户端失败: {e}")
        yield DEFAULT_RESPONSE
        return

    for attempt in range(max_retries):
        try:
            logger.info(f"正在发送请求 (第 {attempt + 1}/{max_retries} 次尝试)...")
            response = client.chat.completions.create(
                model=model,
                messages=messages,
                stream=True
            )
            chunk_received = False
            for chunk in response:
                delta = chunk.choices[0].delta if chunk.choices else None
                if not delta:
                    continue
                content = getattr(delta, "reasoning_content", None) or getattr(delta, "content", None)
                if content:
                    chunk_received = True
                    yield content
            
            if chunk_received:
                logger.info("流式响应成功完成。")
                return
        except (APIConnectionError, RateLimitError, APIError) as e:
            logger.error(f"API 请求失败 (第 {attempt + 1} 次): {type(e).__name__} - {e}")
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                logger.info(f"{wait_time} 秒后重试...")
                time.sleep(wait_time)
            else:
                logger.error("已达到最大重试次数，请求最终失败。")
                yield DEFAULT_RESPONSE
        except Exception as e:
            logger.error(f"发生未知错误 (第 {attempt + 1} 次): {e}", exc_info=True)
            if attempt < max_retries - 1:
                 time.sleep(1)
            else:
                yield DEFAULT_RESPONSE


# --- 4. Flask API 端点 (关键修改部分) ---
# @app.route('/stream-chat', methods=['POST'])
# def stream_chat():
#     """
#     接收 POST 请求，处理并以流式响应返回，同时在日志中记录完整的响应内容。
#     """
#     try:
#         data = request.get_json()
#         if not data or 'messages' not in data:
#             return jsonify({"error": "请求体中缺少 'messages' 字段"}), 400

#         messages = data['messages']
#         model = data.get('model', MODEL_NAME) 

#         # 在日志中记录传入的请求，为了安全可以隐去部分内容
#         logger.info(f"收到来自 {request.remote_addr} 的请求, 模型: {model}, 消息数量: {len(messages)}")

#         def generate():
#             # 用于收集所有响应块的列表
#             response_parts = []
#             try:
#                 generator = query_deepseek_stream(
#                     api_key=ACCESS_TOKEN,
#                     messages=messages,
#                     model=model,
#                     base_url=MODEL_API_URL,
#                     timeout=TIMEOUT
#                 )
#                 for chunk in generator:
#                     response_parts.append(chunk) # 收集块
#                     yield chunk                  # 同时流式传输给客户端
#             except Exception as e:
#                 logger.error(f"在流式生成过程中发生错误: {e}", exc_info=True)
#                 yield DEFAULT_RESPONSE
#             finally:
#                 # --- 新增日志记录 ---
#                 # 在流结束后 (无论成功与否)，将收集到的内容拼接成完整响应并记录
#                 if response_parts:
#                     full_response = "".join(response_parts)
#                     # 为了防止日志文件过大，只记录响应的开头部分
#                     log_preview = (full_response[:500] + '...') if len(full_response) > 500 else full_response
#                     logger.info(f"向 {request.remote_addr} 的流式响应已结束。完整响应预览: {log_preview.strip()}")
#                 else:
#                     logger.warning(f"向 {request.remote_addr} 的流式响应未产生任何内容。")

#         return Response(generate(), mimetype='text/plain; charset=utf-8')

#     except Exception as e:
#         logger.error(f"处理 /stream-chat 请求时发生严重错误: {e}", exc_info=True)
#         return jsonify({"error": "服务器内部错误"}), 500
# ... (文件的其他部分，如 setup_logger, query_deepseek_stream 等保持不变) ...

# --- 4. Flask API 端点 (关键修改部分) ---
# --- 4. Flask API 端点 (已修正) ---
@app.route('/stream-chat', methods=['POST'])
def stream_chat():
    """
    接收 POST 请求，处理并以流式响应返回，同时在命令行中实时打印流式内容。
    """
    try:
        data = request.get_json()
        if not data or 'messages' not in data:
            return jsonify({"error": "请求体中缺少 'messages' 字段"}), 400

        messages = data['messages']
        model = data.get('model', MODEL_NAME) 

        # --- 关键修正 ---
        # 在请求上下文依然有效时，提前获取 IP 地址
        client_ip = request.remote_addr
        
        logger.info(f"收到来自 {client_ip} 的流式请求, 模型: {model}")

        def generate():
            response_parts = []
            try:
                generator = query_deepseek_stream(
                    api_key=ACCESS_TOKEN,
                    messages=messages,
                    model=model,
                    base_url=MODEL_API_URL,
                    timeout=TIMEOUT
                )
                for chunk in generator:
                    print(chunk, end="", flush=True) 
                    response_parts.append(chunk)
                    yield chunk
            
            except Exception as e:
                logger.error(f"在流式生成过程中发生错误: {e}", exc_info=True)
                yield DEFAULT_RESPONSE
            finally:
                print() 
                if response_parts:
                    full_response = "".join(response_parts)
                    log_preview = (full_response[:500] + '...') if len(full_response) > 500 else full_response
                    
                    # --- 关键修正 ---
                    # 使用之前保存的局部变量 client_ip，而不是直接访问 request 对象
                    logger.info(f"向 {client_ip} 的流式响应已结束。完整响应预览: {log_preview.strip()}")
                else:
                    logger.warning(f"向 {client_ip} 的流式响应未产生任何内容。")

        return Response(generate(), mimetype='text/plain; charset=utf-8')

    except Exception as e:
        logger.error(f"处理 /stream-chat 请求时发生严重错误: {e}", exc_info=True)
        return jsonify({"error": "服务器内部错误"}), 500


# ... (文件的其余部分，如 if __name__ == '__main__': ... 保持不变) ...

# --- 5. 启动应用 ---
if __name__ == '__main__':
    if not ACCESS_TOKEN or "sk-" not in ACCESS_TOKEN:
        logger.error("错误：环境变量 ACCESS_TOKEN 未设置或格式不正确。请设置您的 API 密钥。")
        sys.exit(1)
    
    app.run(host='0.0.0.0', port=5000, debug=True)