"""
工具函数模块

这个模块提供通用的工具函数，包括：
- 日志记录
- 目录创建
- 时间戳处理

这些函数被项目中的其他模块使用。
"""

import os
import sys
from datetime import datetime

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from UI_checktest.constants import BASE_RESULT_DIR

# 创建结果目录（如果不存在）
os.makedirs(BASE_RESULT_DIR, exist_ok=True)

# 日志文件路径，包含时间戳
LOG_FILE = os.path.join(BASE_RESULT_DIR, f"test_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


def ensure_dir_exists(path):
    """
    确保目录存在

    如果目录不存在，创建它（包括父目录）。

    Args:
        path (str): 目录路径
    """
    if path:
        os.makedirs(path, exist_ok=True)


def logger(message):
    """
    记录日志信息

    将消息同时输出到控制台和日志文件。
    日志格式： [时间戳] 消息

    Args:
        message (str): 要记录的消息
    """
    # 生成带毫秒的时间戳
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    log_entry = f"[{timestamp}] {message}"

    # 输出到控制台
    print(log_entry)

    # 写入日志文件
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_entry + '\n')
