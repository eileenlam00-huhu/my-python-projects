import os
import time
import subprocess
import cv2
import numpy as np
from .constants import BASE_RESULT_DIR, REMOTE_DIR
from .click import adb_shell, safe_read_line
from .utils import logger, ensure_dir_exists

LOCAL_DIR = BASE_RESULT_DIR

def set_local_dir(path):
    global LOCAL_DIR
    LOCAL_DIR = path
    ensure_dir_exists(path)


def safe_imread(path, flag=0):
    if not path or not os.path.exists(path):
        return None
    try:
        data = np.fromfile(path, dtype=np.uint8)
        img = cv2.imdecode(data, flag)
        return img
    except Exception as e:
        logger(f'[IMREAD ERROR] {path} -> {e}')
        return None


def safe_cv_imread(path, flag=0):
    return safe_imread(path, flag)


def wait_file_ready(path, min_size=20 * 1024, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(path) and os.path.getsize(path) > min_size:
            return True
        time.sleep(0.5)
    return False


def download_file(remote, local):
    ensure_dir_exists(os.path.dirname(local))
    result = subprocess.run(['adb', 'pull', remote, local], capture_output=True)
    return os.path.exists(local)


def capture_screen(lang, name, save_dir=None):
    if save_dir is None:
        save_dir = LOCAL_DIR
    file_full_name = f"{lang}_{name}.png"
    local_file = os.path.join(save_dir, file_full_name)
    temp_remote_file = f"{REMOTE_DIR}/transfer_temp.png"

    ensure_dir_exists(save_dir)
    adb_shell(f'rm -f "{temp_remote_file}"')
    adb_shell('sync')
    time.sleep(5.0)

    logger('[PRE-SHOT] 清理缓存并等待页面渲染...')
    adb_shell('echo 3 > /proc/sys/vm/drop_caches')
    time.sleep(5.0)
    logger('[PRE-SHOT] 等待磁盘 IO 同步...')
    adb_shell('sync')
    time.sleep(5.0)

    logger(f'[SCREENSHOT] 正在请求新截图: {name}')
    shot_cmd = f'PATH=$PATH:/usr/bin:/bin ffmpeg -f fbdev -i /dev/fb0 -vf transpose=1 -frames:v 1 -y "{temp_remote_file}"'
    adb_shell(shot_cmd)

    max_wait = 360
    interval = 10.0
    found = False

    logger('[POLLING] 等待新文件生成...')
    for i in range(int(max_wait / interval)):
        check_cmd = f'[ -s "{temp_remote_file}" ] && echo "NEW_FILE_READY" || echo "STILL_WAITING"'
        adb_shell(check_cmd)
        matched = False
        start_t = time.time()
        while time.time() - start_t < 2.0:
            line = safe_read_line()
            if 'NEW_FILE_READY' in line:
                matched = True
                break
            if 'STILL_WAITING' in line:
                break
        if matched:
            logger('[MATCH] 检测到新图片已就绪')
            adb_shell(f'chmod 777 "{temp_remote_file}"')
            time.sleep(3.0)
            found = True
            break
        logger(f'  ... 检查中({(i + 1) * interval}s)，旧图已删，新图尚未生成 ...')
        time.sleep(interval)

    if found:
        if os.path.exists(local_file):
            os.remove(local_file)
        subprocess.run(['adb', 'pull', temp_remote_file, local_file], capture_output=True)
        if not wait_file_ready(local_file):
            logger(f'[ERROR] 文件未就绪: {local_file}')
            return None
        if os.path.exists(local_file) and os.path.getsize(local_file) > 0:
            logger(f'[SUCCESS] 截图已拉取并命名为: {file_full_name}')
            return local_file
        logger('[ERROR] pull成功但文件无效')
        return None

    if not os.path.exists(local_file) or os.path.getsize(local_file) == 0:
        logger('[ERROR] 截图文件无效')
        return None
    logger('[ERROR] 截图失败，可能超时或设备未响应')
    return None
