import time
import subprocess
from .constants import TOUCH_DEV, CLICK_PAUSE
from .utils import logger

shell_proc = None


def adb_shell(cmd, wait_time=1.0):
    global shell_proc
    if shell_proc is None:
        logger('[ERROR] shell 未初始化，无法下发指令')
        return ''

    try:
        if shell_proc.poll() is not None:
            logger('[FATAL] shell 已断开！需要重连')
            return ''

        clean_cmd = cmd.replace('adb shell ', '')
        shell_proc.stdin.write(f"{clean_cmd}\n")
        shell_proc.stdin.flush()
        time.sleep(wait_time)
        return 'Sent'
    except Exception as e:
        logger(f'[ERROR] 指令下发异常: {e}')
        return ''


def safe_read_line(timeout=2.0):
    start = time.time()
    while time.time() - start < timeout:
        try:
            if shell_proc and shell_proc.stdout:
                line = shell_proc.stdout.readline()
                if line:
                    return line.strip()
        except Exception:
            pass
        time.sleep(0.1)
    return ''


def smart_sleep(total_time, step=2):
    for _ in range(int(total_time / step)):
        time.sleep(step)


def safe_click(x, y, label='', max_retry=3):
    for i in range(max_retry):
        logger(f'[CLICK] {label} 第{i+1}次')
        click_with_evtest(x, y, page_label=label)
        time.sleep(3.0)
        return True


def click_with_evtest(x, y, page_label='未知页面', timeout=150.0, retries=3):
    global shell_proc
    for attempt in range(retries + 1):
        adb_shell('rm -f /tmp/ev.log')
        prefix = f'[RETRY {attempt}]' if attempt > 0 else '[ACTION]'
        logger(f'  {prefix} 正在点击 {page_label} 坐标 ({x}, {y}) 并监控回显...')

        adb_shell(f'/touch_control {TOUCH_DEV} {x} {y}')
        time.sleep(15.0)

        start_t = time.time()
        while time.time() - start_t < timeout:
            shell_proc.stdin.write('[ -s /tmp/ev.log ] && echo "HIT" || echo "MISS"\n')
            shell_proc.stdin.flush()
            line = safe_read_line()
            if 'HIT' in line:
                logger(f'  [CONFIRM] {page_label} 点击成功 (尝试第 {attempt + 1} 次)')
                return True
            time.sleep(0.5)

    error_msg = f'!!! [FATAL ERROR] 页面 [{page_label}] 坐标 ({x}, {y}) 彻底无响应，脚本停止运行 !!!'
    logger(error_msg)
    adb_shell('cat /tmp/ev.log')
    logger(error_msg)
    return False


def click(x, y, wait=CLICK_PAUSE):
    if x is None or y is None:
        return
    logger(f'  [ADB CLICK] 坐标: ({x}, {y})，等待 {wait}s')
    adb_shell(f'/touch_control {TOUCH_DEV} {x} {y}')
    time.sleep(wait)


def click_pulse(x, y, page_label='长按动作', times=25, interval=0.1, wait=5.0):
    if x is None or y is None:
        return False

    adb_shell('> /tmp/ev.log')
    logger(f'  [PULSE ACTION] 正在对 {page_label} ({x}, {y}) 下发 {times} 次脉冲长按...')

    try:
        for i in range(times):
            if i % 10 == 0 and i > 0:
                adb_shell('sync')
            adb_shell(f'/touch_control {TOUCH_DEV} {x} {y}', wait_time=0.01)
            time.sleep(interval)

        time.sleep(1.0)
        adb_shell('sync')

        shell_proc.stdin.write('[ -s /tmp/ev.log ] && echo "PULSE_HIT" || echo "PULSE_MISS"\n')
        shell_proc.stdin.flush()

        response = ''
        start_t = time.time()
        while time.time() - start_t < 3.0:
            line = safe_read_line()
            if 'PULSE_HIT' in line:
                response = 'HIT'
                break
            if 'PULSE_MISS' in line:
                response = 'MISS'
                break

        if response == 'HIT':
            logger(f'  [CONFIRM] {page_label} 脉冲长按内核响应成功')
            time.sleep(wait)
            return True
        logger(f'  [ERROR] {page_label} 脉冲长按内核无响应！')
        raise RuntimeError(f'长按 {page_label} 失败，脚本熔断退出')
    except Exception as e:
        logger(f'  [FATAL] 脉冲函数执行异常: {e}')
        raise
