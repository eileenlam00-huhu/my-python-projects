import subprocess
import time
from . import click
from .constants import DEVICE_USER, DEVICE_PASS, REMOTE_DIR
from .utils import logger


def init_persistent_shell():
    logger('[INIT] 正在建立持久化 Shell 连接...')
    click.shell_proc = subprocess.Popen(
        ['adb', 'shell'],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding='utf-8',
        errors='replace',
        bufsize=1
    )

    def login_handshake(expect_str, send_str, label):
        logger(f'[AUTH] 正在等待 {label} 提示符...')
        start_time = time.time()
        buffer = ''
        while time.time() - start_time < 90:
            char = click.shell_proc.stdout.read(1)
            if char:
                buffer += char
                if expect_str.lower() in buffer.lower():
                    logger(f'[DEVICE] 匹配到提示符: {expect_str}')
                    time.sleep(1.0)
                    click.shell_proc.stdin.write(f'{send_str}\n')
                    click.shell_proc.stdin.flush()
                    return True
            else:
                time.sleep(0.1)
        return False

    click.shell_proc.stdin.write('\n')
    click.shell_proc.stdin.flush()

    if not login_handshake('login:', DEVICE_USER, '用户名'):
        logger('[FATAL] 等待登录提示符超时！')
        return

    if not login_handshake('Password:', DEVICE_PASS, '密码'):
        logger('[FATAL] 等待密码提示符超时！')
        return

    logger('[AUTH] 正在确认 Root 权限...')
    time.sleep(3.0)
    click.shell_proc.stdin.write('id\n')
    click.shell_proc.stdin.flush()

    logger('[SUCCESS] 交互式登录序列完成')
    click.shell_proc.stdin.write(f'mkdir -p {REMOTE_DIR} && chmod 777 {REMOTE_DIR}\n')
    click.shell_proc.stdin.flush()


def start_evtest_background():
    click.adb_shell('( /sbin/evtest /dev/input/event0 < /dev/null > /tmp/ev.log 2>&1 & )')
    logger('[INIT] evtest 监听已通过子 Shell 启动')
