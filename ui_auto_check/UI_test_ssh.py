import paramiko
import os
import time
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage # 用于插入图片

# ==========================
# 1. 更加保守的时间策略 (Ultra-Safe Mode)
# ==========================
DEVICE_IP = "172.23.4.110"
SSH_USER = "root"
SSH_PASS = "creality_2025"
TOUCH_DEV = "/dev/input/event0"
REMOTE_DIR = "/tmp/ui_screens"
LOCAL_DIR = r"C:\Users\119198\Downloads\ui_test\png"

# 时间大幅增加，确保指令物理执行+UI反馈
CLICK_PAUSE = 4.0       # 每次普通点击后的强制停顿 (增加到4秒)
STEP_PAUSE = 2.0        # 连续动作（如Reset）之间的停顿
SCREEN_RENDER = 5.0     # 页面跳转后，等待UI完全渲染完毕的时间 (增加到5秒)
FFMPEG_WRITE = 4.0      # 截图生成并写入Flash的缓冲
LANG_RELOAD = 12.0      # 语言切换后系统重载的最长等待 (增加到12秒)

# ==========================
# 2. 页面结构
# ==========================
HOME_RESET_STEPS = [(63, 617), (63, 617), (63, 617)]

UI_STRUCTURE = {
    "耗材页面": {"entry": (387, 570), "sub_pages": {"P1": None, "P2": (45, 291)}},
    "控制页面": {"entry": (383, 414), "sub_pages": {"P1": None, "P2": (69, 351)}},
    "文件列表": {"entry": (191, 390), "sub_pages": {"P1": None, "P2": (49, 287), "P3": (52, 108)}},
    "设置页面": {"entry": (406, 247), "sub_pages": {"P1": None, "P2": (66, 315), "P3": (54, 108)}},
    "帮助页面": {"entry": (408, 79), "sub_pages": {"P1": None, "P2": (33, 334), "P3": (58, 160)}}
}

LANGUAGES = ["中文（CN）", "英文（EN）English"]
LANGUAGE_COORDS = {
    "中文（CN）": (140, 529),
    "英文（EN）English": (145, 329),
}
GOTO_LANGUAGE_PAGE = [(406, 247), (54, 108), (381, 66), (261, 405)]

# ==========================
# 3. 工具函数
# ==========================
def ssh_connect():
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(DEVICE_IP, username=SSH_USER, password=SSH_PASS, timeout=30)
    return client

def sync_command(client, cmd):
    stdin, stdout, stderr = client.exec_command(cmd)
    exit_status = stdout.channel.recv_exit_status() # 同步等待
    return stdout.read().decode()

def click(client, x, y, post_wait=CLICK_PAUSE):
    if x is None or y is None: return
    print(f"  [ACTION] 同步点击 ({x}, {y}) ...")
    sync_command(client, f"/touch_control {TOUCH_DEV} {x} {y}")
    print(f"  [WAIT] 等待设备响应 {post_wait}s")
    time.sleep(post_wait)

def reset_to_home(client):
    print("[RESET] 缓慢返回首页...")
    for x, y in HOME_RESET_STEPS:
        click(client, x, y, post_wait=STEP_PAUSE)
    time.sleep(SCREEN_RENDER)

def download_file_via_cat(client, remote_path, local_path):
    stdin, stdout, stderr = client.exec_command(f"cat {remote_path}")
    content = stdout.read()
    if content:
        with open(local_path, "wb") as f:
            f.write(content)
        return True
    return False

def capture_screen(client, lang, name):
    time.sleep(2.0)
    remote_file = f"{REMOTE_DIR}/{lang}_{name}.png"
    sync_command(client, f"mkdir -p {REMOTE_DIR}")
    print(f"  [CAPTURE] 正在同步生成截图...")
    sync_command(client, f"ffmpeg -f fbdev -i /dev/fb0 -vf transpose=1 -frames:v 1 -y {remote_file} > /dev/null 2>&1")
    time.sleep(FFMPEG_WRITE)
    return remote_file

# ==========================
# 4. 主流程
# ==========================
def main():
    excel_path = os.path.join(LOCAL_DIR, "UI_Test_Report.xlsx")
    os.makedirs(LOCAL_DIR, exist_ok=True)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UI_Report"
    # 表头增加一列“预览图”
    ws.append(["语言", "一级页面", "分页", "本地路径", "预览图"])
    
    # 设置列宽，方便放图片
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 40 

    client = ssh_connect()
    
    try:
        for index, lang in enumerate(LANGUAGES):
            print(f"\n{'='*50}\n开始语言周期: {lang}\n{'='*50}")

            for main_page, config in UI_STRUCTURE.items():
                print(f"\n[PAGE] 进入一级页面: {main_page}")
                reset_to_home(client)
                
                ex, ey = config["entry"]
                click(client, ex, ey, post_wait=SCREEN_RENDER)
                
                for sub_name, sub_coord in config["sub_pages"].items():
                    print(f"  [SUB] 进入分页: {sub_name}")
                    if sub_coord:
                        click(client, sub_coord[0], sub_coord[1], post_wait=SCREEN_RENDER)
                    
                    full_name = f"{main_page}_{sub_name}"
                    remote_png = capture_screen(client, lang, full_name)
                    local_png = os.path.join(LOCAL_DIR, f"{lang}_{full_name}.png")
                    
                    if download_file_via_cat(client, remote_png, local_png):
                        # 准备写入 Excel
                        current_row = ws.max_row + 1
                        ws.append([lang, main_page, sub_name, local_png, ""])
                        
                        # --- 插入图片逻辑 ---
                        try:
                            img = OpenpyxlImage(local_png)
                            # 缩小图片尺寸以适应单元格 (例如 240x135)
                            img.width, img.height = 280, 175
                            # 设置行高以适应图片高度
                            ws.row_dimensions[current_row].height = 110 
                            # 将图片锚定到 E 列对应行
                            ws.add_image(img, f'E{current_row}')
                        except Exception as e:
                            print(f"  [WARN] 图片插入Excel失败: {e}")

                        # 实时保存 (捕获权限错误)
                        try:
                            wb.save(excel_path)
                            print(f"  [SUCCESS] 数据及预览图已入库: {full_name}")
                        except PermissionError:
                            print(f"  [ERROR] 无法保存Excel！请关闭已经打开的 {excel_path} 文件！")
                        
                        sync_command(client, f"rm -f {remote_png}")
                    else:
                        print(f"  [ERROR] 下载失败: {full_name}")

            # 语言切换
            if index < len(LANGUAGES) - 1:
                next_lang = LANGUAGES[index + 1]
                print(f"\n[LANG_SWITCH] 正在执行语言切换 -> {next_lang}")
                reset_to_home(client)
                for x, y in GOTO_LANGUAGE_PAGE:
                    click(client, x, y, post_wait=CLICK_PAUSE)
                l_x, l_y = LANGUAGE_COORDS[next_lang]
                click(client, l_x, l_y, post_wait=LANG_RELOAD)

    finally:
        client.close()
        print(f"\n[FINISHED] 指令同步下发完毕。请检查: {excel_path}")

if __name__ == "__main__":
    main()