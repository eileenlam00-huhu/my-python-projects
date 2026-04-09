import os
import time
import subprocess
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment

# ==========================
# 1. 更加保守的时间策略 (ADB 模式)
# ==========================
TOUCH_DEV = "/dev/input/event0"
REMOTE_DIR = "/tmp/ui_screens"
LOCAL_DIR = r"C:\Users\119198\Downloads\ui_test\png"

# 时间配置：确保 UI 渲染与设备响应
CLICK_PAUSE = 3.5      
STEP_PAUSE = 2.0       
SCREEN_RENDER = 5.0    
FFMPEG_WRITE = 4.0     
LANG_RELOAD = 15.0     

# ==========================
# 2. 页面结构与语言定义
# ==========================
HOME_RESET_STEPS = [(63, 617), (63, 617), (63, 617)]

UI_STRUCTURE = {
    "耗材页面": {"entry": (387, 570), "sub_pages": {"P1": None, "P2": (45, 291)}},
    "控制页面": {"entry": (383, 414), "sub_pages": {"P1": None, "P2": (69, 351)}},
    "文件列表": {"entry": (191, 390), "sub_pages": {"P1": None, "P2": (49, 287), "P3": (52, 108)}},
    "设置页面": {"entry": (406, 247), "sub_pages": {"P1": None, "P2": (66, 315), "P3": (54, 108)}},
    "帮助页面": {"entry": (408, 79), "sub_pages": {"P1": None, "P2": (33, 334), "P3": (58, 160)}}
}

# 确保第一个是基准语言
LANGUAGES = ["中文（CN）", "英文（EN）English"]
LANGUAGE_COORDS = {
    "中文（CN）": (140, 529),
    "英文（EN）English": (145, 329),
}
GOTO_LANGUAGE_PAGE = [(406, 247), (54, 108), (381, 66), (261, 405)]

# ==========================
# 3. ADB 工具函数
# ==========================
def adb_shell(cmd):
    """
    增强版 ADB 指令下发：
    1. 增加 check=True，如果命令非法直接报错
    2. 检查返回内容，如果没有设备则报错
    """
    full_cmd = ["adb", "shell", cmd]
    try:
        # 增加 capture_output 以捕获错误信息
        result = subprocess.run(full_cmd, capture_output=True, text=True, encoding='utf-8', timeout=30)
        
        # 如果 adb 返回错误码，或者 stderr 里有设备找不到的提示
        if result.returncode != 0 or "device not found" in result.stderr:
            error_msg = result.stderr.strip() if result.stderr else "未知 ADB 错误"
            print(f"\n[FATAL ERROR] ADB 指令失败: {error_msg}")
            raise ConnectionError(f"无法连接到设备，错误信息: {error_msg}")
            
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        raise TimeoutError("ADB 指令执行超时，设备可能已断开连接。")
    except FileNotFoundError:
        raise FileNotFoundError("系统中未找到 adb 执行程序，请检查环境变量 PATH。") 

def click(x, y, wait=CLICK_PAUSE):
    if x is None or y is None: return
    print(f"  [ADB CLICK] 点击坐标: ({x}, {y})")
    # 优先尝试使用 touch_control 保持与原逻辑一致
    adb_shell(f"/touch_control {TOUCH_DEV} {x} {y}")    
    time.sleep(wait)

def reset_to_home():
    print("[RESET] 正在缓慢返回首页...")
    for x, y in HOME_RESET_STEPS:
        click(x, y, wait=STEP_PAUSE)
    time.sleep(1.0)

def capture_screen(lang, name):
    time.sleep(1.5)
    remote_file = f"{REMOTE_DIR}/{lang}_{name}.png"
    adb_shell(f"mkdir -p {REMOTE_DIR}")
    
    print(f"  [CAPTURE] 正在旋转截图...")
    cmd = f"ffmpeg -f fbdev -i /dev/fb0 -vf transpose=1 -frames:v 1 -y {remote_file} > /dev/null 2>&1"
    adb_shell(cmd)
    
    time.sleep(FFMPEG_WRITE)
    return remote_file

def download_file(remote, local):
    """使用 ADB Pull 下载文件"""
    # 确保本地目录存在
    os.makedirs(os.path.dirname(local), exist_ok=True)
    result = subprocess.run(["adb", "pull", remote, local], capture_output=True)
    return os.path.exists(local)

# ==========================
# 4. 主流程
# ==========================
def main():
    excel_path = os.path.join(LOCAL_DIR, "UI_Review_ADB_Report.xlsx")
    os.makedirs(LOCAL_DIR, exist_ok=True)

    # --- 物理连接强制检查 ---
    print("正在检查设备连接状态...")
    check_conn = subprocess.run(["adb", "get-state"], capture_output=True, text=True)
    
    # 正常连接时 stdout 应为 'device\n'
    if "device" not in check_conn.stdout:
        print("\n" + "!"*50)
        print("错误：未检测到任何有效的 ADB 设备！")
        print("1. 请确保已通过 USB 线连接设备")
        print("2. 请确保设备已开启 'USB 调试' 模式")
        print("3. 请在命令行输入 'adb devices' 确认列表不为空")
        print("!"*50 + "\n")
        # 强制退出，不再往下跑
        return  
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UI对比报告"
    ws.append(["语言", "一级页面", "分页", "基准(中文)", "待测(当前)"])
    
    # 设置列宽
    ws.column_dimensions['D'].width = 35 
    ws.column_dimensions['E'].width = 35 

    base_lang = LANGUAGES[0]

    # 检查 ADB 设备状态
    check_dev = subprocess.run(["adb", "get-state"], capture_output=True, text=True)
    if "device" not in check_dev.stdout:
        print("错误：未检测到 ADB 设备，请确保 USB 已连接并开启调试！")
        return

    try:
        for index, lang in enumerate(LANGUAGES):
            print(f"\n{'='*60}\n当前语言周期: {lang} ({index+1}/{len(LANGUAGES)})\n{'='*60}")

            # 语言切换逻辑 (除了第一个语言外执行)
            if index > 0:
                print(f"[LANG] 切换至 {lang}...")
                reset_to_home()
                for x, y in GOTO_LANGUAGE_PAGE:
                    click(x, y, wait=CLICK_PAUSE)
                lx, ly = LANGUAGE_COORDS[lang]
                click(lx, ly, wait=LANG_RELOAD)

            # 遍历页面结构
            for main_page, config in UI_STRUCTURE.items():
                print(f"\n[MENU] 进入: {main_page}")
                reset_to_home()
                
                # 进入一级页面
                ex, ey = config["entry"]
                click(ex, ey, wait=SCREEN_RENDER)
                
                for sub_name, sub_coord in config["sub_pages"].items():
                    print(f"  [SUB] 分页: {sub_name}")
                    if sub_coord:
                        click(sub_coord[0], sub_coord[1], wait=SCREEN_RENDER)
                    
                    # 截图与导出
                    file_id = f"{main_page}_{sub_name}"
                    remote_png = capture_screen(lang, file_id)
                    local_png = os.path.join(LOCAL_DIR, f"{lang}_{file_id}.png")
                    
                    if download_file(remote_png, local_png):
                        current_row = ws.max_row + 1
                        ws.append([lang, main_page, sub_name, "", ""])
                        ws.row_dimensions[current_row].height = 160 # 设置行高以适应图片显示
                        
                        # 插入图片逻辑
                        try:
                            # 1. 插入当前测试语言图 (E列)
                            img_curr = OpenpyxlImage(local_png)
                            img_curr.width, img_curr.height = 240, 200 # 根据旋转后的比例调整
                            ws.add_image(img_curr, f'E{current_row}')
                            
                            # 2. 如果当前不是中文，自动寻找中文基准图并并列 (D列)
                            if lang != base_lang:
                                base_path = os.path.join(LOCAL_DIR, f"{base_lang}_{file_id}.png")
                                if os.path.exists(base_path):
                                    img_base = OpenpyxlImage(base_path)
                                    img_base.width, img_base.height = 240, 200
                                    ws.add_image(img_base, f'D{current_row}')
                                    
                        except Exception as e:
                            print(f"  [WARN] 图片插入Excel失败: {e}")

                        # 实时保存保护
                        try:
                            wb.save(excel_path)
                            print(f"  [SUCCESS] {file_id} 已记录到报告")
                        except PermissionError:
                            print(f"  [CRITICAL] 无法保存 Excel！请关闭 {excel_path} 文件后再继续！")
                        
                        # 清理设备端临时文件
                        adb_shell(f"rm -f {remote_png}")
                    else:
                        print(f"  [ERROR] 文件拉取失败: {file_id}")

    except Exception as e:
        print(f"程序运行中断: {e}")
    finally:
        print(f"\n[FINISHED] 脚本执行完毕。最终报告: {excel_path}")

if __name__ == "__main__":
    main()