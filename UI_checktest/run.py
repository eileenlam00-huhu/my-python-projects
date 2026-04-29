"""
UI 自动化检查主程序

这个模块是 UI 自动化测试的主入口程序。
功能包括：
- 多语言 UI 界面检查
- 截图捕获和对比
- Excel 报告生成
- ADB 设备控制

使用方法：
1. 连接 ADB 设备
2. 运行 UI_checktest run.py
3. 选择要测试的语言
4. 等待自动测试完成

注意：需要安装相关依赖包
"""

import os
import subprocess
import time
import cv2
import numpy as np
from datetime import datetime
from openpyxl.drawing.image import Image as OpenpyxlImage

# 添加项目根目录到路径，以便导入 UI_checktest 包
# 这修复了相对导入问题
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# 导入项目模块
# 改为绝对导入，避免相对导入错误
from UI_checktest.constants import (
    BASE_RESULT_DIR,      # 结果存放目录
    SCREEN_RENDER,        # 页面渲染等待时间
    LANG_RELOAD,          # 语言切换重载时间
    HOME_RESET_STEPS,     # 返回首页的点击步骤
    GOTO_LANGUAGE_PAGE,   # 进入语言设置的步骤
    LANG_PAGE_NEXT_BTN,   # 语言列表翻页按钮
)
from UI_checktest.login import init_persistent_shell, start_evtest_background  # 登录和初始化
from UI_checktest.click import click_with_evtest  # 点击控制
from UI_checktest.page import UI_STRUCTURE, reset_to_home  # 页面结构和导航
from UI_checktest.multilanguage import select_languages, check_language_switch  # 多语言处理
from UI_checktest.screenshot import capture_screen, set_local_dir, safe_imread  # 截图功能
from UI_checktest.compare import analyze_ui_diff  # 图像对比
from UI_checktest.report import create_report_workbook, append_image  # 报告生成
from UI_checktest.html_report import create_html_report  # HTML报告生成
from UI_checktest.utils import logger  # 日志记录


def main():
    run_id = datetime.now().strftime('%Y%m%d_%H%M%S')
    current_run_dir = os.path.join(BASE_RESULT_DIR, f'Run_{run_id}')
    os.makedirs(current_run_dir, exist_ok=True)

    subprocess.run(['adb', 'kill-server'], capture_output=True)
    subprocess.run(['adb', 'start-server'], capture_output=True)
    time.sleep(2.0)

    check_conn = subprocess.run(['adb', 'get-state'], capture_output=True, text=True)
    if 'device' not in check_conn.stdout:
        logger('[ERROR] 未检测到有效的 ADB 设备，程序退出！')
        return

    init_persistent_shell()
    time.sleep(2.0)
    start_evtest_background()

    target_langs = select_languages()
    if not target_langs:
        return

    set_local_dir(current_run_dir)
    logger(f'[INIT] 本次测试结果将存放在: {current_run_dir}')

    excel_path = os.path.join(current_run_dir, f'UI_Report_{run_id}.xlsx')
    wb, ws = create_report_workbook(excel_path)

    # 创建HTML报告
    html_report = create_html_report(current_run_dir, run_id)
    html_report.set_total_languages(len(target_langs))

    # --- 修改点 1: 初始化基准图字典，用于存储中文（CN）的所有页面截图 ---
    baseline_images = {}

    try:
        # 遍历所有语言
        for lang_idx, (lang_id, lang_name, lang_coord) in enumerate(target_langs):
            # --- 修改点 2: 明确判断是否为基准语言（中文） ---
            is_baseline_lang = (lang_name == '中文（CN）')
            logger(f'\n================ 当前准备切换的语言: {lang_name} (是否基准语言: {is_baseline_lang}) ================')

            # 返回首页
            for rx, ry in HOME_RESET_STEPS:
                click_with_evtest(rx, ry, page_label='回首页')
            time.sleep(5)

            # 进入语言设置
            for gx, gy in GOTO_LANGUAGE_PAGE:
                click_with_evtest(gx, gy, page_label='进入语言设置步骤')
                time.sleep(10.0)

            # 点击语言列表或翻页
            if lang_id < 13:
                click_with_evtest(169, 49, page_label='语言唤醒')
            else:
                click_with_evtest(LANG_PAGE_NEXT_BTN[0], LANG_PAGE_NEXT_BTN[1], page_label='翻页')

            time.sleep(10)
            logger('[SCREEN] 语言切换前，开始截图1')
            current_img1 = capture_screen(lang_name, 'LANG_BEFORE')

            # 切换语言
            click_with_evtest(lang_coord[0], lang_coord[1], page_label=f'切换至_{lang_name}')
            logger('[WAIT] 等待语言刷新...')
            time.sleep(LANG_RELOAD)

            logger('[SCREEN] 语言切换后，开始截图2')
            current_img2 = capture_screen(lang_name, 'LANG_AFTER')

            # --- 修改点 3: 语言切换验证的基准图处理 ---
            # 如果是基准语言，保存当前截图作为语言切换的基准图
            if is_baseline_lang:
                baseline_images['LANG_SWITCH'] = current_img2
            
            # 从字典中获取语言切换的基准图
            lang_base_img = baseline_images.get('LANG_SWITCH')
            
            # 如果不是基准语言且找不到基准图，跳过验证
            if not lang_base_img and not is_baseline_lang:
                logger('[WARN] 未找到语言切换的基准图，跳过验证')
            else:
                # 执行对比
                ok1, r1 = check_language_switch(lang_base_img, current_img1)
                ok2, r2 = check_language_switch(lang_base_img, current_img2)
                ok = ok1 or ok2
                ratio = max(r1, r2)
                status = 'PASS' if ok else 'WARN'
                if current_img1 is None or current_img2 is None:
                    status = 'FAIL_NO_IMAGE'

                logger(f'[LANG CHECK] diff={ratio:.4f}')
                if not ok:
                    logger(f'[WARN] 语言可能未切换或已是当前语言: {lang_name} (diff={ratio:.4f})')
                else:
                    logger('✅ 语言切换成功')

                # 写入 Excel
                ws.append([
                    lang_id,
                    lang_name,
                    'LANG',
                    '语言设置',
                    '切换验证',
                    lang_base_img, # 使用动态获取的基准图
                    current_img2,
                    ratio,
                    status
                ])
                append_image(ws, ws.max_row, lang_base_img, 'F')
                append_image(ws, ws.max_row, current_img2, 'G')

                # 更新HTML报告
                html_report.add_language_test(lang_id, lang_name, status, ratio, current_img1, current_img2)

            # 返回首页
            for i in range(2):
                click_with_evtest(63, 617, page_label=f'语言切换后返回首页_{i + 1}')
                time.sleep(15.0)

            last_page_was_home = True
            subprocess.run(['adb', 'shell', 'sync; echo 3 > /proc/sys/vm/drop_caches'], capture_output=True)
            ui_status = 'UNKNOWN'
            ui_diff_ratio = 0

            # 遍历所有UI页面
            for main_page, config in UI_STRUCTURE.items():
                logger(f'\n--- 正在处理一级页面: {main_page} ---')
                if main_page == '首页':
                    last_page_was_home = True
                else:
                    if not last_page_was_home:
                        reset_to_home(main_page)
                    ex, ey = config.get('entry', (None, None))
                    if ex is not None:
                        click_with_evtest(ex, ey, page_label=f'进入_{main_page}')
                        last_page_was_home = False

                # 遍历子页面
                for sub_name, actions in config.get('sub_pages', {}).items():
                    logger(f'  [SUB] {sub_name}')
                    # 构建当前页面的唯一键名
                    page_key = f'{main_page}_{sub_name}'
                    
                    if isinstance(actions, dict):
                        sub_entry = actions.get('entry', (None, None))
                        if sub_entry and sub_entry[0] is not None:
                            click_with_evtest(sub_entry[0], sub_entry[1], page_label=f'{main_page}_{sub_name}_进入')
                            time.sleep(SCREEN_RENDER)

                        img_path = capture_screen(lang_name, page_key)
                        if img_path:
                            # --- 修改点 4: 核心逻辑：根据是否为基准语言处理基准图 ---
                            current_base_img_path = None
                            if is_baseline_lang:
                                # 基准语言：当前截图即为基准图，存入字典
                                current_base_img_path = img_path
                                baseline_images[page_key] = img_path
                            else:
                                # 非基准语言：从字典中取出基准图
                                current_base_img_path = baseline_images.get(page_key)
                                if not current_base_img_path:
                                    logger(f'[WARN] 未找到 {page_key} 的基准图，跳过此页面的对比测试')
                                    continue  # 跳过当前页面的测试

                            # 执行对比
                            ui_status = analyze_ui_diff(current_base_img_path, img_path)
                            try:
                                diff_img = safe_imread(current_base_img_path, 0)
                                test_img = safe_imread(img_path, 0)
                                diff_ratio = -1
                                if diff_img is not None and test_img is not None:
                                    diff_ratio = cv2.absdiff(diff_img, test_img)
                                    diff_ratio = np.count_nonzero(diff_ratio) / diff_ratio.size
                            except Exception:
                                diff_ratio = -1
                            
                            # 写入 Excel
                            ws.append([
                                lang_id,
                                lang_name,
                                'UI',
                                main_page,
                                sub_name,
                                current_base_img_path, # 使用动态计算的基准图路径
                                img_path,
                                diff_ratio,
                                ui_status
                            ])
                            append_image(ws, ws.max_row, current_base_img_path, 'F')
                            append_image(ws, ws.max_row, img_path, 'G')

                        # 处理二级子页面 (SUB2)
                        for sub2_name, sub2_actions in actions.get('sub_pages', {}).items():
                            logger(f'    [SUB2] {sub2_name}')
                            # 构建二级页面的唯一键名
                            page_key_sub2 = f'{main_page}_{sub_name}_{sub2_name}'
                            
                            ax, ay = None, None
                            if isinstance(sub2_actions, tuple):
                                if len(sub2_actions) == 3:
                                    _, ax, ay = sub2_actions
                                else:
                                    ax, ay = sub2_actions
                            elif isinstance(sub2_actions, dict):
                                ax, ay = sub2_actions.get('entry', (None, None))
                            elif isinstance(sub2_actions, list) and sub2_actions:
                                act = sub2_actions[0]
                                if isinstance(act, tuple):
                                    if len(act) == 3:
                                        _, ax, ay = act
                                    else:
                                        ax, ay = act
                            if ax is not None:
                                click_with_evtest(ax, ay, page_label=f'{sub_name}_{sub2_name}')
                            time.sleep(SCREEN_RENDER)

                            img_path = capture_screen(lang_name, page_key_sub2)
                            if not img_path:
                                logger('[SKIP] SUB2截图失败')
                                continue
                            
                            # --- 修改点 5: SUB2 的基准图处理 ---
                            current_base_img_path = None
                            if is_baseline_lang:
                                current_base_img_path = img_path
                                baseline_images[page_key_sub2] = img_path
                            else:
                                current_base_img_path = baseline_images.get(page_key_sub2)
                                if not current_base_img_path:
                                    logger(f'[WARN] 未找到 {page_key_sub2} 的基准图，跳过此页面的对比测试')
                                    continue

                            # 执行对比
                            ui_status = analyze_ui_diff(current_base_img_path, img_path)
                            try:
                                diff_img = safe_imread(current_base_img_path, 0)
                                test_img = safe_imread(img_path, 0)
                                if diff_img is None or test_img is None:
                                    diff_ratio = -1
                                else:
                                    diff_ratio = cv2.absdiff(diff_img, test_img)
                                    diff_ratio = np.count_nonzero(diff_ratio) / diff_ratio.size
                            except Exception as e:
                                logger(f'[DIFF ERROR] {e}')
                                diff_ratio = -1

                            # 写入 Excel
                            ws.append([
                                lang_id,
                                lang_name,
                                'UI',
                                main_page,
                                f'{sub_name}_{sub2_name}',
                                current_base_img_path, # 使用动态计算的基准图路径
                                img_path,
                                diff_ratio,
                                ui_status
                            ])
                            append_image(ws, ws.max_row, current_base_img_path, 'F')
                            append_image(ws, ws.max_row, img_path, 'G')

                            # 更新HTML报告
                            html_report.add_ui_test(lang_name, 'UI', main_page, f'{sub_name}_{sub2_name}', ui_status, diff_ratio, current_base_img_path, img_path)
                        continue

                    # 处理列表项页面
                    if actions is None:
                        actions = [(None, None)]
                    elif isinstance(actions, tuple):
                        actions = [actions]
                    elif not isinstance(actions, list):
                        actions = [actions]

                    for i, act in enumerate(actions):
                        ax, ay = None, None
                        if act is None:
                            act = (None, None)
                        if isinstance(act, tuple):
                            if len(act) == 3:
                                _, ax, ay = act
                            else:
                                ax, ay = act
                        if ax is not None:
                            click_with_evtest(ax, ay, page_label=f'{main_page}_{sub_name}')
                        time.sleep(SCREEN_RENDER)

                        # 构建列表项页面的唯一键名
                        page_key_list = f'{main_page}_{sub_name}_{i}'
                        
                        img_path = capture_screen(lang_name, page_key_list)
                        
                        # --- 修改点 6: 列表项的基准图处理 ---
                        current_base_img_path = None
                        if is_baseline_lang:
                            current_base_img_path = img_path
                            baseline_images[page_key_list] = img_path
                        else:
                            current_base_img_path = baseline_images.get(page_key_list)
                            if not current_base_img_path:
                                logger(f'[WARN] 未找到 {page_key_list} 的基准图，跳过此页面的对比测试')
                                continue

                        # 执行对比
                        status = analyze_ui_diff(current_base_img_path, img_path) if img_path else 'Skip'
                        
                        # 写入 Excel
                        ws.append([
                            lang_id,
                            lang_name,
                            'UI',
                            main_page,
                            f'{sub_name}_{i}',
                            current_base_img_path, # 使用动态计算的基准图路径
                            img_path,
                            'OK' if img_path else 'FAIL',
                            status
                        ])
                        append_image(ws, ws.max_row, current_base_img_path, 'F')
                        append_image(ws, ws.max_row, img_path, 'G')

                        # 更新HTML报告
                        html_report.add_ui_test(lang_name, 'UI', main_page, f'{sub_name}_{i}', status, -1, current_base_img_path, img_path)
                    logger(f'[LANG DONE] {lang_name}')
    finally:
        # --- 修改点 7: 统一保存 Excel，解决列宽问题 ---
        logger(f'\n[FINISHED] 任务结束。报告: {excel_path}')
        wb.save(excel_path)
        logger('[EXCEL] 报告已保存')

        # 标记HTML报告为完成状态
        html_report.mark_completed()
        logger('[HTML] 报告已完成')


if __name__ == '__main__':
    main()
