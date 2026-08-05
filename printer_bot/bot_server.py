#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
飞书机器人服务 - 多用户隔离版

每个用户独立管理自己的表格配置，互不影响。
定时任务按用户各自的设置触发，结果只通知对应用户。

启动方式: python bot_server.py
"""

import json
import re
import os
import threading
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
import requests

import lark_oapi as lark
from lark_oapi.api.im.v1 import *

import printer_monitor as pm

# ========== 配置 ==========

CONFIG_FILE = "bot_config.json"

config = {}
scheduler = BackgroundScheduler()

# ========== 配置管理 ==========

def load_config():
    global config

    # 先尝试从 .env 加载环境变量
    try:
        from dotenv import load_dotenv
        load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
    except ImportError:
        pass

    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {
            "feishu_app_id": "",
            "feishu_app_secret": "",
            "users": {}
        }
        save_config()
        print(f"⚠️  已生成 {CONFIG_FILE}，请填入 feishu_app_id/secret 后重启")

    if "users" not in config:
        config["users"] = {}

    # 优先使用环境变量中的密钥
    if os.environ.get("FEISHU_APP_ID"):
        config["feishu_app_id"] = os.environ["FEISHU_APP_ID"]
    if os.environ.get("FEISHU_APP_SECRET"):
        config["feishu_app_secret"] = os.environ["FEISHU_APP_SECRET"]

    if not config.get("feishu_app_id") or not config.get("feishu_app_secret"):
        raise ValueError(f"请在 .env 或 {CONFIG_FILE} 中配置 feishu_app_id 和 feishu_app_secret")

    return config


def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def get_user_config(open_id):
    """获取用户配置，不存在则创建"""
    users = config.setdefault("users", {})
    if open_id not in users:
        users[open_id] = {
            "sheets": [],
            "schedule_times": []
        }
        save_config()
    return users[open_id]

# ========== 飞书 API ==========

def get_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": config["feishu_app_id"], "app_secret": config["feishu_app_secret"]}
    resp = requests.post(url, json=payload)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"获取 token 失败: {data}")
    return data["tenant_access_token"]


def reply_text(token, open_id, text):
    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "receive_id": open_id,
        "msg_type": "text",
        "content": json.dumps({"text": text})
    }
    requests.post(url, headers=headers, json=payload)

# ========== 表格操作 ==========

def get_spreadsheet_token(token, wiki_token):
    url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={wiki_token}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"获取Wiki节点失败: code={data.get('code')}, msg={data.get('msg')}")
    node = data["data"]["node"]
    if node.get("obj_type") != "sheet":
        raise ValueError(f"节点不是电子表格，而是 '{node.get('obj_type')}'")
    return node["obj_token"]


def get_sheet_id(token, spreadsheet_token, sheet_name):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"获取子表失败: {data.get('msg')}")
    for s in data["data"]["sheets"]:
        if s["title"] == sheet_name:
            return s["sheet_id"]
    titles = [s["title"] for s in data["data"]["sheets"]]
    raise ValueError(f"找不到子表 '{sheet_name}'，现有: {titles}")


def extract_cell_text(cell):
    """从飞书单元格提取纯文本"""
    if cell is None:
        return ""
    if isinstance(cell, str):
        return cell.strip()
    if isinstance(cell, (int, float)):
        return str(cell).strip()
    if isinstance(cell, list):
        texts = []
        for item in cell:
            if isinstance(item, dict):
                texts.append(item.get("text", item.get("link", "")))
            else:
                texts.append(str(item))
        return "".join(texts).strip()
    return str(cell).strip()


def read_printers_from_sheet(token, spreadsheet_token, sheet_name):
    """动态读取打印机列表：根据表头定位 名称/IP/型号/负责人/邮箱 列"""
    sheet_id = get_sheet_id(token, spreadsheet_token, sheet_name)

    # 先读表头确定各列位置
    header_range = f"{sheet_id}!A1:Z1"
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{header_range}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"读取表头失败: {data.get('msg')}")

    header_rows = data.get("data", {}).get("valueRange", {}).get("values", [])
    if not header_rows:
        raise ValueError("表头行为空")

    header = [extract_cell_text(c) for c in header_rows[0]]

    # 通过关键字匹配列位置
    col_map = {}
    keywords = {"名称": "name", "IP": "ip", "型号": "model", "负责人": "owner_name", "邮箱": "owner_email"}
    for i, h in enumerate(header):
        for kw, field in keywords.items():
            if kw.lower() in h.lower():
                col_map[field] = i
                break

    # 确定数据读取范围（到邮箱列为止）
    email_col_idx = col_map.get("owner_email")
    if email_col_idx is None:
        raise ValueError(f"表头中未找到'邮箱'列，表头: {header}")

    end_col_letter = chr(65 + email_col_idx)  # 0-based -> A,B,C...
    range_str = f"{sheet_id}!A2:{end_col_letter}200"
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}"
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"读取表格失败: {data.get('msg')}")

    rows = data.get("data", {}).get("valueRange", {}).get("values", [])
    printers = []
    for row in rows:
        if not row or not row[0]:
            break
        name = extract_cell_text(row[col_map["name"]]) if "name" in col_map and len(row) > col_map["name"] else ""
        ip = extract_cell_text(row[col_map["ip"]]) if "ip" in col_map and len(row) > col_map["ip"] else ""
        model = extract_cell_text(row[col_map["model"]]) if "model" in col_map and len(row) > col_map["model"] else ""
        owner_name = extract_cell_text(row[col_map["owner_name"]]) if "owner_name" in col_map and len(row) > col_map["owner_name"] else ""
        owner_email = extract_cell_text(row[col_map["owner_email"]]) if "owner_email" in col_map and len(row) > col_map["owner_email"] else ""
        if name and ip:
            printers.append({
                "name": name, "ip": ip, "model": model,
                "owner_name": owner_name or "未指定",
                "owner_email": owner_email
            })
    return printers

# ========== 执行监控 ==========

def run_single_sheet(token, sheet_cfg):
    """对单个表格执行监控，返回打印机数量"""
    wiki_token = sheet_cfg["wiki_token"]
    sheet_name = sheet_cfg["sheet_name"]

    spreadsheet_token = get_spreadsheet_token(token, wiki_token)
    printers = read_printers_from_sheet(token, spreadsheet_token, sheet_name)
    if not printers:
        return 0

    pm.FEISHU_APP_ID = config["feishu_app_id"]
    pm.FEISHU_APP_SECRET = config["feishu_app_secret"]
    pm.PRINTERS = printers
    pm.SPREADSHEET_TOKEN = spreadsheet_token
    pm.WIKI_NODE_TOKEN = wiki_token
    pm.SHEET_NAME = sheet_name
    pm.START_ROW = 2

    pm.main()
    return len(printers)


def run_monitor_for_user(open_id, sheet_index=None):
    """为指定用户执行监控"""
    try:
        token = get_token()
    except Exception as e:
        print(f"❌ 获取 token 失败: {e}")
        return

    user_cfg = get_user_config(open_id)
    sheets = user_cfg.get("sheets", [])

    if not sheets:
        reply_text(token, open_id, "❌ 你还没添加表格，请先:\n添加表格 <URL> <子表名>")
        return

    if sheet_index is not None:
        if sheet_index < 1 or sheet_index > len(sheets):
            reply_text(token, open_id, f"❌ 编号无效，你当前共 {len(sheets)} 个表格")
            return
        target_sheets = [sheets[sheet_index - 1]]
    else:
        target_sheets = sheets

    total_printers = 0
    success_count = 0
    errors = []

    for sheet_cfg in target_sheets:
        alias = sheet_cfg.get("name", sheet_cfg["sheet_name"])
        try:
            count = run_single_sheet(token, sheet_cfg)
            total_printers += count
            if count > 0:
                success_count += 1
            print(f"   ✅ [{alias}] 完成，{count} 台")
        except Exception as e:
            errors.append(f"[{alias}] {e}")
            print(f"   ❌ [{alias}] 失败: {e}")
            import traceback
            traceback.print_exc()

    msg = f"✅ 监控完成！\n查询 {success_count}/{len(target_sheets)} 个表格，共 {total_printers} 台打印机"
    if errors:
        msg += "\n\n❌ 失败:\n" + "\n".join(errors)
    reply_text(token, open_id, msg)


def run_all_scheduled():
    """定时任务：遍历所有用户，按各自设定执行"""
    # 此函数由 scheduler 内部按时间匹配调用
    pass

# ========== 定时任务 ==========

def setup_scheduler():
    """根据所有用户的定时配置设置 scheduler"""
    scheduler.remove_all_jobs()

    users = config.get("users", {})
    for open_id, user_cfg in users.items():
        times = user_cfg.get("schedule_times", [])
        for t in times:
            try:
                hour, minute = t.split(":")
                job_id = f"{open_id}_{t}"
                scheduler.add_job(
                    run_monitor_for_user,
                    CronTrigger(hour=int(hour), minute=int(minute)),
                    args=[open_id],
                    id=job_id,
                    replace_existing=True
                )
            except Exception as e:
                print(f"   ❌ 设置定时 {open_id} {t} 失败: {e}")

    job_count = len(scheduler.get_jobs())
    print(f"   ⏰ 共 {job_count} 个定时任务")

# ========== 消息处理 ==========

def handle_message(open_id, text):
    """处理用户消息"""
    print(f"📩 收到消息: {text} (from: {open_id})")
    token = get_token()

    if text.startswith("查询") or text in ("执行", "run", "监控"):
        user_cfg = get_user_config(open_id)
        if not user_cfg.get("sheets"):
            reply_text(token, open_id, "❌ 你还没添加表格:\n添加表格 <URL> <子表名>")
            return
        num_match = re.search(r'(\d+)', text)
        sheet_index = int(num_match.group(1)) if num_match else None
        reply_text(token, open_id, "⏳ 正在执行监控，请稍候...")
        threading.Thread(target=run_monitor_for_user, args=(open_id, sheet_index)).start()

    elif text.startswith("添加表格"):
        handle_add_sheet(token, open_id, text)

    elif text.startswith("删除表格"):
        handle_remove_sheet(token, open_id, text)

    elif text.startswith("设置时间") or text.startswith("时间"):
        handle_set_schedule(token, open_id, text)

    elif text in ("查看定时", "定时列表", "我的定时"):
        handle_show_schedule(token, open_id)

    elif text.startswith("删除定时") or text.startswith("清除定时"):
        handle_clear_schedule(token, open_id)

    elif text in ("状态", "配置", "info", "列表"):
        handle_show_status(token, open_id)

    elif text in ("帮助", "help", "?", "？"):
        handle_help(token, open_id)

    else:
        handle_help(token, open_id)


def handle_add_sheet(token, open_id, text):
    """添加表格: 添加表格 <URL> <子表名>"""
    url_match = re.search(r'feishu\.cn/wiki/(\w+)', text)
    if not url_match:
        reply_text(token, open_id,
                  "❌ 格式:\n添加表格 <飞书wiki链接> <子表名>\n\n"
                  "例: 添加表格 https://xx.feishu.cn/wiki/xxx F022+NANO")
        return

    wiki_token = url_match.group(1)
    remaining = text[url_match.end():].strip()

    if not remaining:
        reply_text(token, open_id,
                  "❌ 请指定子表名:\n添加表格 <URL> <子表名>\n\n"
                  "例: 添加表格 https://xx.feishu.cn/wiki/xxx F022+NANO")
        return

    sheet_name = remaining
    user_cfg = get_user_config(open_id)
    sheets = user_cfg.setdefault("sheets", [])

    for s in sheets:
        if s["wiki_token"] == wiki_token and s["sheet_name"] == sheet_name:
            reply_text(token, open_id, f"⚠️ 该表格已存在: {sheet_name}")
            return

    sheets.append({
        "name": sheet_name,
        "wiki_token": wiki_token,
        "sheet_name": sheet_name
    })
    save_config()

    idx = len(sheets)
    reply_text(token, open_id,
              f"✅ 表格已添加！编号: {idx}\n"
              f"子表: {sheet_name}\n\n"
              f"发送「查询」执行全部，或「查询 {idx}」执行此表")


def handle_remove_sheet(token, open_id, text):
    """删除表格: 删除表格 <编号>"""
    num_match = re.search(r'(\d+)', text)
    if not num_match:
        reply_text(token, open_id, "❌ 格式: 删除表格 <编号>\n发送「状态」查看编号")
        return

    idx = int(num_match.group(1))
    user_cfg = get_user_config(open_id)
    sheets = user_cfg.get("sheets", [])

    if idx < 1 or idx > len(sheets):
        reply_text(token, open_id, f"❌ 编号无效，你当前共 {len(sheets)} 个表格")
        return

    removed = sheets.pop(idx - 1)
    save_config()
    reply_text(token, open_id, f"✅ 已删除: [{idx}] {removed['name']}")


def handle_set_schedule(token, open_id, text):
    """设置定时: 设置时间 09:00,14:00"""
    time_match = re.findall(r'(\d{1,2}:\d{2})', text)
    if not time_match:
        reply_text(token, open_id, "❌ 格式:\n设置时间 09:00,14:00,18:00")
        return

    user_cfg = get_user_config(open_id)
    user_cfg["schedule_times"] = time_match
    save_config()
    setup_scheduler()

    reply_text(token, open_id, f"✅ 你的定时已更新！\n每天执行: {'、'.join(time_match)}")


def handle_show_schedule(token, open_id):
    """查看定时"""
    user_cfg = get_user_config(open_id)
    times = user_cfg.get("schedule_times", [])

    if not times:
        reply_text(token, open_id, "⏰ 你当前没有设置定时任务\n\n设置方式: 设置时间 09:00,14:00")
    else:
        reply_text(token, open_id, f"⏰ 你的定时任务:\n每天 {'、'.join(times)} 自动执行\n\n修改: 设置时间 HH:MM,...\n删除: 删除定时")


def handle_clear_schedule(token, open_id):
    """删除/清除定时"""
    user_cfg = get_user_config(open_id)
    user_cfg["schedule_times"] = []
    save_config()
    setup_scheduler()

    reply_text(token, open_id, "✅ 已清除所有定时任务")


def handle_show_status(token, open_id):
    """显示用户自己的配置"""
    user_cfg = get_user_config(open_id)
    sheets = user_cfg.get("sheets", [])
    times = "、".join(user_cfg.get("schedule_times", []))

    lines = ["📋 你的配置:", "━━━━━━━━━━━━━━"]

    if not sheets:
        lines.append("表格: 未添加")
    else:
        lines.append(f"表格 ({len(sheets)} 个):")
        for i, s in enumerate(sheets, 1):
            lines.append(f"  [{i}] {s['name']}")

    lines.append(f"\n定时: {times or '未设置'}")
    lines.append("━━━━━━━━━━━━━━")
    reply_text(token, open_id, "\n".join(lines))


def handle_help(token, open_id):
    reply_text(token, open_id,
              "🖨 打印机监控机器人\n"
              "━━━━━━━━━━━━━━━━━━━━━━\n\n"
              "📌 查询:\n"
              "  查询 — 执行你的全部表格\n"
              "  查询 1 — 只执行第1个表格\n"
              "  状态 — 查看你的配置\n\n"
              "📊 表格管理:\n"
              "  添加表格 <wiki链接> <子表名>\n"
              "    例: 添加表格 https://xx.feishu.cn/wiki/xxx F022+NANO\n"
              "  删除表格 <编号>\n\n"
              "⏰ 定时:\n"
              "  设置时间 09:00,14:00,18:00\n"
              "  查看定时 — 查看当前定时\n"
              "  删除定时 — 清除所有定时\n\n"
              "📝 表格格式 (从第2行开始):\n"
              "  A:名称 B:IP C:型号 ... 邮箱列\n"
              "  邮箱列右侧: 自动写入查询结果\n\n"
              "💡 流程: 添加表格 → 查询/设置定时\n"
              "━━━━━━━━━━━━━━━━━━━━━━")

# ========== 飞书 SDK 长连接 ==========

def on_p2_im_message_receive_v1(data: P2ImMessageReceiveV1) -> None:
    try:
        event = data.event
        msg = event.message
        sender = event.sender
        if msg.message_type != "text":
            return
        open_id = sender.sender_id.open_id
        content = json.loads(msg.content)
        text = content.get("text", "").strip()
        threading.Thread(target=handle_message, args=(open_id, text)).start()
    except Exception as e:
        print(f"❌ 处理消息异常: {e}")
        import traceback
        traceback.print_exc()

# ========== 启动 ==========

def main():
    print("=" * 50)
    print("  🖨 打印机监控机器人（多用户版）")
    print("=" * 50)

    load_config()
    app_id = config['feishu_app_id']
    users = config.get("users", {})
    print(f"\n📋 配置:")
    print(f"   App ID: {app_id[:8]}...{app_id[-4:]}")
    print(f"   用户数: {len(users)}")

    setup_scheduler()
    scheduler.start()

    event_handler = lark.EventDispatcherHandler.builder("", "") \
        .register_p2_im_message_receive_v1(on_p2_im_message_receive_v1) \
        .build()

    cli = lark.ws.Client(
        config["feishu_app_id"],
        config["feishu_app_secret"],
        event_handler=event_handler,
        log_level=lark.LogLevel.INFO,
        domain=lark.FEISHU_DOMAIN
    )

    print(f"\n🔌 正在建立飞书长连接...")
    print(f"   Ctrl+C 退出\n")
    cli.start()


if __name__ == "__main__":
    main()
