#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打印机状态监控 V4 - 异常自动通知负责人
- 批量查询打印机状态
- 写入飞书表格（异常标红加粗）
- 检测到异常时，按负责人分发飞书消息通知
用法: python printer_monitor_v4.py
"""

import requests
import json
import time
import re
from datetime import datetime
from collections import defaultdict

# ========== 配置区 ===========

# ---- 飞书应用凭证（从 .env 文件读取） ----
import os
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
except ImportError:
    pass

FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")

# ---- 飞书表格配置 ----
# 注意：如果表格在知识库(wiki)中，需要先通过 wiki API 获取真正的 spreadsheet_token
WIKI_NODE_TOKEN = "ETouw8WKkiiKAOkdAgEcZ0Ydnsh"  # 从 wiki URL 中获取的 node_token
SPREADSHEET_TOKEN = ""  # 留空，运行时自动从 wiki 获取
SHEET_NAME = "Sheet1"
START_ROW = 2
WRITE_COL = None  # 运行时动态定位：邮箱列右边一列

# ---- 打印机列表（含负责人配置） ----
# owner_name: 负责人姓名（用于显示）
# owner_email: 负责人飞书邮箱（用于查询 open_id 并发通知）
PRINTERS = [
    # ---- NANO (9台) ----
    {"name": "NANO-212.39", "ip": "172.23.212.39", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-214.26", "ip": "172.23.214.26", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-218.118", "ip": "172.23.218.118", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-218.33", "ip": "172.23.218.33", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-223.24", "ip": "172.23.223.24", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-222.145", "ip": "172.23.222.145", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-218.227", "ip": "172.23.218.227", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-223.19", "ip": "172.23.223.19", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "NANO-209.148", "ip": "172.23.209.148", "model": "I7 NANO", "owner_name": "林惠玲", "owner_email": ""},
    # ---- LITE (15台) ----
    {"name": "LITE-219.41", "ip": "172.23.219.41", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-220.104", "ip": "172.23.220.104", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-2.222", "ip": "172.23.2.222", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-214.114", "ip": "172.23.214.114", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-215.66", "ip": "172.23.215.66", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-218.144", "ip": "172.23.218.144", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-218.32", "ip": "172.23.218.32", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-220.165", "ip": "172.23.220.165", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-208.39", "ip": "172.23.208.39", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-214.58", "ip": "172.23.214.58", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-223.219", "ip": "172.23.223.219", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-214.214", "ip": "172.23.214.214", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-208.199", "ip": "172.23.208.199", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-208.107", "ip": "172.23.208.107", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    {"name": "LITE-214.155", "ip": "172.23.214.155", "model": "I7 LITE", "owner_name": "林惠玲", "owner_email": ""},
    # ---- MINI (1台) ----
    {"name": "MINI-223.28", "ip": "172.23.223.28", "model": "I7 MINI", "owner_name": "林惠玲", "owner_email": ""},
    # ---- I7 (1台) ----
    {"name": "I7-4.232", "ip": "172.23.4.232", "model": "I7", "owner_name": "林惠玲", "owner_email": ""},
]

PORT = 4408
TIMEOUT = 5
STATUS_API = "/printer/objects/query?print_stats&display_status&virtual_sdcard"
CONSOLE_API = "/server/gcode_store?count=100"

# 通知配置（已取消冷却期，每次运行都通知）

# ========== 飞书 API ==========

def get_feishu_token():
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}
    resp = requests.post(url, json=payload)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"获取token失败: code={data.get('code')}, msg={data.get('msg')}, "
                         f"请检查 APP_ID={FEISHU_APP_ID} 和 APP_SECRET 是否正确")
    return data["tenant_access_token"]

def get_spreadsheet_token_from_wiki(token):
    """从知识库节点获取真正的 spreadsheet_token"""
    global SPREADSHEET_TOKEN
    
    if SPREADSHEET_TOKEN:
        return SPREADSHEET_TOKEN
    
    url = f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_NODE_TOKEN}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    
    print(f"   [DEBUG] wiki get_node HTTP {resp.status_code}")
    print(f"   [DEBUG] 响应: {json.dumps(data, ensure_ascii=False)[:400]}")
    
    if data.get("code") != 0:
        raise ValueError(
            f"获取Wiki节点失败: code={data.get('code')}, msg={data.get('msg')}\n"
            f"   → 请确认: 1) 应用已开通 wiki:wiki:readonly 权限  "
            f"2) 知识库已添加应用为协作者"
        )
    
    node = data["data"]["node"]
    obj_type = node.get("obj_type")
    obj_token = node.get("obj_token")
    
    print(f"   [INFO] Wiki节点类型: {obj_type}, obj_token: {obj_token}")
    
    if obj_type != "sheet":
        raise ValueError(
            f"该Wiki节点不是电子表格类型，而是 '{obj_type}'。"
            f"请确认链接指向的是一个电子表格。"
        )
    
    SPREADSHEET_TOKEN = obj_token
    return obj_token

def get_user_open_id(token, printer_info):
    """获取用户 open_id，通过邮箱查询"""
    email = printer_info.get("owner_email", "")
    if not email:
        print(f"   ⚠️ 未配置 owner_email，无法发送通知")
        return None
    
    url = "https://open.feishu.cn/open-apis/contact/v3/users/batch_get_id?user_id_type=open_id"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"emails": [email]}
    resp = requests.post(url, headers=headers, json=payload)
    data = resp.json()
    if data.get("code") != 0:
        print(f"   ⚠️ 邮箱查询失败: code={data.get('code')}, msg={data.get('msg')}")
        return None
    user_list = data.get("data", {}).get("user_list", [])
    if user_list and user_list[0].get("user_id"):
        return user_list[0]["user_id"]
    print(f"   ⚠️ 邮箱 {email} 未匹配到用户")
    return None


def send_feishu_message(token, user_open_id, content):
    """给用户发飞书消息（富文本卡片）"""
    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "receive_id": user_open_id,
        "msg_type": "interactive",
        "content": json.dumps(content, ensure_ascii=False)
    }
    resp = requests.post(url, headers=headers, json=payload)
    result = resp.json()
    if result.get("code") != 0:
        print(f"   发消息API返回: code={result.get('code')}, msg={result.get('msg')}")
    return result

def build_error_card(owner_email, error_printers):
    """构建异常通知卡片"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    elements = []
    for p in error_printers:
        status = p["status"]
        error_lines = []
        
        # 设备名 + 状态
        error_lines.append(f"**【{status['状态']}】{p['name']}**（{p['model']}）")
        error_lines.append(f"IP: {p['ip']}")
        
        if status["文件名"] != "-":
            error_lines.append(f"文件: {status['文件名'][:40]}")
        
        if status["进度"] != "-":
            error_lines.append(f"进度: {status['进度']}（已打印 {status['已打印']}）")
        
        if status["错误码"]:
            error_lines.append(f"🚨 错误码: **{status['错误码']}**")
        
        if status["错误详情"]:
            detail = status["错误详情"]
            if len(detail) > 60:
                detail = detail[:57] + "..."
            error_lines.append(f"异常: {detail}")
        
        elements.append({
            "tag": "div",
            "text": {
                "tag": "lark_md",
                "content": "\n".join(error_lines)
            }
        })
        elements.append({"tag": "hr"})
    
    # 去掉最后一个分割线
    if elements and elements[-1].get("tag") == "hr":
        elements.pop()
    
    # 查看详情按钮
    elements.append({
        "tag": "action",
        "actions": [{
            "tag": "button",
            "text": {
                "tag": "plain_text",
                "content": "查看状态跟踪表"
            },
            "type": "danger",
            "url": f"https://gcmomk3i2c.feishu.cn/wiki/{WIKI_NODE_TOKEN}"
        }]
    })
    
    count = len(error_printers)
    card = {
        "config": {"wide_screen_mode": True},
        "header": {
            "title": {
                "tag": "plain_text",
                "content": f"🚨 打印机异常通知（{count}台）"
            },
            "template": "red"
        },
        "elements": elements
    }
    return card

# ========== 飞书表格写入 ==========

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


def get_sheet_id(token):
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/query"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    
    # 调试：打印完整响应（排查后可删）
    print(f"   [DEBUG] sheets query HTTP {resp.status_code}")
    print(f"   [DEBUG] 响应: {json.dumps(data, ensure_ascii=False)[:300]}")
    
    if data.get("code") != 0:
        raise ValueError(
            f"获取子表失败: code={data.get('code')}, msg={data.get('msg')}\n"
            f"   → 请确认: 1) 应用已开通 sheets:spreadsheet 权限  "
            f"2) 表格已添加应用为协作者  "
            f"3) SPREADSHEET_TOKEN={SPREADSHEET_TOKEN} 是否正确"
        )
    
    if "data" not in data or "sheets" not in data.get("data", {}):
        raise ValueError(f"API返回格式异常，完整响应: {json.dumps(data, ensure_ascii=False)[:500]}")
    
    sheets = data["data"]["sheets"]
    sheet_titles = [s["title"] for s in sheets]
    for s in sheets:
        if s["title"] == SHEET_NAME:
            return s["sheet_id"]
    raise ValueError(f"找不到子表 '{SHEET_NAME}'，现有子表: {sheet_titles}")


def find_email_col(token, sheet_id):
    """读取第1行表头，找到'邮箱'列的位置（1-based），返回邮箱列+1作为写入列"""
    range_str = f"{sheet_id}!A1:Z1"
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{range_str}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers)
    data = resp.json()
    if data.get("code") != 0:
        raise ValueError(f"读取表头失败: {data.get('msg')}")

    rows = data.get("data", {}).get("valueRange", {}).get("values", [])
    if not rows:
        raise ValueError("表头行为空，无法定位邮箱列")

    header_row = rows[0]
    for i, cell in enumerate(header_row):
        cell_text = extract_cell_text(cell)
        if "邮箱" in cell_text:
            email_col = i + 1  # 转为1-based
            write_col = email_col + 1
            print(f"   📍 检测到邮箱列: {col_letter(email_col)}列(第{email_col}列)，写入列: {col_letter(write_col)}列(第{write_col}列)")
            return write_col

    raise ValueError(f"表头中未找到包含'邮箱'的列，表头内容: {[extract_cell_text(c) for c in header_row]}")

def insert_column(token, sheet_id, col_index):
    """在指定位置插入一列（把现有数据往右推）
    
    col_index: 从第几列前插入（1-based），API用0-based
    """
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/insert_dimension_range"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    payload = {
        "dimension": {
            "sheetId": sheet_id,
            "majorDimension": "COLUMNS",
            "startIndex": col_index - 1,  # 0-based，E列=4
            "endIndex": col_index          # 插入1列
        },
        "inheritStyle": "AFTER"
    }
    
    resp = requests.post(url, headers=headers, json=payload)
    result = resp.json()
    if result.get("code") != 0:
        print(f"   ⚠️ 插入列失败: {result.get('msg')}")
    return result


def write_rich_cells(token, sheet_id, rows_data):
    """批量写入单元格（纯文本方式，兼容 values_batch_update）"""
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_update"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    value_ranges = []
    for row_data in rows_data:
        col = col_letter(row_data['col'])
        row = row_data['row']
        range_str = f"{sheet_id}!{col}{row}:{col}{row}"
        # 将富文本元素拼接为纯文本
        text_content = build_plain_text(row_data["elements"])
        value_ranges.append({
            "range": range_str,
            "values": [[text_content]]
        })
    
    payload = {"valueRanges": value_ranges}
    resp = requests.post(url, headers=headers, json=payload)
    return resp.json()


def build_plain_text(elements):
    """将富文本元素列表拼成纯文本字符串"""
    lines = []
    current_line = ""
    for el in elements:
        if el.get("newline"):
            lines.append(current_line)
            current_line = ""
        else:
            current_line += el.get("text", "")
    if current_line:
        lines.append(current_line)
    return "\n".join(lines)

def col_letter(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def build_richtext_json(elements):
    lines = []
    current_line = []
    
    for el in elements:
        if el.get("newline"):
            if current_line:
                lines.append(current_line)
                current_line = []
            continue
        
        text_el = {"text_run": {"content": el["text"]}}
        style = {}
        if el.get("bold"):
            style["bold"] = True
        if el.get("color"):
            style["fore_color"] = {
                "alpha": 1,
                "red": int(el["color"][1:3], 16) / 255,
                "green": int(el["color"][3:5], 16) / 255,
                "blue": int(el["color"][5:7], 16) / 255
            }
        if style:
            text_el["text_run"]["text_element_style"] = style
        
        current_line.append(text_el)
    
    if current_line:
        lines.append(current_line)
    
    return json.dumps({"lines": [{"elements": line} for line in lines]}, ensure_ascii=False)

def build_summary_richtext(printer_info, status):
    is_error = status["状态"] in ("异常", "离线")
    error_color = "#F53F3F"
    normal_color = "#1D2129"
    
    elements = []
    elements.append({
        "text": f"【{status['状态']}】{printer_info['name']}",
        "bold": is_error,
        "color": error_color if is_error else normal_color
    })
    elements.append({"newline": True})
    elements.append({"text": f"型号: {printer_info['model']}  |  IP: {printer_info['ip']}"})
    elements.append({"newline": True})
    
    if status["文件名"] != "-":
        elements.append({"text": f"文件: {status['文件名'][:50]}"})
        elements.append({"newline": True})
    
    if status["进度"] != "-":
        elements.append({"text": f"进度: {status['进度']}  [{status['已打印']} / {status['总预估']}]"})
        elements.append({"newline": True})
    
    if status["预计完成"] != "-":
        elements.append({"text": f"预计完成: {status['预计完成']}"})
        elements.append({"newline": True})
    
    if status["总打印时长"] != "-":
        elements.append({"text": f"总打印时长: {status['总打印时长']}"})
        elements.append({"newline": True})
    
    if status["错误码"]:
        elements.append({"newline": True})
        elements.append({
            "text": f"⚠ 错误码: {status['错误码']}",
            "bold": True,
            "color": error_color
        })
        elements.append({"newline": True})
    
    if status["错误详情"]:
        detail = status["错误详情"]
        if len(detail) > 80:
            detail = detail[:77] + "..."
        elements.append({"text": f"异常: {detail}", "color": error_color})
    
    return elements

# ========== 打印机查询 ==========

def query_printer(ip):
    result = {"状态": "离线", "文件名": "-", "进度": "-", 
              "已打印": "-", "总预估": "-", "预计完成": "-", 
              "错误码": "", "错误详情": "", "总打印时长": "-"}
    
    try:
        url = f"http://{ip}:{PORT}{STATUS_API}"
        resp = requests.get(url, timeout=TIMEOUT)
        resp.raise_for_status()
        status_info = parse_status(resp.json())
        result.update(status_info)
    except Exception as e:
        result["状态"] = "离线"
        result["错误详情"] = str(e)
        return result
    
    try:
        url = f"http://{ip}:{PORT}{CONSOLE_API}"
        resp = requests.get(url, timeout=TIMEOUT)
        resp.raise_for_status()
        errors = parse_console_errors(resp.json())
        if errors:
            result["错误码"] = "、".join(errors["codes"])
            result["错误详情"] = errors["detail"]
    except Exception:
        pass
    
    return result

def parse_status(data):
    r = {}
    ps = data.get("result", {}).get("status", {}).get("print_stats", {})
    ds = data.get("result", {}).get("status", {}).get("display_status", {})
    
    state = ps.get("state", "unknown")
    state_map = {
        "printing": "打印中", "paused": "已暂停", "complete": "已完成",
        "error": "异常", "standby": "待机", "ready": "就绪",
        "cancelled": "已取消", "startup": "启动中"
    }
    r["状态"] = state_map.get(state, state)
    r["文件名"] = ps.get("filename", "-") or "-"
    progress = ds.get("progress", 0)
    r["进度"] = f"{progress * 100:.1f}%"
    dur = ps.get("print_duration", 0)
    r["已打印"] = fmt_time(dur)
    
    if progress > 0.001 and dur > 0:
        total = dur / progress
        r["总预估"] = fmt_time(total)
        finish = time.time() + (total - dur)
        r["预计完成"] = datetime.fromtimestamp(finish).strftime("%m-%d %H:%M")
    else:
        r["总预估"] = "-"
        r["预计完成"] = "-"
    
    msg = ps.get("message", "")
    if state == "error" and msg:
        r["错误详情"] = msg
    
    # 总打印时长
    total_duration = ps.get("total_duration", 0)
    r["总打印时长"] = fmt_time(total_duration) if total_duration > 0 else "-"
    
    return r

def parse_console_errors(data):
    gcode_store = data.get("result", {}).get("gcode_store", [])
    errors = []
    error_codes = set()
    
    pattern = re.compile(r'!!\s*(\{["\']code["\']\s*:\s*["\']key[^"\'}]+["\'].*?\})', re.IGNORECASE)
    
    for entry in gcode_store:
        if isinstance(entry, list):
            msg = entry[1] if len(entry) > 1 else ""
        elif isinstance(entry, dict):
            msg = entry.get("message", "")
        else:
            msg = str(entry)
        
        matches = pattern.findall(msg)
        for match in matches:
            try:
                err_obj = json.loads(match)
                code = err_obj.get("code", "")
                err_msg = err_obj.get("msg", "")
                if code:
                    error_codes.add(code)
                    errors.append(f"{code}: {err_msg}" if err_msg else code)
            except json.JSONDecodeError:
                errors.append(match)
    
    if error_codes:
        return {
            "codes": sorted(list(error_codes)),
            "detail": " | ".join(errors[-5:])
        }
    return None

def fmt_time(sec):
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    if h > 0:
        return f"{h}h{m}m"
    elif m > 0:
        return f"{m}m{s}s"
    return f"{s}s"

# ========== 主函数 ==========

def main():
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n{'='*70}")
    print(f"  打印机状态监控 V4 - {now}")
    print(f"{'='*70}\n")
    
    # 1. 查询所有打印机
    results = []
    for i, p in enumerate(PRINTERS):
        status = query_printer(p["ip"])
        results.append({"printer": p, "status": status})
        
        flag = "🔴" if status["状态"] in ("异常", "离线") else "🟢" if status["状态"] == "打印中" else "🟡"
        err_tag = f"  [{status['错误码']}]" if status["错误码"] else ""
        total_tag = f"  累计:{status['总打印时长']}" if status["总打印时长"] != "-" else ""
        print(f"{flag} {p['name']:15s} {status['状态']:6s} "
              f"{status['进度']:7s} {status['已打印']:>6s}/{status['总预估']:<6s}{total_tag}{err_tag}")
    
    # 统计
    printing = sum(1 for r in results if r["status"]["状态"] == "打印中")
    error = sum(1 for r in results if r["status"]["状态"] in ("异常", "离线"))
    done = sum(1 for r in results if r["status"]["状态"] == "已完成")
    has_error_code = sum(1 for r in results if r["status"]["错误码"])
    print(f"\n📊 总计: {len(results)}台 | "
          f"🟢 打印中: {printing} | "
          f"✅ 已完成: {done} | "
          f"🔴 异常/离线: {error} | "
          f"⚠️  有错误码: {has_error_code}")
    
    # 2. 写入飞书表格
    print(f"\n📝 正在写入飞书表格...")
    try:
        token = get_feishu_token()
        
        # 从 Wiki 节点获取真正的 spreadsheet_token
        spreadsheet_token = get_spreadsheet_token_from_wiki(token)
        print(f"   ✅ 获取到 spreadsheet_token: {spreadsheet_token}")
        
        sheet_id = get_sheet_id(token)
        
        # 动态定位邮箱列，确定写入列位置
        write_col = find_email_col(token, sheet_id)
        
        # 在邮箱列右边插入新列（旧数据往右推，确保最新数据始终紧邻邮箱列）
        print(f"   📊 在 {col_letter(write_col)} 列前插入新列...")
        insert_result = insert_column(token, sheet_id, write_col)
        if insert_result.get("code") == 0:
            print(f"   ✅ 新列已插入")
        else:
            print(f"   ⚠️ 插入列返回: {insert_result}")
        
        # 写入表头（查询时间）
        query_time = datetime.now().strftime("%m-%d %H:%M")
        header_data = [{
            "row": 1,
            "col": write_col,
            "elements": [{"text": query_time}]
        }]
        write_rich_cells(token, sheet_id, header_data)
        
        # 写入各打印机状态数据
        rows_data = []
        for i, r in enumerate(results):
            row_num = START_ROW + i
            richtext = build_summary_richtext(r["printer"], r["status"])
            rows_data.append({
                "row": row_num,
                "col": write_col,
                "elements": richtext
            })
        
        result = write_rich_cells(token, sheet_id, rows_data)
        if result.get("code") == 0:
            print(f"✅ 写入成功！共 {len(rows_data)} 行，时间: {query_time}")
        else:
            print(f"❌ 写入失败: {result}")
    except Exception as e:
        print(f"❌ 写入飞书表格失败: {e}")
        import traceback
        traceback.print_exc()
    
    # 3. 异常通知（通知暂停/异常/离线/有错误码的机器）
    print(f"\n📨 正在发送异常通知...")
    try:
        token = get_feishu_token()
        
        # 按负责人分组异常机器
        error_by_owner = defaultdict(list)
        owner_info_map = {}  # owner_name -> printer_info（用于查 open_id）
        for r in results:
            status = r["status"]
            if status["状态"] in ("已暂停", "异常", "离线") or status["错误码"]:
                owner_name = r["printer"].get("owner_name", "")
                if owner_name:
                    error_by_owner[owner_name].append({
                        "name": r["printer"]["name"],
                        "ip": r["printer"]["ip"],
                        "model": r["printer"]["model"],
                        "status": status
                    })
                    # 记录该负责人的配置信息（用于查 open_id）
                    if owner_name not in owner_info_map:
                        owner_info_map[owner_name] = r["printer"]
        
        if not error_by_owner:
            print("ℹ️  没有异常状态的机器，无需通知")
        else:
            total_notified = 0
            for owner_name, printers in error_by_owner.items():
                # 获取 open_id
                print(f"   🔍 获取用户 {owner_name} 的 open_id...")
                open_id = get_user_open_id(token, owner_info_map[owner_name])
                if not open_id:
                    print(f"   ⚠️  无法获取 {owner_name} 的 open_id，跳过通知")
                    continue
                
                print(f"   ✅ open_id: {open_id}")
                
                # 发卡片消息
                card = build_error_card(owner_name, printers)
                result = send_feishu_message(token, open_id, card)
                
                if result.get("code") == 0:
                    print(f"   ✅ 已通知 {owner_name}（{len(printers)}台异常）")
                    total_notified += 1
                else:
                    print(f"   ❌ 通知 {owner_name} 失败: code={result.get('code')}, msg={result.get('msg')}")
                    if result.get("code") == 230001:
                        print(f"      → 用户未与机器人对话过，需先在飞书搜索并打开机器人对话")
            
            print(f"\n📊 共通知 {total_notified} 位负责人")
    except Exception as e:
        print(f"❌ 发送通知失败: {e}")
        import traceback
        traceback.print_exc()
    
    print()

if __name__ == "__main__":
    main()
