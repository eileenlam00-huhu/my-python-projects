import tkinter as tk
from tkinter import filedialog, messagebox
import re
from datetime import datetime
import pandas as pd
from collections import defaultdict
import os


class PrintProcessAnalyzer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.keywords = {
            'start_print': 'Starting SD card print',
            'end_print': 'Finished SD card print',
            'cut_start': 'cmd_CR_BOX_CUT return None',
            'flush_start': 'slow_kiss_to_pipe',
            'flush_end': 'sh: restore speed factor:'
        }

    def select_file(self):
        """弹出文件选择对话框"""
        file_path = filedialog.askopenfilename(
            title="选择3D打印日志文件",
            filetypes=[("文本文件", "*.txt"), ("日志文件", "*.log"), ("所有文件", "*.*")]
        )
        return file_path

    def select_output_path(self):
        """选择输出Excel文件路径"""
        output_path = filedialog.asksaveasfilename(
            title="保存分析结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        return output_path

    def parse_timestamp(self, line):
        """从日志行中解析时间戳"""
        timestamp_patterns = [
            r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+)',
            r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})',
            r'(\d{2}:\d{2}:\d{2}\.\d+)',
            r'(\d{2}:\d{2}:\d{2})',
            r'(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})',
            r'(\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2})',
        ]

        for pattern in timestamp_patterns:
            match = re.search(pattern, line)
            if match:
                timestamp_str = match.group(1)
                try:
                    if len(timestamp_str) > 8 and ('-' in timestamp_str or '/' in timestamp_str):
                        if '-' in timestamp_str and timestamp_str.count('-') == 2:
                            if '.' in timestamp_str:
                                return datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S.%f')
                            else:
                                return datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                        elif '-' in timestamp_str and timestamp_str.count('-') == 1:
                            return datetime.strptime(timestamp_str, '%m-%d-%Y %H:%M:%S')
                        elif '/' in timestamp_str:
                            return datetime.strptime(timestamp_str, '%m/%d/%Y %H:%M:%S')
                    else:
                        today = datetime.now().strftime('%Y-%m-%d')
                        full_timestamp = f"{today} {timestamp_str}"
                        if '.' in timestamp_str:
                            return datetime.strptime(full_timestamp, '%Y-%m-%d %H:%M:%S.%f')
                        else:
                            return datetime.strptime(full_timestamp, '%Y-%m-%d %H:%M:%S')
                except ValueError as e:
                    continue
        return None

    def debug_exact_matches(self, file_path):
        """精确调试：显示包含关键字的实际行内容"""
        print(f"\n=== 文件 {os.path.basename(file_path)} 的精确调试 ===")
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                lines = file.readlines()

            keyword_counts = {key: 0 for key in self.keywords}

            for i, line in enumerate(lines):
                line = line.strip()
                for key, keyword in self.keywords.items():
                    if keyword in line:
                        keyword_counts[key] += 1
                        timestamp = self.parse_timestamp(line)
                        timestamp_info = f"时间戳: {timestamp}" if timestamp else "无时间戳"
                        if keyword_counts[key] <= 10:  # 只显示前10次匹配
                            print(f"行{i + 1} - {key}: {timestamp_info}")
                            print(f"    内容: {line}")
                            print()
                        break

            print(f"关键字匹配统计:")
            for key, count in keyword_counts.items():
                print(f"  {key}: {count} 次")
            print(f"=== 文件 {os.path.basename(file_path)} 调试结束 ===\n")

        except Exception as e:
            print(f"调试时出错: {e}")

    def analyze_print_process(self, file_path):
        """分析3D打印换色流程 - 无论是否找到开始结束都统计所有数据"""
        print("开始分析3D打印日志...")

        # 先进行精确调试
        self.debug_exact_matches(file_path)

        # 创建一个默认的会话用于存放所有数据
        all_color_changes = []
        print_sessions = []
        current_session = None
        current_color_change = None
        session_count = 0

        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                lines = file.readlines()

            for line_num, line in enumerate(lines, 1):
                line = line.strip()

                # 检查开始打印 - 如果有开始打印，创建新会话
                if self.keywords['start_print'] in line:
                    timestamp = self.parse_timestamp(line)
                    if timestamp:
                        # 保存当前会话（如果有）
                        if current_session:
                            print_sessions.append(current_session)

                        session_count += 1
                        current_session = {
                            'session_id': session_count,
                            'start_time': timestamp,
                            'end_time': None,
                            'color_changes': [],
                            'incomplete': True,
                            'has_start': True,
                            'has_end': False
                        }
                        current_color_change = None
                        print(f"✅ 创建会话 {current_session['session_id']}，开始时间: {timestamp}")

                # 检查打印结束
                elif current_session and self.keywords['end_print'] in line:
                    timestamp = self.parse_timestamp(line)
                    if timestamp:
                        current_session['end_time'] = timestamp
                        current_session['incomplete'] = False
                        current_session['has_end'] = True
                        print(f"✅ 会话 {current_session['session_id']} 完成，结束时间: {timestamp}")

                # 检查切料开始 - 无论是否有会话都记录
                if self.keywords['cut_start'] in line:
                    timestamp = self.parse_timestamp(line)
                    if timestamp:
                        # 如果没有当前会话，创建一个默认会话
                        if not current_session:
                            session_count += 1
                            current_session = {
                                'session_id': session_count,
                                'start_time': None,  # 没有开始时间
                                'end_time': None,
                                'color_changes': [],
                                'incomplete': True,
                                'has_start': False,
                                'has_end': False
                            }
                            print(f"📝 创建默认会话 {current_session['session_id']} 用于存放换色数据")

                        # 如果已经有未完成的换色流程，先完成它
                        if current_color_change:
                            print(f"找到新的切料开始，保存当前换色流程 {current_color_change['change_id']}")
                            current_session['color_changes'].append(current_color_change)
                            all_color_changes.append(current_color_change)

                        current_color_change = {
                            'change_id': len(current_session['color_changes']) + 1,
                            'cut_start_time': timestamp,
                            'flush_start_time': None,
                            'flush_end_time': None,
                            'complete': False,
                            'session_id': current_session['session_id']
                        }
                        print(f"✅ 记录换色流程 {current_color_change['change_id']}，切料开始: {timestamp}")

                # 检查冲刷开始
                elif current_color_change and self.keywords['flush_start'] in line:
                    timestamp = self.parse_timestamp(line)
                    if timestamp:
                        current_color_change['flush_start_time'] = timestamp
                        print(f"✅ 换色 {current_color_change['change_id']} 冲刷开始: {timestamp}")

                # 检查冲刷结束
                elif current_color_change and self.keywords['flush_end'] in line:
                    timestamp = self.parse_timestamp(line)
                    if timestamp:
                        current_color_change['flush_end_time'] = timestamp
                        current_color_change['complete'] = True
                        # 完成当前换色流程
                        current_session['color_changes'].append(current_color_change)
                        all_color_changes.append(current_color_change)

                        if current_color_change['cut_start_time']:
                            total_duration = (timestamp - current_color_change['cut_start_time']).total_seconds()
                            print(f"✅ 换色 {current_color_change['change_id']} 完成，总耗时: {total_duration:.2f}秒")
                        else:
                            print(f"✅ 换色 {current_color_change['change_id']} 完成，冲刷结束: {timestamp}")
                        current_color_change = None

            # 文件处理完毕，保存当前会话和最后一个换色流程
            if current_color_change:
                if current_session:
                    current_session['color_changes'].append(current_color_change)
                all_color_changes.append(current_color_change)

            if current_session:
                print_sessions.append(current_session)

            # 如果没有找到任何会话但找到了换色数据，创建一个汇总会话
            if not print_sessions and all_color_changes:
                session_count += 1
                summary_session = {
                    'session_id': session_count,
                    'start_time': None,
                    'end_time': None,
                    'color_changes': all_color_changes,
                    'incomplete': True,
                    'has_start': False,
                    'has_end': False,
                    'is_summary': True
                }
                print_sessions.append(summary_session)
                print(f"📊 创建汇总会话 {session_count} 包含所有换色数据")

            print(f"\n📊 分析完成，统计结果:")
            print(f"• 找到会话数: {len(print_sessions)}")
            print(f"• 总换色次数: {len(all_color_changes)}")

            for session in print_sessions:
                status = "未完成" if session['incomplete'] else "完成"
                has_start = "有开始" if session['has_start'] else "无开始"
                has_end = "有结束" if session['has_end'] else "无结束"
                start_time = session['start_time'].strftime('%Y-%m-%d %H:%M:%S') if session['start_time'] else "无开始时间"
                end_time = session['end_time'].strftime('%Y-%m-%d %H:%M:%S') if session['end_time'] else "无结束时间"

                print(f"会话 {session['session_id']}: {status} [{has_start}, {has_end}]")
                print(f"  开始: {start_time}")
                print(f"  结束: {end_time}")
                print(f"  换色次数: {len(session['color_changes'])}")

                for color_change in session['color_changes']:
                    if color_change['cut_start_time'] and color_change['flush_end_time']:
                        duration = (color_change['flush_end_time'] - color_change['cut_start_time']).total_seconds()
                        print(f"    换色 {color_change['change_id']}: 完成, 耗时: {duration:.2f}秒")
                    elif color_change['cut_start_time']:
                        print(
                            f"    换色 {color_change['change_id']}: 未完成, 切料时间: {color_change['cut_start_time'].strftime('%H:%M:%S')}")
                    else:
                        print(f"    换色 {color_change['change_id']}: 数据不完整")

            return print_sessions

        except Exception as e:
            print(f"分析日志时出错: {e}")
            import traceback
            traceback.print_exc()
            return []

    def calculate_statistics(self, print_sessions):
        """计算统计信息 - 包含所有数据"""
        if not print_sessions:
            return None

        # 收集所有换色数据
        all_color_changes = []
        for session in print_sessions:
            all_color_changes.extend(session['color_changes'])

        statistics = {
            'total_sessions': len(print_sessions),
            'sessions_with_start': len([s for s in print_sessions if s['has_start']]),
            'sessions_with_end': len([s for s in print_sessions if s['has_end']]),
            'total_color_changes': len(all_color_changes),
            'completed_color_changes': len([c for c in all_color_changes if c.get('complete', False)]),
            'total_print_duration': 0,
            'session_durations': [],
            'flush_durations': [],
            'color_change_durations': []
        }

        for session in print_sessions:
            # 计算会话耗时（只有有开始和结束时间的会话）
            if session['start_time'] and session['end_time']:
                session_duration = (session['end_time'] - session['start_time']).total_seconds()
                statistics['session_durations'].append(session_duration)
                statistics['total_print_duration'] += session_duration

        # 计算换色相关统计
        for color_change in all_color_changes:
            # 冲刷耗时（只计算完成的）
            if color_change['flush_start_time'] and color_change['flush_end_time']:
                flush_duration = (color_change['flush_end_time'] - color_change['flush_start_time']).total_seconds()
                statistics['flush_durations'].append(flush_duration)

            # 整个换色过程耗时（从切料开始到冲刷结束，只计算完成的）
            if (color_change['cut_start_time'] and color_change['flush_end_time']):
                color_change_duration = (
                            color_change['flush_end_time'] - color_change['cut_start_time']).total_seconds()
                statistics['color_change_durations'].append(color_change_duration)

        # 计算平均值
        if statistics['session_durations']:
            statistics['avg_session_duration'] = sum(statistics['session_durations']) / len(
                statistics['session_durations'])

        if statistics['flush_durations']:
            statistics['avg_flush_duration'] = sum(statistics['flush_durations']) / len(statistics['flush_durations'])

        if statistics['color_change_durations']:
            statistics['avg_color_change_duration'] = sum(statistics['color_change_durations']) / len(
                statistics['color_change_durations'])

        return statistics

    def generate_excel_report(self, print_sessions, statistics, output_path):
        """生成Excel报告 - 确保所有数据都写入"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

                # 1. 详细数据表
                detailed_data = []
                for session in print_sessions:
                    # 会话基本信息
                    session_duration = "N/A"
                    if session['start_time'] and session['end_time']:
                        session_duration = (session['end_time'] - session['start_time']).total_seconds()

                    session_status = "完成" if not session['incomplete'] else "未完成"
                    if not session['has_start'] and not session['has_end']:
                        session_status = "仅换色数据"
                    elif session['has_start'] and not session['has_end']:
                        session_status = "有开始无结束"

                    session_row = {
                        '打印会话ID': session['session_id'],
                        '会话状态': session_status,
                        '开始时间': session['start_time'].strftime('%Y-%m-%d %H:%M:%S') if session[
                            'start_time'] else 'N/A',
                        '结束时间': session['end_time'].strftime('%Y-%m-%d %H:%M:%S') if session['end_time'] else 'N/A',
                        '总打印耗时(秒)': session_duration,
                        '换色次数': len(session['color_changes'])
                    }
                    detailed_data.append(session_row)

                    # 换色详细数据
                    for color_change in session['color_changes']:
                        # 计算耗时
                        flush_duration = "N/A"
                        if color_change['flush_start_time'] and color_change['flush_end_time']:
                            flush_duration = (color_change['flush_end_time'] - color_change[
                                'flush_start_time']).total_seconds()

                        total_duration = "N/A"
                        if color_change['cut_start_time'] and color_change['flush_end_time']:
                            total_duration = (
                                        color_change['flush_end_time'] - color_change['cut_start_time']).total_seconds()

                        color_status = "完成" if color_change.get('complete', False) else "未完成"

                        color_row = {
                            '打印会话ID': session['session_id'],
                            '换色序号': color_change['change_id'],
                            '换色状态': color_status,
                            '切料开始时间': color_change['cut_start_time'].strftime('%Y-%m-%d %H:%M:%S') if
                            color_change['cut_start_time'] else 'N/A',
                            '冲刷开始时间': color_change['flush_start_time'].strftime('%Y-%m-%d %H:%M:%S') if
                            color_change['flush_start_time'] else 'N/A',
                            '冲刷结束时间': color_change['flush_end_time'].strftime('%Y-%m-%d %H:%M:%S') if
                            color_change['flush_end_time'] else 'N/A',
                            '冲刷耗时(秒)': flush_duration,
                            '总换色耗时(秒)': total_duration
                        }
                        detailed_data.append(color_row)

                df_detailed = pd.DataFrame(detailed_data)
                df_detailed.to_excel(writer, sheet_name='详细数据', index=False)

                # 2. 统计汇总表
                summary_data = []
                if statistics:
                    summary_data = [
                        ['统计项目', '数值'],
                        ['总会话数', statistics['total_sessions']],
                        ['有开始时间的会话', statistics['sessions_with_start']],
                        ['有结束时间的会话', statistics['sessions_with_end']],
                        ['总换色次数', statistics['total_color_changes']],
                        ['完成换色次数', statistics['completed_color_changes']],
                        ['总打印耗时(秒)', round(statistics['total_print_duration'], 2)],
                        ['平均每次打印耗时(秒)', round(statistics.get('avg_session_duration', 0), 2)],
                        ['平均每次换色耗时(秒)', round(statistics.get('avg_color_change_duration', 0), 2)],
                        ['平均冲刷耗时(秒)', round(statistics.get('avg_flush_duration', 0), 2)]
                    ]

                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='统计汇总', index=False, header=False)

            print(f"Excel报告已生成: {output_path}")
            return True

        except Exception as e:
            print(f"生成Excel报告时出错: {e}")
            import traceback
            traceback.print_exc()
            return False

    def run(self):
        """运行分析程序"""
        try:
            # 选择日志文件
            file_path = self.select_file()
            if not file_path:
                messagebox.showinfo("信息", "未选择文件，程序退出。")
                return

            print(f"选择的文件: {file_path}")

            # 分析日志
            messagebox.showinfo("信息", "开始分析日志，请稍候...")
            print_sessions = self.analyze_print_process(file_path)

            if not print_sessions:
                messagebox.showwarning("分析结果", "未找到任何打印流程数据。")
                return

            # 计算统计信息
            statistics = self.calculate_statistics(print_sessions)

            # 选择输出路径
            output_path = self.select_output_path()
            if not output_path:
                messagebox.showinfo("信息", "未选择输出路径，程序退出。")
                return

            # 生成Excel报告
            success = self.generate_excel_report(print_sessions, statistics, output_path)

            if success:
                summary = (
                    f"分析完成！\n\n"
                    f"📊 统计概要:\n"
                    f"• 会话数: {statistics['total_sessions']} 个\n"
                    f"• 有开始时间: {statistics['sessions_with_start']} 个\n"
                    f"• 有结束时间: {statistics['sessions_with_end']} 个\n"
                    f"• 总换色次数: {statistics['total_color_changes']} 次\n"
                    f"• 完成换色: {statistics['completed_color_changes']} 次\n"
                    f"• 总打印耗时: {statistics['total_print_duration']:.2f} 秒\n\n"
                    f"结果已保存到:\n{output_path}"
                )
                messagebox.showinfo("完成", summary)
            else:
                messagebox.showerror("错误", "生成Excel报告失败，请检查文件路径和权限。")

        except Exception as e:
            error_msg = f"分析过程中出现错误: {str(e)}"
            messagebox.showerror("错误", error_msg)
            print(f"错误详情: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    analyzer = PrintProcessAnalyzer()
    analyzer.run()