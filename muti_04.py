def compare_excel_files(self):
    """对比两个Excel文件，按指定语言顺序排列结果"""
    try:
        # 获取用户选择的语言
        selected_languages = self.get_selected_languages()
        print(f"调试: 选择的语言列表: {selected_languages}")
        self.update_progress(0, f"已选择语言: {', '.join(selected_languages)}")

        self.update_progress(5, "正在选择源文件...")
        source_file = filedialog.askopenfilename(
            title="选择源Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not source_file:
            self.update_progress(0, "操作已取消")
            return
        print(f"调试: 源文件: {source_file}")

        self.update_progress(10, "正在选择翻译文件...")
        trans_file = filedialog.askopenfilename(
            title="选择翻译Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not trans_file:
            self.update_progress(0, "操作已取消")
            return
        print(f"调试: 翻译文件: {trans_file}")

        output_filename = self.get_output_filename("comparison_result")
        output_file = filedialog.asksaveasfilename(
            title="保存对比结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=output_filename
        )
        if not output_file:
            self.update_progress(0, "操作已取消")
            return
        print(f"调试: 输出文件: {output_file}")

        self.update_progress(20, "正在加载文件...")

        # 加载工作簿
        source_wb = openpyxl.load_workbook(source_file)
        trans_wb = openpyxl.load_workbook(trans_file)
        output_wb = openpyxl.Workbook()

        source_ws = source_wb.active
        trans_ws = trans_wb.active
        output_ws = output_wb.active
        output_ws.title = "对比结果"

        print(f"调试: 源文件工作表 - 行数: {source_ws.max_row}, 列数: {source_ws.max_column}")
        print(f"调试: 翻译文件工作表 - 行数: {trans_ws.max_row}, 列数: {trans_ws.max_column}")

        # 定义颜色和样式
        RED_FILL = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        GREEN_FILL = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid')
        HEADER_FONT = Font(bold=True, size=12)
        DIFF_FONT = Font(bold=True, color='FF0000')
        MATCH_FONT = Font(color='006400')
        HEADER_FILL = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')

        # 设置列宽
        output_ws.column_dimensions['A'].width = 25  # 键名
        output_ws.column_dimensions['B'].width = 20  # 语言
        output_ws.column_dimensions['C'].width = 30  # 源文件内容
        output_ws.column_dimensions['D'].width = 30  # 翻译文件内容
        output_ws.column_dimensions['E'].width = 15  # 对比结果

        # 添加表头
        headers = ["键名", "语言", "源文件内容", "翻译文件内容", "对比结果"]
        output_ws.append(headers)
        print("调试: 已添加表头")

        # 设置表头样式
        for col in range(1, 6):
            cell = output_ws.cell(row=1, column=col)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL

        # 1. 找到中文列的索引
        def find_chinese_col(ws, file_type):
            chinese_names = ["中文（CN）", "中文(CN)", "中文", "Chinese", "CN"]
            print(f"调试: 在{file_type}文件中查找中文列...")
            for col in range(1, ws.max_column + 1):
                header = str(ws.cell(row=1, column=col).value).strip()
                print(f"调试: {file_type}文件第{col}列: '{header}'")
                if any(name in header for name in chinese_names):
                    print(f"调试: 找到中文列: 第{col}列 '{header}'")
                    return col
                if re.search(r'中[文国]|汉', header):
                    print(f"调试: 通过正则找到中文列: 第{col}列 '{header}'")
                    return col
            print(f"调试: 未找到中文列，使用第1列")
            return 1

        source_cn_col = find_chinese_col(source_ws, "源")
        trans_cn_col = find_chinese_col(trans_ws, "翻译")
        print(f"调试: 源文件中文列: {source_cn_col}, 翻译文件中文列: {trans_cn_col}")

        # 2. 构建键映射
        def build_key_map(ws, cn_col, file_type):
            key_map = {}
            print(f"调试: 构建{file_type}文件键映射，使用第{cn_col}列作为键")
            for row in range(2, ws.max_row + 1):
                key = ws.cell(row=row, column=cn_col).value
                if key is not None:
                    try:
                        key_str = str(key).strip()
                        if key_str:
                            key_map[key_str] = row
                            if len(key_map) <= 5:  # 只打印前5个键用于调试
                                print(f"调试: {file_type}文件键 '{key_str}' -> 行 {row}")
                    except Exception as e:
                        print(f"调试: 处理{file_type}文件行{row}时出错: {e}")
            print(f"调试: {file_type}文件总共找到 {len(key_map)} 个键")
            return key_map

        source_key_map = build_key_map(source_ws, source_cn_col, "源")
        trans_key_map = build_key_map(trans_ws, trans_cn_col, "翻译")

        # 3. 收集所有键
        all_keys = set()
        for k in source_key_map:
            try:
                all_keys.add(str(k).strip())
            except:
                pass

        for k in trans_key_map:
            try:
                all_keys.add(str(k).strip())
            except:
                pass

        print(f"调试: 总共找到 {len(all_keys)} 个唯一键")

        # 4. 安全排序
        def safe_key_func(k):
            try:
                return str(k).lower()
            except:
                return ""

        sorted_keys = sorted(all_keys, key=safe_key_func)
        print(f"调试: 前5个排序后的键: {sorted_keys[:5]}")

        # 5. 定义语言列映射
        def build_lang_col_map(ws):
            lang_map = {}
            print(f"调试: 构建语言列映射...")
            for col in range(1, ws.max_column + 1):
                header = str(ws.cell(row=1, column=col).value).strip()
                lang_map[header] = col
                print(f"调试: 列{col}: '{header}'")
            print(f"调试: 总共找到 {len(lang_map)} 个语言列")
            return lang_map

        source_lang_map = build_lang_col_map(source_ws)
        trans_lang_map = build_lang_col_map(trans_ws)

        # 找出共同的语言
        common_langs_all = set(source_lang_map.keys()) & set(trans_lang_map.keys())
        print(f"调试: 两文件共同的语言列: {common_langs_all}")

        # 只保留用户选择的语言
        common_langs = [lang for lang in common_langs_all if lang in selected_languages]
        print(f"调试: 用户选择的共同语言: {common_langs}")

        # 按照LANGUAGE_ORDER排序
        common_langs_sorted = sorted(common_langs,
                                     key=lambda x: self.LANGUAGE_ORDER.index(
                                         x) if x in self.LANGUAGE_ORDER else len(self.LANGUAGE_ORDER))
        print(f"调试: 排序后的语言顺序: {common_langs_sorted}")

        # 6. 执行对比 - 只对比用户选择的语言
        differences_found = False
        output_row = 2  # 从第2行开始输出
        total_comparisons = 0

        print(f"调试: 开始对比 {len(sorted_keys)} 个键...")

        for i, cn_key in enumerate(sorted_keys):
            src_row = source_key_map.get(cn_key)
            trans_row = trans_key_map.get(cn_key)

            print(f"调试: 处理键 '{cn_key}' - 源文件行: {src_row}, 翻译文件行: {trans_row}")

            for lang in common_langs_sorted:
                if lang == "中文（CN）":
                    continue  # 跳过中文列，因为它是我们的键

                # 获取值
                src_val = ""
                if src_row:
                    try:
                        src_val = str(source_ws.cell(
                            row=src_row,
                            column=source_lang_map[lang]
                        ).value or "")
                    except:
                        src_val = ""

                trans_val = ""
                if trans_row:
                    try:
                        trans_val = str(trans_ws.cell(
                            row=trans_row,
                            column=trans_lang_map[lang]
                        ).value or "")
                    except:
                        trans_val = ""

                print(f"调试: 语言 '{lang}' - 源值: '{src_val}', 翻译值: '{trans_val}'")

                # 写入键名（只在该语言组的第一行写入）
                if lang == common_langs_sorted[0]:
                    output_ws.cell(row=output_row, column=1).value = cn_key
                    output_ws.cell(row=output_row, column=1).font = Font(bold=True)

                # 写入语言标签
                output_ws.cell(row=output_row, column=2).value = lang

                # 写入源内容
                output_ws.cell(row=output_row, column=3).value = src_val

                # 写入翻译内容
                output_ws.cell(row=output_row, column=4).value = trans_val

                # 对比并标记结果
                if src_val.strip() == trans_val.strip():
                    output_ws.cell(row=output_row, column=5).value = "✓ 一致"
                    output_ws.cell(row=output_row, column=5).font = MATCH_FONT
                    output_ws.cell(row=output_row, column=3).fill = GREEN_FILL
                    output_ws.cell(row=output_row, column=4).fill = GREEN_FILL
                    print(f"调试: 键 '{cn_key}' 语言 '{lang}' - 一致")
                else:
                    output_ws.cell(row=output_row, column=5).value = "✗ 差异"
                    output_ws.cell(row=output_row, column=5).font = DIFF_FONT
                    output_ws.cell(row=output_row, column=3).fill = RED_FILL
                    output_ws.cell(row=output_row, column=4).fill = RED_FILL
                    differences_found = True
                    print(f"调试: 键 '{cn_key}' 语言 '{lang}' - 发现差异")

                output_row += 1
                total_comparisons += 1

            # 添加空行分隔不同键
            output_row += 1

            # 更新进度
            progress = 40 + (i / len(sorted_keys)) * 50
            self.update_progress(progress, f"正在对比 {i + 1}/{len(sorted_keys)}...")

        print(f"调试: 总共完成了 {total_comparisons} 次对比")
        print(f"调试: 发现差异: {differences_found}")

        # 添加总结信息
        summary_row = output_row + 2
        output_ws.cell(row=summary_row, column=1).value = "对比总结"
        output_ws.cell(row=summary_row, column=1).font = HEADER_FONT

        if differences_found:
            output_ws.cell(row=summary_row + 1, column=1).value = "发现差异内容已用红色标记"
            output_ws.cell(row=summary_row + 1, column=1).font = DIFF_FONT
        else:
            output_ws.cell(row=summary_row + 1, column=1).value = "所有内容完全一致"
            output_ws.cell(row=summary_row + 1, column=1).font = MATCH_FONT

        # 添加颜色说明
        output_ws.cell(row=summary_row + 3, column=1).value = "颜色说明:"
        output_ws.cell(row=summary_row + 3, column=1).font = Font(bold=True)

        # 绿色说明
        green_cell = output_ws.cell(row=summary_row + 4, column=1)
        green_cell.value = "绿色: 内容完全一致"
        green_cell.fill = GREEN_FILL

        # 红色说明
        red_cell = output_ws.cell(row=summary_row + 5, column=1)
        red_cell.value = "红色: 内容存在差异"
        red_cell.fill = RED_FILL

        self.update_progress(95, "正在保存结果...")
        output_wb.save(output_file)

        msg = "对比完成！\n\n"
        msg += "结果文件格式说明:\n"
        msg += "- 第1列: 键名\n"
        msg += "- 第2列: 语言标签\n"
        msg += "- 第3列: 源文件内容\n"
        msg += "- 第4列: 翻译文件内容\n"
        msg += "- 第5列: 对比结果\n\n"
        msg += f"结果已保存到:\n{os.path.basename(output_file)}"
        msg += f"\n\n调试信息: 完成了 {total_comparisons} 次对比"

        self.update_progress(100, "对比完成！")
        messagebox.showinfo("完成", msg)

    except Exception as e:
        print(f"调试: 发生异常: {str(e)}")
        import traceback
        traceback.print_exc()
        self.update_progress(0, f"错误: {str(e)}")
        messagebox.showerror("错误", f"对比失败:\n{str(e)}")
    finally:
        self.root.after(2000, lambda: self.update_progress(0, "准备就绪"))