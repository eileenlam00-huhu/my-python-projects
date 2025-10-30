import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from pathlib import Path


def select_input_file():
    """选择输入Excel文件"""
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="选择原始Excel文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
    )


def select_output_file():
    """选择输出Excel文件"""
    root = tk.Tk()
    root.withdraw()
    return filedialog.asksaveasfilename(
        title="保存结果Excel文件",
        defaultextension=".xlsx",
        filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
    )


def extract_matrices_by_sheet(input_path, output_path):
    """按原始Sheet提取矩阵到对应Sheet"""
    try:
        # 读取原始Excel的所有Sheet名
        xls = pd.ExcelFile(input_path)
        sheet_names = xls.sheet_names

        # 创建新Excel文件
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet in sheet_names:
                # 读取每个Sheet的数据
                df = pd.read_excel(input_path, sheet_name=sheet, header=None)
                content = "\n".join(["\t".join(row.astype(str)) for _, row in df.iterrows()])

                # 提取Before矩阵
                pattern = r"Before Probed_matrix:(?:\s+\S+){9}([\s\S]+?)(?=\n\s*Now Probed_matrix:|\n\*{5,})"
                matrices = re.findall(pattern, content)

                if matrices:
                    # 只取第一个匹配的矩阵(假设每个Sheet只有一个Before矩阵)
                    rows = [line.split() for line in matrices[0].strip().split('\n') if line.strip()]

                    if len(rows) == 9 and all(len(row) >= 9 for row in rows):
                        # 写入同名Sheet
                        pd.DataFrame(rows[:9]).to_excel(
                            writer,
                            sheet_name=sheet,
                            index=False,
                            header=False
                        )

        messagebox.showinfo("完成", f"矩阵已提取到:\n{output_path}")
    except Exception as e:
        messagebox.showerror("错误", f"处理失败:\n{str(e)}")


def main():
    input_file = select_input_file()
    if not input_file:
        return

    output_file = select_output_file()
    if output_file:
        extract_matrices_by_sheet(input_file, output_file)


if __name__ == "__main__":
    main()