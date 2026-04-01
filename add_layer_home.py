import re
import os


def add_g28_after_each_layer(input_file, output_file):
    # 尝试多种编码读取文件
    encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1', 'cp1252']

    content = None
    used_encoding = None

    for enc in encodings:
        try:
            with open(input_file, 'r', encoding=enc) as f:
                content = f.read()
                used_encoding = enc
                print(f"成功使用 {enc} 编码读取文件")
                break
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"尝试 {enc} 编码时出错: {e}")
            continue

    if content is None:
        print("无法读取文件，请检查文件格式")
        return

    # 按行处理
    lines = content.splitlines(keepends=True)
    output_lines = []
    layer_count = 0
    inserted = False

    for i, line in enumerate(lines):
        output_lines.append(line)

        # 检测层结束标记 (TIME_ELAPSED)
        if line.startswith(';TIME_ELAPSED:'):
            # 标记可以插入位置
            inserted = False

            # 检查是否应该插入回零指令
            # 避免在文件末尾插入
            if i + 1 < len(lines):
                # 获取后续几行内容
                next_lines = []
                for j in range(1, min(5, len(lines) - i)):
                    next_lines.append(lines[i + j].strip() if lines[i + j] else '')

                # 检查是否下一层开始
                is_layer_end = False
                for nl in next_lines:
                    if 'LAYER_CHANGE' in nl:
                        is_layer_end = True
                        break

                if is_layer_end:
                    # 在 TIME_ELAPSED 后添加回零指令
                    output_lines.append('; ===== 层结束，回零 X/Y =====\n')
                    output_lines.append('G28 X Y\n')
                    layer_count += 1
                    inserted = True

    # 写入输出文件
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.writelines(output_lines)
        print(f"处理完成！共在 {layer_count} 层后添加了 G28 X Y 指令")
        print(f"输出文件: {output_file}")
    except Exception as e:
        print(f"写入文件时出错: {e}")


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("用法: python add_layer_home.py <输入文件.gcode> [输出文件.gcode]")
        print("示例: python add_layer_home.py 立方体_PLA_3h43m.gcode")
        sys.exit(1)

    input_file = sys.argv[1]

    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误: 文件 '{input_file}' 不存在")
        sys.exit(1)

    # 生成输出文件名
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    else:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_with_home{ext}"

    print(f"输入文件: {input_file}")
    print(f"输出文件: {output_file}")
    print("-" * 50)

    add_g28_after_each_layer(input_file, output_file)