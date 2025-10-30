import os


def generate_gcode(file_path, target_size_mb):
    target_bytes = target_size_mb * 1024 * 1024
    line = "G1 X10 Y10 Z0.5 F5000\n"  # 约20字节/行
    line_size = len(line.encode('utf-8'))

    with open(file_path, 'w') as f:
        written_bytes = 0
        while written_bytes < target_bytes:
            f.write(line)
            written_bytes += line_size


# 自定义路径（修改为你想要的路径）
generate_gcode("D:/test_file/1GB_test.gcode", 1024)  # 生成1GB文件
generate_gcode("D:/test_file/500MB_test.gcode", 500)  # 生成500MB文件
print("生成完成！")