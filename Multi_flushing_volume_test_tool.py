"""
Creality Print自动颜色冲刷测试塔生成器
版本: 18.0
功能: 生成多颜色标记测试塔 + 自动换料脚本
特性: 颜色单选/多选 + 自定义冲刷量
作者: AI助手
日期: 2024年
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json
import os
import math
import struct
import sys
from datetime import datetime

# ============================================================================
# 默认颜色配置
# ============================================================================

DEFAULT_COLORS = {
    "黑色": {
        "name": "黑色",
        "recommended": [50, 100, 150, 200, 250],
        "description": "深色材料，需要较多冲刷量",
        "color_code": "#000000"
    },
    "红色": {
        "name": "红色",
        "recommended": [30, 60, 90, 120, 150],
        "description": "中等深色，冲刷量适中",
        "color_code": "#FF0000"
    },
    "蓝色": {
        "name": "蓝色",
        "recommended": [30, 60, 90, 120, 150],
        "description": "中等颜色，冲刷量适中",
        "color_code": "#0000FF"
    },
    "黄色": {
        "name": "黄色",
        "recommended": [20, 40, 60, 80, 100],
        "description": "浅色材料，需要较少冲刷量",
        "color_code": "#FFFF00"
    },
    "绿色": {
        "name": "绿色",
        "recommended": [30, 60, 90, 120, 150],
        "description": "中等颜色，冲刷量适中",
        "color_code": "#00FF00"
    },
    "白色": {
        "name": "白色",
        "recommended": [10, 20, 30, 40, 50],
        "description": "最浅色，需要最少冲刷量",
        "color_code": "#FFFFFF"
    },
    "紫色": {
        "name": "紫色",
        "recommended": [40, 80, 120, 160, 200],
        "description": "深色材料，需要较多冲刷量",
        "color_code": "#800080"
    },
    "橙色": {
        "name": "橙色",
        "recommended": [25, 50, 75, 100, 125],
        "description": "浅色材料，冲刷量较少",
        "color_code": "#FFA500"
    }
}


# ============================================================================
# 配置文件生成器
# ============================================================================

class AutoConfigGenerator:
    """自动配置生成器"""

    @staticmethod
    def create_auto_guide(config):
        """创建自动化操作指南"""

        guide = f"""🎯 Creality Print自动颜色测试塔使用指南
===========================================
生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
模型文件：{config['filename']}.stl

✨ 测试配置：
• 测试颜色：{', '.join([c['name'].split('→')[0] for c in config['colors']])}
• 总测试点：{sum(len(c['volumes']) for c in config['colors'])} 个
• 测试塔高度：{config['base_height'] + sum(len(c['volumes']) for c in config['colors']) * config['layer_height']:.1f} mm

===========================================
第1步：模型特点
===========================================
测试塔具有以下特点：
"""

        # 描述颜色段
        current_z = config["base_height"]
        for color_idx, color_test in enumerate(config["colors"]):
            color_name = color_test["name"].split("→")[0].strip()
            num_tests = len(color_test["volumes"])
            segment_height = num_tests * config["layer_height"]

            guide += f"""
【{color_name}测试段】
• 高度范围：{current_z:.1f} - {current_z + segment_height:.1f} mm
• 测试冲刷量：{', '.join(map(str, color_test['volumes']))} mm³
• 标记方式：凹槽数量表示冲刷量大小（1-9个凹槽）
"""

            current_z += segment_height + config["separator_height"]

        guide += f"""
===========================================
第2步：全自动处理流程
===========================================
📁 生成的文件：
{config['filename']}.stl        # 测试塔模型
{config['filename']}_auto.py    # 自动处理脚本
{config['filename']}_guide.txt  # 本指南

🚀 自动化流程：
1. 在Creality Print中导入 {config['filename']}.stl
2. 正常切片，保存为G-code文件
3. 运行自动脚本处理G-code
4. 使用处理后的G-code打印

===========================================
第3步：详细操作步骤
===========================================
请在以下高度进行颜色切换：
"""

        # 生成切换点表格
        current_z = config["base_height"]
        test_num = 1

        for color_idx, color_test in enumerate(config["colors"]):
            color_name = color_test["name"].split("→")[0].strip()

            guide += f"\n【{color_name}测试】\n"
            guide += "序号 | 高度(mm) | 操作 | 冲刷量 | 标记\n"
            guide += "-----|----------|------|--------|------\n"

            for vol_idx, volume in enumerate(color_test["volumes"]):
                layer_z = current_z + vol_idx * config["layer_height"]
                dot_count = min(9, max(1, volume // 30))

                guide += f"{test_num:4} | {layer_z:8.2f} | {color_name}→白色 | {volume:6} | {dot_count}个凹槽\n"
                test_num += 1

            current_z += len(color_test["volumes"]) * config["layer_height"]

            if color_idx < len(config["colors"]) - 1:
                guide += f"     | {current_z:8.2f} | 分隔层 | -      | ---\n"
                current_z += config["separator_height"]

        guide += f"""
===========================================
第4步：结果记录表
===========================================
请打印完成后填写：

测试日期：__________________
打印机型号：________________
环境温度：__________________

【测试结果记录】
序号 | 颜色 | 冲刷量(mm³) | 效果评分 | 最佳值 | 备注
-----|------|-------------|----------|--------|------
"""

        # 结果表格
        test_num = 1
        for color_test in config["colors"]:
            color_name = color_test["name"].split("→")[0].strip()
            for volume in color_test["volumes"]:
                guide += f"{test_num:4} | {color_name:4} | {volume:11} |        |        |      \n"
                test_num += 1

        guide += f"""
效果评分标准：
★★★★★ 完美：完全无污染，白色纯净
★★★★☆ 优秀：极轻微污染，几乎看不见
★★★☆☆ 良好：轻微污染，可接受范围
★★☆☆☆ 一般：明显污染，需要增加冲刷量
★☆☆☆☆ 差：严重污染，无法接受

===========================================
第5步：智能推荐
===========================================
根据您的测试结果，推荐以下冲刷量：

【基础冲刷量】
• 找到第一个达到★★★★☆效果的冲刷量
• 将其作为基础值

【安全冲刷量】
• 安全值 = 基础值 × 1.5
• 日常打印使用安全值

【快速切换】
• 浅色→浅色：基础值 × 0.8
• 深色→浅色：基础值 × 2.0
• 深色→深色：基础值 × 1.2

===========================================
第6步：常见问题
===========================================
❓ M600命令不工作？
✅ 解决方法：
1. 检查打印机固件是否支持M600
2. 使用手动暂停功能
3. 在G-code中搜索Z高度，手动暂停

❓ 颜色切换后挤出不足？
✅ 解决方法：
1. 提高喷嘴温度5-10°C
2. 在暂停后手动挤出5-10mm
3. 确保材料干燥

❓ 标记看不清楚？
✅ 解决方法：
1. 降低打印速度至30mm/s
2. 增加标记尺寸（修改配置）
3. 使用放大镜观察

===========================================
📞 技术支持
===========================================
请记录以下信息以便诊断：
1. 打印机型号和固件版本
2. 使用的材料品牌和类型
3. 打印温度设置
4. 观察到的具体问题
5. 拍摄清晰的照片

祝您测试顺利！ 🎉
"""

        return guide


# ============================================================================
# 智能STL生成器
# ============================================================================

class SmartSTLGenerator:
    """智能STL生成器 - 带颜色标记"""

    @staticmethod
    def generate_color_tower(config, filename):
        """生成彩色测试塔"""
        try:
            vertices = []
            faces = []
            face_count = 0

            print(f"🔄 正在生成彩色测试塔...")

            # 1. 创建白色底座
            base_vertices, base_faces = SmartSTLGenerator._create_cylinder(
                0, 0, config["base_height"] / 2,
                      config["width"] / 2, config["base_height"],
                segments=32
            )
            vertices.extend(base_vertices)
            for face in base_faces:
                faces.append([face[0], face[1], face[2]])
                face_count += 1

            # 2. 创建测试层
            current_z = config["base_height"]

            for color_idx, color_test in enumerate(config["colors"]):
                color_name = color_test["name"].split("→")[0].strip()
                num_layers = len(color_test["volumes"])

                print(f"  生成 {color_name} 测试段...")

                # 为每个冲刷量创建测试层
                for layer_idx, volume in enumerate(color_test["volumes"]):
                    layer_z = current_z + config["layer_height"] / 2

                    # 主体圆柱
                    body_radius = config["width"] * 0.35
                    body_vertices, body_faces = SmartSTLGenerator._create_cylinder(
                        0, 0, layer_z,
                        body_radius, config["layer_height"],
                        segments=32
                    )

                    base_idx = len(vertices)
                    vertices.extend(body_vertices)
                    for face in body_faces:
                        faces.append([face[0] + base_idx, face[1] + base_idx, face[2] + base_idx])
                        face_count += 1

                    # 添加冲刷量标记（凹槽）
                    dot_count = min(9, max(1, volume // 30))
                    mark_radius = body_radius * 1.1

                    for dot_idx in range(dot_count):
                        angle = (dot_idx * 2 * math.pi) / dot_count
                        mark_x = mark_radius * math.cos(angle)
                        mark_y = mark_radius * math.sin(angle)
                        mark_z = layer_z

                        # 创建凹槽标记
                        mark_vertices, mark_faces = SmartSTLGenerator._create_marker(
                            mark_x, mark_y, mark_z,
                            config["layer_height"] * 0.8,
                            dot_count, dot_idx
                        )

                        base_idx = len(vertices)
                        vertices.extend(mark_vertices)
                        for face in mark_faces:
                            faces.append([face[0] + base_idx, face[1] + base_idx, face[2] + base_idx])
                            face_count += 1

                    current_z += config["layer_height"]

                # 添加颜色段分隔层
                if color_idx < len(config["colors"]) - 1:
                    separator_z = current_z + config["separator_height"] / 2
                    sep_vertices, sep_faces = SmartSTLGenerator._create_separator(
                        0, 0, separator_z,
                        config["width"] * 0.4,
                        config["separator_height"]
                    )

                    base_idx = len(vertices)
                    vertices.extend(sep_vertices)
                    for face in sep_faces:
                        faces.append([face[0] + base_idx, face[1] + base_idx, face[2] + base_idx])
                        face_count += 1

                    current_z += config["separator_height"]

            # 3. 保存STL文件
            with open(filename, 'wb') as f:
                # STL头部
                header = f"ColorTestTower v3.0 {config['filename']}".ljust(80, ' ')
                f.write(header.encode('ascii'))

                # 面片数量
                f.write(struct.pack('<I', face_count))

                # 写入每个面片
                for face in faces:
                    v1 = vertices[face[0]]
                    v2 = vertices[face[1]]
                    v3 = vertices[face[2]]

                    # 计算法向量
                    normal = SmartSTLGenerator._calculate_normal(v1, v2, v3)

                    # 写入数据
                    f.write(struct.pack('<3f', *normal))
                    f.write(struct.pack('<3f', *v1))
                    f.write(struct.pack('<3f', *v2))
                    f.write(struct.pack('<3f', *v3))
                    f.write(struct.pack('<H', 0))

            print(f"✅ STL文件已生成: {filename}")
            print(f"   面片数量: {face_count}")
            print(f"   测试塔高度: {current_z:.1f}mm")

            return True

        except Exception as e:
            print(f"❌ 生成STL失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    @staticmethod
    def _create_cylinder(x, y, z, radius, height, segments=32):
        """创建圆柱体"""
        vertices = []
        faces = []

        # 底部圆心
        vertices.append((x, y, z - height / 2))

        # 顶部圆心
        vertices.append((x, y, z + height / 2))

        # 侧面顶点
        for i in range(segments):
            angle = 2 * math.pi * i / segments
            vx = x + radius * math.cos(angle)
            vy = y + radius * math.sin(angle)

            # 底部顶点
            vertices.append((vx, vy, z - height / 2))
            # 顶部顶点
            vertices.append((vx, vy, z + height / 2))

        # 侧面三角形
        for i in range(segments):
            next_i = (i + 1) % segments

            # 底部索引
            b1 = 2 + i * 2
            b2 = 2 + next_i * 2

            # 顶部索引
            t1 = b1 + 1
            t2 = b2 + 1

            # 侧面四边形（两个三角形）
            faces.append((b1, t1, t2))
            faces.append((b1, t2, b2))

        # 底部面
        for i in range(segments):
            next_i = (i + 1) % segments
            faces.append((0, 2 + next_i * 2, 2 + i * 2))

        # 顶部面
        for i in range(segments):
            next_i = (i + 1) % segments
            faces.append((1, 2 + i * 2 + 1, 2 + next_i * 2 + 1))

        return vertices, faces

    @staticmethod
    def _create_marker(x, y, z, height, total_dots, dot_index):
        """创建标记点（凹槽）"""
        vertices = []
        faces = []

        radius = height * 0.2
        segments = 8

        # 创建圆柱形凹槽
        base_idx = len(vertices)

        # 底部圆心
        vertices.append((x, y, z - height / 2))

        # 顶部圆心
        vertices.append((x, y, z + height / 2))

        # 侧面顶点
        for i in range(segments):
            angle = 2 * math.pi * i / segments
            vx = x + radius * math.cos(angle)
            vy = y + radius * math.sin(angle)

            vertices.append((vx, vy, z - height / 2))
            vertices.append((vx, vy, z + height / 2))

        # 侧面（向内，形成凹槽）
        for i in range(segments):
            next_i = (i + 1) % segments

            b1 = base_idx + 2 + i * 2
            b2 = base_idx + 2 + next_i * 2
            t1 = b1 + 1
            t2 = b2 + 1

            # 注意：法向量方向反转
            faces.append((b1, b2, t2))
            faces.append((b1, t2, t1))

        # 底部（凹槽底部）
        for i in range(segments):
            next_i = (i + 1) % segments
            faces.append((base_idx, base_idx + 2 + i * 2, base_idx + 2 + next_i * 2))

        # 顶部（凹槽顶部）
        for i in range(segments):
            next_i = (i + 1) % segments
            faces.append((base_idx + 1, base_idx + 2 + next_i * 2 + 1, base_idx + 2 + i * 2 + 1))

        return vertices, faces

    @staticmethod
    def _create_separator(x, y, z, radius, height):
        """创建分隔层（锥形）"""
        vertices = []
        faces = []

        segments = 32
        base_idx = len(vertices)

        # 底部大圆
        bottom_radius = radius
        top_radius = radius * 0.6

        # 底部圆心
        vertices.append((x, y, z - height / 2))

        # 顶部圆心
        vertices.append((x, y, z + height / 2))

        # 侧面顶点
        for i in range(segments):
            angle = 2 * math.pi * i / segments

            # 底部顶点
            vx1 = x + bottom_radius * math.cos(angle)
            vy1 = y + bottom_radius * math.sin(angle)
            vertices.append((vx1, vy1, z - height / 2))

            # 顶部顶点
            vx2 = x + top_radius * math.cos(angle)
            vy2 = y + top_radius * math.sin(angle)
            vertices.append((vx2, vy2, z + height / 2))

        # 侧面三角形
        for i in range(segments):
            next_i = (i + 1) % segments

            b1 = base_idx + 2 + i * 2
            b2 = base_idx + 2 + next_i * 2
            t1 = b1 + 1
            t2 = b2 + 1

            faces.append((b1, t1, t2))
            faces.append((b1, t2, b2))

        return vertices, faces

    @staticmethod
    def _calculate_normal(v1, v2, v3):
        """计算法向量"""
        u = [v2[i] - v1[i] for i in range(3)]
        v = [v3[i] - v1[i] for i in range(3)]

        normal = [
            u[1] * v[2] - u[2] * v[1],
            u[2] * v[0] - u[0] * v[2],
            u[0] * v[1] - u[1] * v[0]
        ]

        length = math.sqrt(normal[0] ** 2 + normal[1] ** 2 + normal[2] ** 2)
        if length > 0:
            normal = [n / length for n in normal]

        return normal


# ============================================================================
# 智能G-code处理器 - 简化版
# ============================================================================

class SmartGcodeProcessor:
    """智能G-code处理器"""

    @staticmethod
    def create_auto_script(config):
        """创建自动化处理脚本"""

        script = f'''"""
🎯 Creality Print自动颜色测试处理器
版本: 3.0
功能: 自动处理G-code，添加颜色切换命令
生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
模型: {config['filename']}

使用方法:
  模式1: python {config['filename']}_auto.py --mode auto input.gcode output.gcode
  模式2: python {config['filename']}_auto.py --mode guide
"""

import sys
import os
import re
import json

def generate_color_changes(config):
    """生成颜色切换点"""
    changes = []
    current_z = config["base_height"]

    print("🔍 生成颜色切换点...")

    for color_test in config["colors"]:
        color_name = color_test["name"].split("→")[0].strip()

        for vol_idx, volume in enumerate(color_test["volumes"]):
            layer_z = current_z + vol_idx * config["layer_height"]

            changes.append({{
                "z": layer_z,
                "action": "color_change",
                "from_color": color_name,
                "to_color": "白色",
                "flush_volume": volume,
                "description": f"{{color_name}}→白色，冲刷{{volume}}mm³"
            }})

            print(f"   ✓ Z={{layer_z:.2f}}mm: {{color_name}}→白色 ({{volume}}mm³)")

        current_z += len(color_test["volumes"]) * config["layer_height"]

        if config["colors"].index(color_test) < len(config["colors"]) - 1:
            changes.append({{
                "z": current_z,
                "action": "separator",
                "description": "分隔层开始"
            }})
            current_z += config["separator_height"]

    return changes

def process_gcode(input_file, output_file, color_changes):
    """处理G-code文件"""
    print(f"🔄 处理G-code文件: {{input_file}}")

    try:
        with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()

        output_lines = []
        inserted_count = 0

        for line_num, line in enumerate(lines):
            output_lines.append(line)

            if "G1" in line and "Z" in line:
                match = re.search(r'Z([-+]?[0-9]*\\.?[0-9]+)', line)
                if match:
                    z_value = float(match.group(1))

                    for change in color_changes:
                        if abs(z_value - change["z"]) < 0.01:
                            if change["action"] == "color_change":
                                comment = f"; === 颜色切换点 === {{change['description']}}"
                                m600_cmd = "M600 ; 暂停更换材料"
                                nozzle_clean = "G1 E5 F100 ; 挤出5mm清洁喷嘴"

                                output_lines.append(f"\\n{{comment}}\\n")
                                output_lines.append(f"{{m600_cmd}}\\n")
                                output_lines.append(f"{{nozzle_clean}}\\n")

                                inserted_count += 1
                                print(f"   ✅ 在 Z={{z_value:.2f}}mm 插入颜色切换")

                            break

        with open(output_file, 'w', encoding='utf-8') as f:
            f.writelines(output_lines)

        print(f"✅ 处理完成！")
        print(f"   输入文件: {{input_file}}")
        print(f"   输出文件: {{output_file}}")
        print(f"   插入切换点: {{inserted_count}} 个")

        return True

    except Exception as e:
        print(f"❌ 处理失败: {{e}}")
        return False

def generate_guide(config):
    """生成操作指南"""
    guide = f"""
🎮 Creality Print手动操作指南
=============================

测试配置：
模型名称: {config['filename']}
测试颜色: {', '.join([c['name'].split('→')[0] for c in config['colors']])}

🔍 颜色切换高度表：
"""

    current_z = config["base_height"]
    step_num = 1

    for color_test in config["colors"]:
        color_name = color_test["name"].split("→")[0].strip()

        guide += f"""
【{{color_name}}测试段】
"""

        for vol_idx, volume in enumerate(color_test["volumes"]):
            layer_z = current_z + vol_idx * config["layer_height"]

            guide += f"""
步骤{{step_num}}: Z={{layer_z:.2f}}mm - {{color_name}}→白色
  冲刷量: {{volume}} mm³
  操作:
  1. 监控打印机Z高度显示
  2. 接近{{layer_z:.2f}}mm时准备暂停
  3. 到达高度后立即暂停打印机
  4. 挤出旧材料10-20mm
  5. 剪断旧材料，装入{{color_name}}
  6. 手动挤出约{{volume//3}}mm新材料
  7. 继续打印
"""
            step_num += 1

        current_z += len(color_test["volumes"]) * config["layer_height"]

        if config["colors"].index(color_test) < len(config["colors"]) - 1:
            guide += f"""
• Z={{current_z:.2f}}mm - 分隔层开始
"""
            current_z += config["separator_height"]

    return guide

def main():
    """主函数"""

    if len(sys.argv) < 2:
        print("请指定操作模式！")
        print("可用模式:")
        print("  --mode guide       生成操作指南")
        print("  --mode auto        自动处理G-code")
        print("")
        print("示例:")
        print(f'  python {{sys.argv[0]}} --mode guide')
        print(f'  python {{sys.argv[0]}} --mode auto input.gcode output.gcode')
        return

    mode = sys.argv[1]

    if mode == "--mode" and len(sys.argv) > 2:
        mode = sys.argv[2]

    if mode == "guide":
        # 尝试从配置文件加载
        try:
            config_file = f"{os.path.splitext(sys.argv[0])[0]}_config.json"
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except:
            print("❌ 无法加载配置文件，请确保配置文件存在")
            return

        guide = generate_guide(config)
        print(guide)

        guide_file = f"{config['filename']}_manual_guide.txt"
        with open(guide_file, 'w', encoding='utf-8') as f:
            f.write(guide)

        print(f"✅ 操作指南已保存到: {{guide_file}}")

    elif mode == "auto":
        if len(sys.argv) < 5:
            print("请指定输入和输出文件！")
            print(f"示例: python {{sys.argv[0]}} --mode auto input.gcode output.gcode")
            return

        input_file = sys.argv[3]
        output_file = sys.argv[4]

        if not os.path.exists(input_file):
            print(f"错误: 文件不存在 - {{input_file}}")
            return

        # 尝试从配置文件加载
        try:
            config_file = f"{os.path.splitext(sys.argv[0])[0]}_config.json"
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except:
            print("❌ 无法加载配置文件，请确保配置文件存在")
            return

        color_changes = generate_color_changes(config)
        success = process_gcode(input_file, output_file, color_changes)

        if success:
            print("\\n🎉 处理成功！")
            print("下一步:")
            print("  1. 使用输出文件打印")
            print("  2. 按照指南操作")
            print("  3. 记录测试结果")

    else:
        print(f"未知模式: {{mode}}")

if __name__ == "__main__":
    main()
'''

        return script


# ============================================================================
# 主界面 - 优化版
# ============================================================================

class AutoTestGenerator:
    """自动测试生成器主界面"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Creality自动颜色测试生成器 v18.0")
        self.root.geometry("800x700")

        # 初始化变量
        self.selected_colors = []
        self.color_vars = {}
        self.custom_volume_entries = {}

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""

        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            title_frame,
            text="🎨 Creality自动颜色测试塔生成器",
            font=("Arial", 16, "bold")
        ).pack(side=tk.LEFT)

        ttk.Label(
            title_frame,
            text="v18.0 - 支持颜色单选/多选 + 自定义冲刷量",
            font=("Arial", 9)
        ).pack(side=tk.RIGHT)

        # 创建Notebook（选项卡）
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=10)

        # ==================== 颜色选择选项卡 ====================
        color_frame = ttk.Frame(notebook, padding="15")
        notebook.add(color_frame, text="1. 颜色选择")

        # 选择说明
        ttk.Label(color_frame, text="选择要测试的颜色（可多选）：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))

        # 创建颜色选择框架
        color_select_frame = ttk.Frame(color_frame)
        color_select_frame.pack(fill=tk.BOTH, expand=True)

        # 左列
        left_frame = ttk.Frame(color_select_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # 右列
        right_frame = ttk.Frame(color_select_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 创建颜色复选框
        row = 0
        col = 0
        for i, (color_name, color_info) in enumerate(DEFAULT_COLORS.items()):
            frame = left_frame if i % 2 == 0 else right_frame

            # 创建复选框和标签
            var = tk.BooleanVar(value=(color_name in ["黑色", "红色", "蓝色"]))  # 默认选中几个常用颜色
            self.color_vars[color_name] = var

            # 使用Frame包装，确保复选框和标签在同一行
            item_frame = ttk.Frame(frame)
            item_frame.pack(fill=tk.X, pady=2)

            # 复选框
            cb = ttk.Checkbutton(
                item_frame,
                text=color_name,
                variable=var,
                command=lambda c=color_name: self.update_selected_colors()
            )
            cb.pack(side=tk.LEFT, padx=(0, 10))

            # 颜色标签
            color_label = tk.Label(
                item_frame,
                text="■■",
                fg=color_info["color_code"],
                font=("Arial", 12, "bold"),
                bg="#f0f0f0"
            )
            color_label.pack(side=tk.LEFT, padx=(0, 10))

            # 描述
            ttk.Label(
                item_frame,
                text=color_info["description"],
                font=("Arial", 8),
                foreground="#666666"
            ).pack(side=tk.LEFT)

        # 分隔线
        ttk.Separator(color_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)

        # 冲刷量设置说明
        ttk.Label(color_frame, text="冲刷量设置（单位：mm³）：",
                  font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))

        # 冲刷量预设选项
        volume_frame = ttk.Frame(color_frame)
        volume_frame.pack(fill=tk.X, pady=(0, 10))

        self.volume_preset = tk.StringVar(value="recommended")

        ttk.Radiobutton(
            volume_frame,
            text="使用推荐值",
            variable=self.volume_preset,
            value="recommended"
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Radiobutton(
            volume_frame,
            text="自定义冲刷量",
            variable=self.volume_preset,
            value="custom",
            command=self.toggle_custom_volumes
        ).pack(side=tk.LEFT)

        # 自定义冲刷量输入区域
        self.custom_volume_frame = ttk.Frame(color_frame)

        # 冲刷量输入说明
        ttk.Label(self.custom_volume_frame,
                  text="为每个选中的颜色设置冲刷量（用逗号分隔，如：50,100,150）：",
                  font=("Arial", 9)).pack(anchor=tk.W, pady=(0, 5))

        self.custom_volume_frame.pack_forget()  # 初始隐藏

        # ==================== 模型设置选项卡 ====================
        model_frame = ttk.Frame(notebook, padding="15")
        notebook.add(model_frame, text="2. 模型设置")

        # 输出目录
        dir_frame = ttk.Frame(model_frame)
        dir_frame.pack(fill=tk.X, pady=5)

        ttk.Label(dir_frame, text="保存到:", width=10).pack(side=tk.LEFT)
        self.output_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "color_tests"))
        ttk.Entry(dir_frame, textvariable=self.output_dir_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X,
                                                                              expand=True)
        ttk.Button(dir_frame, text="浏览", command=self.browse_output_dir, width=8).pack(side=tk.LEFT)

        # 文件名
        name_frame = ttk.Frame(model_frame)
        name_frame.pack(fill=tk.X, pady=5)

        ttk.Label(name_frame, text="文件名:", width=10).pack(side=tk.LEFT)
        self.filename_var = tk.StringVar(value=f"flush_test_{datetime.now().strftime('%Y%m%d')}")
        ttk.Entry(name_frame, textvariable=self.filename_var, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X,
                                                                             expand=True)

        # 模型尺寸
        size_frame = ttk.Frame(model_frame)
        size_frame.pack(fill=tk.X, pady=10)

        ttk.Label(size_frame, text="模型尺寸:", width=10).pack(side=tk.LEFT)
        self.width_var = tk.DoubleVar(value=20.0)
        ttk.Spinbox(size_frame, from_=10, to=50, textvariable=self.width_var,
                    width=8, increment=5.0).pack(side=tk.LEFT, padx=5)
        ttk.Label(size_frame, text="×").pack(side=tk.LEFT)
        self.depth_var = tk.DoubleVar(value=20.0)
        ttk.Spinbox(size_frame, from_=10, to=50, textvariable=self.depth_var,
                    width=8, increment=5.0).pack(side=tk.LEFT, padx=5)
        ttk.Label(size_frame, text="mm").pack(side=tk.LEFT)

        # 底座高度
        base_frame = ttk.Frame(model_frame)
        base_frame.pack(fill=tk.X, pady=5)

        ttk.Label(base_frame, text="底座高度:", width=10).pack(side=tk.LEFT)
        self.base_height_var = tk.DoubleVar(value=3.0)
        ttk.Spinbox(base_frame, from_=1, to=10, textvariable=self.base_height_var,
                    width=8, increment=1.0).pack(side=tk.LEFT, padx=5)
        ttk.Label(base_frame, text="mm").pack(side=tk.LEFT)

        # 层高
        layer_frame = ttk.Frame(model_frame)
        layer_frame.pack(fill=tk.X, pady=5)

        ttk.Label(layer_frame, text="层高:", width=10).pack(side=tk.LEFT)
        self.layer_height_var = tk.DoubleVar(value=0.2)
        ttk.Spinbox(layer_frame, from_=0.1, to=0.4, textvariable=self.layer_height_var,
                    width=8, increment=0.1, format="%.2f").pack(side=tk.LEFT, padx=5)
        ttk.Label(layer_frame, text="mm").pack(side=tk.LEFT)

        # 分隔层高度
        sep_frame = ttk.Frame(model_frame)
        sep_frame.pack(fill=tk.X, pady=5)

        ttk.Label(sep_frame, text="分隔层高度:", width=10).pack(side=tk.LEFT)
        self.separator_height_var = tk.DoubleVar(value=1.0)
        ttk.Spinbox(sep_frame, from_=0.5, to=5, textvariable=self.separator_height_var,
                    width=8, increment=0.5).pack(side=tk.LEFT, padx=5)
        ttk.Label(sep_frame, text="mm").pack(side=tk.LEFT)

        # ==================== 预览选项卡 ====================
        preview_frame = ttk.Frame(notebook, padding="15")
        notebook.add(preview_frame, text="3. 预览")

        # 预览文本框
        self.preview_text = scrolledtext.ScrolledText(
            preview_frame,
            height=20,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.status_var = tk.StringVar(value="🟢 就绪 - 请选择测试颜色")
        status_bar = ttk.Frame(main_frame, relief=tk.SUNKEN, padding=(5, 2))
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Label(status_bar, textvariable=self.status_var).pack(side=tk.LEFT)

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="刷新预览",
                   command=self.update_preview, width=15).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="生成测试塔",
                   command=self.generate_all, width=15).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="退出",
                   command=self.root.quit, width=15).pack(side=tk.LEFT, padx=5)

        # 初始化预览
        self.update_selected_colors()
        self.update_preview()

    def update_selected_colors(self):
        """更新选中的颜色"""
        self.selected_colors = []

        for color_name, var in self.color_vars.items():
            if var.get():
                self.selected_colors.append(color_name)

        # 更新自定义冲刷量输入框
        if self.volume_preset.get() == "custom":
            self.create_custom_volume_inputs()

        # 更新状态
        if self.selected_colors:
            self.status_var.set(f"🟡 已选择 {len(self.selected_colors)} 种颜色: {', '.join(self.selected_colors)}")
        else:
            self.status_var.set("🟡 请至少选择一种颜色")

    def toggle_custom_volumes(self):
        """切换自定义冲刷量输入"""
        if self.volume_preset.get() == "custom":
            self.custom_volume_frame.pack(fill=tk.X, pady=(0, 10))
            self.create_custom_volume_inputs()
        else:
            self.custom_volume_frame.pack_forget()

        self.update_preview()

    def create_custom_volume_inputs(self):
        """创建自定义冲刷量输入框"""
        # 清除旧的输入框
        for widget in self.custom_volume_frame.winfo_children():
            if widget not in [self.custom_volume_frame.winfo_children()[0]]:  # 保留说明标签
                widget.destroy()

        # 为每个选中的颜色创建输入框
        for i, color_name in enumerate(self.selected_colors):
            frame = ttk.Frame(self.custom_volume_frame)
            frame.pack(fill=tk.X, pady=2)

            ttk.Label(frame, text=f"{color_name}:", width=10).pack(side=tk.LEFT)

            default_value = ",".join(map(str, DEFAULT_COLORS[color_name]["recommended"]))
            entry_var = tk.StringVar(value=default_value)
            self.custom_volume_entries[color_name] = entry_var

            ttk.Entry(frame, textvariable=entry_var, width=30).pack(side=tk.LEFT, padx=5)

            # 显示推荐值
            recommended_str = ",".join(map(str, DEFAULT_COLORS[color_name]["recommended"]))
            ttk.Label(frame, text=f"(推荐: {recommended_str})",
                      font=("Arial", 8), foreground="#666666").pack(side=tk.LEFT)

    def browse_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if directory:
            self.output_dir_var.set(directory)

    def get_current_config(self):
        """获取当前配置"""
        # 检查是否选择了颜色
        if not self.selected_colors:
            messagebox.showwarning("警告", "请至少选择一种测试颜色！")
            return None

        try:
            # 获取冲刷量设置
            colors_config = []

            if self.volume_preset.get() == "recommended":
                # 使用推荐值
                for color_name in self.selected_colors:
                    colors_config.append({
                        "name": f"{color_name}→白色",
                        "volumes": DEFAULT_COLORS[color_name]["recommended"]
                    })
            else:
                # 使用自定义值
                for color_name in self.selected_colors:
                    if color_name in self.custom_volume_entries:
                        volume_str = self.custom_volume_entries[color_name].get()
                        try:
                            # 解析冲刷量（逗号分隔的数字）
                            volumes = [int(v.strip()) for v in volume_str.split(",") if v.strip()]
                            if not volumes:
                                raise ValueError(f"{color_name}的冲刷量不能为空")

                            # 确保冲刷量为正数
                            volumes = [max(1, v) for v in volumes]

                            colors_config.append({
                                "name": f"{color_name}→白色",
                                "volumes": volumes
                            })
                        except ValueError as e:
                            messagebox.showerror("输入错误",
                                                 f"{color_name}的冲刷量格式错误！\n请使用逗号分隔的数字，如：50,100,150\n错误：{str(e)}")
                            return None

            # 生成唯一文件名
            timestamp = datetime.now().strftime("%H%M%S")
            base_filename = self.filename_var.get().strip()
            if not base_filename:
                base_filename = "flush_test"

            config = {
                "width": float(self.width_var.get()),
                "depth": float(self.depth_var.get()),
                "base_height": float(self.base_height_var.get()),
                "layer_height": float(self.layer_height_var.get()),
                "separator_height": float(self.separator_height_var.get()),
                "colors": colors_config,
                "output_dir": self.output_dir_var.get(),
                "filename": f"{base_filename}_{timestamp}",
            }

            return config

        except ValueError as e:
            messagebox.showerror("配置错误", f"数值格式错误:\n{str(e)}")
            return None

    def update_preview(self):
        """更新预览"""
        config = self.get_current_config()
        if not config:
            # 清空预览
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "请选择至少一种测试颜色以查看预览")
            return

        # 计算统计信息
        total_tests = sum(len(c["volumes"]) for c in config["colors"])
        total_height = config["base_height"] + total_tests * config["layer_height"]

        preview = f"""
📊 测试塔预览
{"=" * 40}

📐 模型信息:
• 尺寸: {config['width']} × {config['depth']} mm
• 总高度: {total_height:.1f} mm
• 底座高度: {config['base_height']} mm
• 层高: {config['layer_height']} mm
• 分隔层: {config['separator_height']} mm

🎨 测试配置:
• 测试颜色: {len(config['colors'])} 种
• 总测试点: {total_tests} 个
• 测试方式: {'推荐冲刷量' if self.volume_preset.get() == 'recommended' else '自定义冲刷量'}

🔍 详细测试计划:
"""

        current_z = config["base_height"]
        for color_test in config["colors"]:
            color_name = color_test["name"].split("→")[0].strip()
            num_tests = len(color_test["volumes"])
            segment_height = num_tests * config["layer_height"]

            preview += f"""
【{color_name}测试段】
• 高度范围: {current_z:.1f} - {current_z + segment_height:.1f} mm
• 测试数量: {num_tests} 个
"""

            for vol_idx, volume in enumerate(color_test["volumes"]):
                layer_z = current_z + vol_idx * config["layer_height"]
                dot_count = min(9, max(1, volume // 30))
                preview += f"  - Z={layer_z:.1f}mm: 冲刷{volume}mm³ (标记: {dot_count}个凹槽)\n"

            current_z += segment_height + config["separator_height"]

        preview += f"""
🚀 生成的文件:
1. {config['filename']}.stl
   - 多颜色标记测试塔
   - 包含可视化标记

2. {config['filename']}_auto.py
   - 自动G-code处理脚本
   - 支持M600自动插入

3. {config['filename']}_guide.txt
   - 详细操作指南
   - 包含故障排除

4. {config['filename']}_config.json
   - 配置文件备份
   - 方便重复使用

📋 操作流程:
1. 在Creality Print中打开STL文件
2. 切片并导出G-code
3. 运行自动脚本处理G-code
4. 使用处理后的G-code打印
5. 按照指南进行颜色切换
"""

        # 更新预览文本框
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(1.0, preview)
        self.preview_text.config(state=tk.NORMAL)

    def generate_all(self):
        """生成所有文件"""
        config = self.get_current_config()
        if not config:
            return

        output_dir = config["output_dir"]

        try:
            # 创建输出目录
            os.makedirs(output_dir, exist_ok=True)

            self.status_var.set("🔄 正在生成文件...")
            self.root.update()

            print("=" * 60)
            print("🎨 Creality自动颜色测试塔生成器 v18.0")
            print("=" * 60)

            # 1. 保存配置文件
            config_file = os.path.join(output_dir, f"{config['filename']}_config.json")
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f"✅ 配置文件: {config_file}")

            # 2. 生成STL文件
            stl_file = os.path.join(output_dir, f"{config['filename']}.stl")
            if SmartSTLGenerator.generate_color_tower(config, stl_file):
                print(f"✅ STL文件: {stl_file}")
            else:
                raise Exception("STL生成失败")

            # 3. 生成操作指南
            guide = AutoConfigGenerator.create_auto_guide(config)
            guide_file = os.path.join(output_dir, f"{config['filename']}_guide.txt")
            with open(guide_file, 'w', encoding='utf-8') as f:
                f.write(guide)
            print(f"✅ 操作指南: {guide_file}")

            # 4. 生成自动脚本
            script = SmartGcodeProcessor.create_auto_script(config)
            script_file = os.path.join(output_dir, f"{config['filename']}_auto.py")
            with open(script_file, 'w', encoding='utf-8') as f:
                f.write(script)
            print(f"✅ 自动脚本: {script_file}")

            # 5. 生成快速开始脚本
            quick_start = self.create_quick_start_script(config)
            quick_file = os.path.join(output_dir, f"{config['filename']}_quick.py")
            with open(quick_file, 'w', encoding='utf-8') as f:
                f.write(quick_start)
            print(f"✅ 快速开始: {quick_file}")

            # 完成
            print("\n" + "=" * 60)
            print("🎉 所有文件已生成完成！")
            print(f"📁 输出目录: {output_dir}")
            print(f"🎨 测试颜色: {len(config['colors'])} 种")
            print(f"🔢 测试点数: {sum(len(c['volumes']) for c in config['colors'])} 个")
            print("\n🚀 下一步操作:")
            print(f"  1. 打开目录: {output_dir}")
            print(f"  2. 在Creality Print中打开: {config['filename']}.stl")
            print(f"  3. 切片后运行: python {config['filename']}_auto.py --mode auto")
            print("  4. 按照指南操作")
            print("=" * 60)

            self.status_var.set("✅ 生成完成！")

            # 显示完成对话框
            self.show_success_dialog(output_dir, config['filename'])

        except Exception as e:
            self.status_var.set("❌ 生成失败")
            messagebox.showerror("生成错误", f"生成过程中发生错误:\n{str(e)}")
            print(f"❌ 错误: {e}")
            import traceback
            traceback.print_exc()

    def create_quick_start_script(self, config):
        """创建快速开始脚本"""
        return f'''"""
🚀 快速开始脚本 - {config['filename']}
运行此脚本开始颜色测试
"""

import os
import sys

def main():
    print("🎯 Creality颜色测试快速开始")
    print("=" * 50)

    # 显示测试信息
    print("测试配置:")
    print(f"• 模型名称: {config['filename']}")
    print(f"• 测试颜色: {len(config['colors'])} 种")
    print(f"• 总测试点: {sum(len(c['volumes']) for c in config['colors'])} 个")

    print("\\n📋 操作步骤:")
    print("1. 在Creality Print中打开 .stl 文件")
    print("2. 切片并导出G-code文件")
    print("3. 运行以下命令处理G-code:")
    print(f'   python {config["filename"]}_auto.py --mode auto 你的文件.gcode 输出文件.gcode')
    print("4. 使用处理后的G-code打印")
    print("5. 按照指南操作颜色切换")

    print("\\n💡 小贴士:")
    print("• 使用白色材料作为主体")
    print("• 深色材料需要更多冲刷量")
    print("• 记录测试结果以便优化")

    input("\\n按Enter键退出...")

if __name__ == "__main__":
    main()
'''

    def show_success_dialog(self, output_dir, filename):
        """显示成功对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("🎉 生成成功！")
        dialog.geometry("500x450")

        text = f"""✅ 所有文件已成功生成！

📁 输出目录：
{output_dir}

📋 生成的文件：
1. {filename}.stl
   - 多颜色标记测试塔
   - 包含可视化冲刷量标记

2. {filename}_auto.py
   - ★自动处理脚本★
   - 自动插入M600命令
   - 简化操作流程

3. {filename}_guide.txt
   - 完整操作指南
   - 包含故障排除和结果记录

4. {filename}_config.json
   - 配置文件
   - 可重复使用或修改

5. {filename}_quick.py
   - 快速开始脚本
   - 简化启动流程

🚀 立即开始：
1. 打开输出目录查看文件
2. 仔细阅读操作指南
3. 按照步骤进行测试
4. 记录测试结果优化参数
"""

        text_widget = tk.Text(dialog, wrap=tk.WORD, height=22)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_widget.insert("1.0", text)
        text_widget.config(state=tk.DISABLED)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="📂 打开目录",
                   command=lambda: os.startfile(output_dir)).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="📖 查看指南",
                   command=lambda: self.open_guide(output_dir, filename)).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="关闭",
                   command=dialog.destroy).pack(side=tk.LEFT, padx=5)

    def open_guide(self, output_dir, filename):
        """打开操作指南"""
        guide_file = os.path.join(output_dir, f"{filename}_guide.txt")
        if os.path.exists(guide_file):
            try:
                # Windows
                os.startfile(guide_file)
            except:
                # Linux/Mac
                import subprocess
                subprocess.run(['xdg-open', guide_file])

    def run(self):
        """运行程序"""
        self.root.mainloop()


# ============================================================================
# 启动程序
# ============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("🎨 Creality自动颜色测试塔生成器 v18.0")
    print("=" * 60)
    print("功能特性:")
    print("• 颜色单选/多选（默认选中黑、红、蓝）")
    print("• 支持自定义冲刷量设置")
    print("• 可视化标记测试塔")
    print("• 自动G-code处理脚本")
    print("=" * 60)

    app = AutoTestGenerator()
    app.run()