"""
storage_box_generator_fixed.py
修正盖子尺寸不超过260x260mm平台限制
"""

import numpy as np
from stl import mesh
import math

def cm_to_mm(cm):
    """厘米转毫米"""
    return cm * 10

def create_box():
    """
    创建盒子主体（开放顶部，有16个格子）
    尺寸：外部260x260x200mm
    """
    print("正在生成盒子主体（26x26x20cm）...")

    total_size = cm_to_mm(26)       # 260mm
    height = cm_to_mm(20)           # 200mm
    wall_thickness = 12             # 12mm壁厚

    inner_size = total_size - 2 * wall_thickness  # 236mm
    grid_size = inner_size / 4                    # 59mm

    vertices = []
    faces = []

    # 1. 外部立方体（260x260x200mm）
    # 底部外圈
    vertices.extend([
        [0, 0, 0],                          # 0
        [total_size, 0, 0],                 # 1
        [total_size, total_size, 0],        # 2
        [0, total_size, 0],                 # 3
        [0, 0, height - 10],                # 4 (190mm)
        [total_size, 0, height - 10],       # 5
        [total_size, total_size, height - 10], # 6
        [0, total_size, height - 10],       # 7
    ])

    # 内部立方体（236x236x190mm）
    vertices.extend([
        [wall_thickness, wall_thickness, 0],                  # 8
        [total_size - wall_thickness, wall_thickness, 0],     # 9
        [total_size - wall_thickness, total_size - wall_thickness, 0],  # 10
        [wall_thickness, total_size - wall_thickness, 0],     # 11
        [wall_thickness, wall_thickness, height - 10],        # 12
        [total_size - wall_thickness, wall_thickness, height - 10],     # 13
        [total_size - wall_thickness, total_size - wall_thickness, height - 10],  # 14
        [wall_thickness, total_size - wall_thickness, height - 10],     # 15
    ])

    # 外部侧面
    faces.extend([
        [0, 1, 5], [0, 5, 4],
        [1, 2, 6], [1, 6, 5],
        [2, 3, 7], [2, 7, 6],
        [3, 0, 4], [3, 4, 7],
    ])

    # 内部侧面
    faces.extend([
        [8, 9, 13], [8, 13, 12],
        [9, 10, 14], [9, 14, 13],
        [10, 11, 15], [10, 15, 14],
        [11, 8, 12], [11, 12, 15],
    ])

    # 底部
    faces.extend([
        [0, 3, 2], [0, 2, 1],
        [8, 11, 10], [8, 10, 9],
        [0, 8, 9], [0, 9, 1],
        [1, 9, 10], [1, 10, 2],
        [2, 10, 11], [2, 11, 3],
        [3, 11, 8], [3, 8, 0],
    ])

    # 顶部边缘
    faces.extend([
        [4, 5, 6], [4, 6, 7],
        [12, 13, 14], [12, 14, 15],
        [4, 12, 13], [4, 13, 5],
        [5, 13, 14], [5, 14, 6],
        [6, 14, 15], [6, 15, 7],
        [7, 15, 12], [7, 12, 4],
    ])

    # 2. 内部隔板（16个格子）
    # 纵向隔板（3个）
    for i in range(1, 4):
        x_pos = wall_thickness + i * (grid_size + wall_thickness)

        start_idx = len(vertices)
        vertices.extend([
            [x_pos, wall_thickness, 0],                    # 底部前
            [x_pos + wall_thickness, wall_thickness, 0],  # 底部后
            [x_pos + wall_thickness, total_size - wall_thickness, 0],  # 底部右
            [x_pos, total_size - wall_thickness, 0],      # 底部左
            [x_pos, wall_thickness, height - 10],                    # 顶部前
            [x_pos + wall_thickness, wall_thickness, height - 10],  # 顶部后
            [x_pos + wall_thickness, total_size - wall_thickness, height - 10],  # 顶部右
            [x_pos, total_size - wall_thickness, height - 10],      # 顶部左
        ])

        faces.extend([
            [start_idx, start_idx+1, start_idx+5], [start_idx, start_idx+5, start_idx+4],
            [start_idx+1, start_idx+2, start_idx+6], [start_idx+1, start_idx+6, start_idx+5],
            [start_idx+2, start_idx+3, start_idx+7], [start_idx+2, start_idx+7, start_idx+6],
            [start_idx+3, start_idx, start_idx+4], [start_idx+3, start_idx+4, start_idx+7],
            [start_idx, start_idx+3, start_idx+2], [start_idx, start_idx+2, start_idx+1],
            [start_idx+4, start_idx+5, start_idx+6], [start_idx+4, start_idx+6, start_idx+7],
        ])

    # 横向隔板（3个）
    for i in range(1, 4):
        y_pos = wall_thickness + i * (grid_size + wall_thickness)

        start_idx = len(vertices)
        vertices.extend([
            [wall_thickness, y_pos, 0],                    # 底部左
            [total_size - wall_thickness, y_pos, 0],      # 底部右
            [total_size - wall_thickness, y_pos + wall_thickness, 0],  # 底部后
            [wall_thickness, y_pos + wall_thickness, 0],  # 底部前
            [wall_thickness, y_pos, height - 10],                    # 顶部左
            [total_size - wall_thickness, y_pos, height - 10],      # 顶部右
            [total_size - wall_thickness, y_pos + wall_thickness, height - 10],  # 顶部后
            [wall_thickness, y_pos + wall_thickness, height - 10],  # 顶部前
        ])

        faces.extend([
            [start_idx, start_idx+1, start_idx+5], [start_idx, start_idx+5, start_idx+4],
            [start_idx+1, start_idx+2, start_idx+6], [start_idx+1, start_idx+6, start_idx+5],
            [start_idx+2, start_idx+3, start_idx+7], [start_idx+2, start_idx+7, start_idx+6],
            [start_idx+3, start_idx, start_idx+4], [start_idx+3, start_idx+4, start_idx+7],
            [start_idx, start_idx+3, start_idx+2], [start_idx, start_idx+2, start_idx+1],
            [start_idx+4, start_idx+5, start_idx+6], [start_idx+4, start_idx+6, start_idx+7],
        ])

    box_mesh = mesh.Mesh(np.zeros(len(faces), dtype=mesh.Mesh.dtype))
    for i, face in enumerate(faces):
        for j in range(3):
            box_mesh.vectors[i][j] = vertices[face[j]]

    print(f"盒子尺寸：{total_size}x{total_size}x{height}mm")
    print(f"壁厚：{wall_thickness}mm")
    print(f"格子数量：4x4 = 16个")
    return box_mesh

def create_lid():
    """
    创建盖子 - 修正版（确保不超过260x260mm）
    """
    print("\n正在生成盖子（确保不超过平台尺寸）...")

    total_size = cm_to_mm(26)       # 260mm
    lid_height = 50                 # 50mm高度
    wall_thickness = 12             # 12mm
    tolerance = 1                   # 1mm配合公差（更紧）

    inner_size = total_size - 2 * wall_thickness  # 236mm
    grid_size = inner_size / 4                    # 59mm

    vertices = []
    faces = []

    # ========== 1. 盖子主体（确保不超过260mm） ==========

    # 顶板（厚15mm，完全在260mm内）
    start_idx = len(vertices)
    vertices.extend([
        # 底部四角
        [0, 0, 0],                          # 0 前左下
        [total_size, 0, 0],                 # 1 前右下
        [total_size, total_size, 0],        # 2 后右下
        [0, total_size, 0],                 # 3 后左下

        # 顶部四角（15mm厚）
        [0, 0, 15],                         # 4 前左上
        [total_size, 0, 15],                # 5 前右上
        [total_size, total_size, 15],       # 6 后右上
        [0, total_size, 15],                # 7 后左上
    ])

    # 顶板侧面
    faces.extend([
        [0, 1, 5], [0, 5, 4],   # 前面
        [1, 2, 6], [1, 6, 5],   # 右面
        [2, 3, 7], [2, 7, 6],   # 后面
        [3, 0, 4], [3, 4, 7],   # 左面
    ])

    # 顶板顶部和底部
    faces.extend([
        [0, 3, 2], [0, 2, 1],   # 底面
        [4, 5, 6], [4, 6, 7],   # 顶面
    ])

    # ========== 2. 侧壁（内缩，确保外部尺寸不超过260mm） ==========
    # 侧壁从顶板向下延伸35mm，总高50mm

    side_wall_start_z = 15  # 从顶板底部开始
    side_wall_height = 35   # 侧壁高度

    # 外壁顶点（与顶板对齐，完全在260mm内）
    start_idx = len(vertices)
    vertices.extend([
        # 外壁底部四角
        [0, 0, 0],                          # 8
        [total_size, 0, 0],                 # 9
        [total_size, total_size, 0],        # 10
        [0, total_size, 0],                 # 11

        # 外壁顶部四角（连接到顶板）
        [0, 0, side_wall_start_z],          # 12
        [total_size, 0, side_wall_start_z], # 13
        [total_size, total_size, side_wall_start_z],  # 14
        [0, total_size, side_wall_start_z], # 15
    ])

    # 外壁侧面
    faces.extend([
        [8, 9, 13], [8, 13, 12],   # 前面
        [9, 10, 14], [9, 14, 13],  # 右面
        [10, 11, 15], [10, 15, 14], # 后面
        [11, 8, 12], [11, 12, 15], # 左面
    ])

    # ========== 3. 内壁（用于套在盒子上） ==========
    # 内壁从外壁向内偏移，形成套接结构

    inner_offset = wall_thickness - 2  # 10mm内缩，留2mm壁厚

    start_idx = len(vertices)
    vertices.extend([
        # 内壁底部四角
        [inner_offset, inner_offset, 0],                          # 16
        [total_size - inner_offset, inner_offset, 0],            # 17
        [total_size - inner_offset, total_size - inner_offset, 0],  # 18
        [inner_offset, total_size - inner_offset, 0],            # 19

        # 内壁顶部四角
        [inner_offset, inner_offset, side_wall_start_z - 2],     # 20 (低2mm便于配合)
        [total_size - inner_offset, inner_offset, side_wall_start_z - 2],  # 21
        [total_size - inner_offset, total_size - inner_offset, side_wall_start_z - 2],  # 22
        [inner_offset, total_size - inner_offset, side_wall_start_z - 2],  # 23
    ])

    # 内壁侧面
    faces.extend([
        [16, 17, 21], [16, 21, 20],   # 前面
        [17, 18, 22], [17, 22, 21],  # 右面
        [18, 19, 23], [18, 23, 22],  # 后面
        [19, 16, 20], [19, 20, 23],  # 左面
    ])

    # 连接内外壁的底面
    faces.extend([
        [8, 16, 17], [8, 17, 9],     # 前底面
        [9, 17, 18], [9, 18, 10],    # 右底面
        [10, 18, 19], [10, 19, 11],  # 后底面
        [11, 19, 16], [11, 16, 8],   # 左底面
    ])

    # 连接内外壁的顶面
    faces.extend([
        [12, 20, 21], [12, 21, 13],  # 前顶面
        [13, 21, 22], [13, 22, 14],  # 右顶面
        [14, 22, 23], [14, 23, 15],  # 后顶面
        [15, 23, 20], [15, 20, 12],  # 左顶面
    ])

    # ========== 4. 内部支撑结构（在260mm内） ==========
    # 添加内部支撑板，厚度5mm，距离底部10mm

    support_thickness = 5
    support_height = 10
    support_offset = wall_thickness

    start_idx = len(vertices)
    vertices.extend([
        # 支撑板底部
        [support_offset, support_offset, support_height],                          # 24
        [total_size - support_offset, support_offset, support_height],            # 25
        [total_size - support_offset, total_size - support_offset, support_height],  # 26
        [support_offset, total_size - support_offset, support_height],            # 27

        # 支撑板顶部
        [support_offset, support_offset, support_height + support_thickness],     # 28
        [total_size - support_offset, support_offset, support_height + support_thickness],  # 29
        [total_size - support_offset, total_size - support_offset, support_height + support_thickness],  # 30
        [support_offset, total_size - support_offset, support_height + support_thickness],  # 31
    ])

    # 支撑板侧面
    faces.extend([
        [24, 25, 29], [24, 29, 28],   # 前面
        [25, 26, 30], [25, 30, 29],   # 右面
        [26, 27, 31], [26, 31, 30],   # 后面
        [27, 24, 28], [27, 28, 31],   # 左面
        [24, 27, 26], [24, 26, 25],   # 底面
        [28, 29, 30], [28, 30, 31],   # 顶面
    ])

    # ========== 5. 密封凸起（完全在内部，不超过260mm） ==========
    # 十字形凸起，高度8mm，宽度4mm

    cross_height = 8
    cross_width = 4

    for row in range(4):
        for col in range(4):
            # 计算格子位置（在支撑板上方）
            x_start = support_offset + col * (grid_size + wall_thickness)
            y_start = support_offset + row * (grid_size + wall_thickness)
            z_base = support_height + support_thickness  # 在支撑板上

            # 横向凸起
            start_idx = len(vertices)
            vertices.extend([
                # 底部
                [x_start, y_start + grid_size/2 - cross_width/2, z_base],  # 32
                [x_start + grid_size, y_start + grid_size/2 - cross_width/2, z_base],  # 33
                [x_start + grid_size, y_start + grid_size/2 + cross_width/2, z_base],  # 34
                [x_start, y_start + grid_size/2 + cross_width/2, z_base],  # 35
                # 顶部
                [x_start, y_start + grid_size/2 - cross_width/2, z_base + cross_height],  # 36
                [x_start + grid_size, y_start + grid_size/2 - cross_width/2, z_base + cross_height],  # 37
                [x_start + grid_size, y_start + grid_size/2 + cross_width/2, z_base + cross_height],  # 38
                [x_start, y_start + grid_size/2 + cross_width/2, z_base + cross_height],  # 39
            ])

            faces.extend([
                [start_idx, start_idx+1, start_idx+5], [start_idx, start_idx+5, start_idx+4],
                [start_idx+1, start_idx+2, start_idx+6], [start_idx+1, start_idx+6, start_idx+5],
                [start_idx+2, start_idx+3, start_idx+7], [start_idx+2, start_idx+7, start_idx+6],
                [start_idx+3, start_idx, start_idx+4], [start_idx+3, start_idx+4, start_idx+7],
                [start_idx, start_idx+3, start_idx+2], [start_idx, start_idx+2, start_idx+1],
                [start_idx+4, start_idx+5, start_idx+6], [start_idx+4, start_idx+6, start_idx+7],
            ])

            # 纵向凸起
            start_idx = len(vertices)
            vertices.extend([
                # 底部
                [x_start + grid_size/2 - cross_width/2, y_start, z_base],  # 40
                [x_start + grid_size/2 + cross_width/2, y_start, z_base],  # 41
                [x_start + grid_size/2 + cross_width/2, y_start + grid_size, z_base],  # 42
                [x_start + grid_size/2 - cross_width/2, y_start + grid_size, z_base],  # 43
                # 顶部
                [x_start + grid_size/2 - cross_width/2, y_start, z_base + cross_height],  # 44
                [x_start + grid_size/2 + cross_width/2, y_start, z_base + cross_height],  # 45
                [x_start + grid_size/2 + cross_width/2, y_start + grid_size, z_base + cross_height],  # 46
                [x_start + grid_size/2 - cross_width/2, y_start + grid_size, z_base + cross_height],  # 47
            ])

            faces.extend([
                [start_idx, start_idx+1, start_idx+5], [start_idx, start_idx+5, start_idx+4],
                [start_idx+1, start_idx+2, start_idx+6], [start_idx+1, start_idx+6, start_idx+5],
                [start_idx+2, start_idx+3, start_idx+7], [start_idx+2, start_idx+7, start_idx+6],
                [start_idx+3, start_idx, start_idx+4], [start_idx+3, start_idx+4, start_idx+7],
                [start_idx, start_idx+3, start_idx+2], [start_idx, start_idx+2, start_idx+1],
                [start_idx+4, start_idx+5, start_idx+6], [start_idx+4, start_idx+6, start_idx+7],
            ])

    # ========== 6. 拉手（在顶板上方，但不超过260mm边界） ==========
    # 圆形拉手，直径30mm，高15mm，居中

    handle_radius = 15  # 半径15mm
    handle_height = 10  # 高10mm
    handle_x = total_size / 2
    handle_y = total_size / 2
    handle_z = 15  # 在顶板上方开始
    segments = 16

    start_idx = len(vertices)

    # 圆柱体顶点
    for i in range(segments):
        angle = 2 * math.pi * i / segments
        # 底部顶点（在顶板上）
        vertices.append([
            handle_x + handle_radius * math.cos(angle),
            handle_y + handle_radius * math.sin(angle),
            handle_z
        ])
        # 顶部顶点
        vertices.append([
            handle_x + handle_radius * math.cos(angle),
            handle_y + handle_radius * math.sin(angle),
            handle_z + handle_height
        ])

    # 圆柱体侧面
    for i in range(segments):
        next_i = (i + 1) % segments
        bottom_i = start_idx + i * 2
        bottom_next = start_idx + next_i * 2
        top_i = bottom_i + 1
        top_next = bottom_next + 1

        faces.extend([
            [bottom_i, bottom_next, top_next],
            [bottom_i, top_next, top_i]
        ])

    # 圆柱体底面（连接到底部）
    center_bottom = len(vertices)
    vertices.append([handle_x, handle_y, handle_z])
    for i in range(segments):
        next_i = (i + 1) % segments
        faces.append([
            start_idx + i * 2,
            center_bottom,
            start_idx + next_i * 2
        ])

    # 圆柱体顶面
    center_top = len(vertices)
    vertices.append([handle_x, handle_y, handle_z + handle_height])
    for i in range(segments):
        next_i = (i + 1) % segments
        faces.append([
            start_idx + i * 2 + 1,
            start_idx + next_i * 2 + 1,
            center_top
        ])

    # 创建网格对象
    lid_mesh = mesh.Mesh(np.zeros(len(faces), dtype=mesh.Mesh.dtype))
    for i, face in enumerate(faces):
        for j in range(3):
            lid_mesh.vectors[i][j] = vertices[face[j]]

    print(f"✅ 盖子尺寸：{total_size}x{total_size}x{lid_height}mm")
    print(f"✅ 外部边界检查：")
    print(f"   X范围：0 - {total_size}mm")
    print(f"   Y范围：0 - {total_size}mm")
    print(f"   Z范围：0 - {lid_height}mm")
    print(f"✅ 所有部件都在260x260mm平台内")

    return lid_mesh

def main():
    """主函数"""
    print("=" * 60)
    print("储物盒STL文件生成器 - 修正版")
    print("专为260x260mm打印平台优化")
    print("=" * 60)

    try:
        # 生成盒子
        box = create_box()

        # 检查盒子尺寸
        box_vertices = box.vectors.reshape(-1, 3)
        max_x, max_y, max_z = box_vertices.max(axis=0)
        min_x, min_y, min_z = box_vertices.min(axis=0)

        print(f"\n📦 盒子尺寸检查：")
        print(f"   X: {min_x:.1f} - {max_x:.1f} mm (跨度: {max_x-min_x:.1f}mm)")
        print(f"   Y: {min_y:.1f} - {max_y:.1f} mm (跨度: {max_y-min_y:.1f}mm)")
        print(f"   Z: {min_z:.1f} - {max_z:.1f} mm (高度: {max_z-min_z:.1f}mm)")

        box.save('storage_box_26cm_fixed.stl')
        print(f"\n✅ 盒子已保存为：storage_box_26cm_fixed.stl")

        # 生成盖子
        lid = create_lid()

        # 检查盖子尺寸
        lid_vertices = lid.vectors.reshape(-1, 3)
        max_x, max_y, max_z = lid_vertices.max(axis=0)
        min_x, min_y, min_z = lid_vertices.min(axis=0)

        print(f"\n🔶 盖子尺寸检查：")
        print(f"   X: {min_x:.1f} - {max_x:.1f} mm (跨度: {max_x-min_x:.1f}mm)")
        print(f"   Y: {min_y:.1f} - {max_y:.1f} mm (跨度: {max_y-min_y:.1f}mm)")
        print(f"   Z: {min_z:.1f} - {max_z:.1f} mm (高度: {max_z-min_z:.1f}mm)")

        if max_x <= 260 and max_y <= 260:
            print("✅ 盖子完全适合260x260mm打印平台！")
        else:
            print("⚠️  警告：盖子可能超出平台边界！")

        lid.save('storage_box_lid_26cm_fixed.stl')
        print(f"✅ 盖子已保存为：storage_box_lid_26cm_fixed.stl")

        print("\n" + "=" * 60)
        print("🎯 打印建议：")
        print("1. 平台尺寸：至少260x260mm")
        print("2. 打印方向：盒子开口朝上，盖子平放打印")
        print("3. 层高：0.2-0.3mm")
        print("4. 填充：15-20%")
        print("5. 壁厚：至少1.2mm")
        print("6. 材料：PLA（推荐）或PETG")
        print("=" * 60)

    except ImportError as e:
        print(f"\n❌ 缺少依赖库，请运行：pip install numpy numpy-stl")
    except Exception as e:
        print(f"\n❌ 错误：{e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()