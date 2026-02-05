"""
storage_box_true_hollow.py
创建真正的空心储物盒 - 使用"壳"方法而不是实心块
"""

import numpy as np
import zipfile
import os
from datetime import datetime
import math

# 打印机平台尺寸
PRINTER_X = 260.0
PRINTER_Y = 260.0
PRINTER_Z = 255.0

def create_true_hollow_box():
    """创建真正的空心盒子 - 只有2mm厚的壁"""
    print("正在创建真正的空心储物盒...")

    # 格子尺寸
    GRID_WIDTH = 60.0
    GRID_DEPTH = 60.0
    GRID_HEIGHT = 200.0

    # 布局
    COLS = 4
    ROWS = 3

    # 结构参数
    WALL_THICKNESS = 2.0
    PARTITION_THICKNESS = 2.0
    BOTTOM_THICKNESS = 2.0

    # 计算尺寸
    total_inner_width = COLS * GRID_WIDTH + (COLS - 1) * PARTITION_THICKNESS
    total_inner_depth = ROWS * GRID_DEPTH + (ROWS - 1) * PARTITION_THICKNESS
    total_outer_width = total_inner_width + 2 * WALL_THICKNESS
    total_outer_depth = total_inner_depth + 2 * WALL_THICKNESS
    total_height = GRID_HEIGHT + BOTTOM_THICKNESS

    print(f"📦 设计:")
    print(f"  外部: {total_outer_width}x{total_outer_depth}x{total_height}mm")
    print(f"  壁厚: {WALL_THICKNESS}mm")
    print(f"  格子: {COLS}x{ROWS}个，每个{GRID_WIDTH}x{GRID_DEPTH}x{GRID_HEIGHT}mm")

    vertices = []
    triangles = []

    # ========== 方法：创建2mm厚的"壳" ==========
    # 思路：先创建外表面，再创建内表面，中间是空的

    # 1. 外表面顶点
    print("\n🔧 创建外表面...")

    # 外底面四个角
    v0 = [0, 0, 0]
    v1 = [total_outer_width, 0, 0]
    v2 = [total_outer_width, total_outer_depth, 0]
    v3 = [0, total_outer_depth, 0]

    # 外壁顶部四个角
    v4 = [0, 0, total_height]
    v5 = [total_outer_width, 0, total_height]
    v6 = [total_outer_width, total_outer_depth, total_height]
    v7 = [0, total_outer_depth, total_height]

    outer_indices = []
    for v in [v0, v1, v2, v3, v4, v5, v6, v7]:
        vertices.append(v)
        outer_indices.append(len(vertices) - 1)

    # 外表面三角形
    # 外底面（朝下）
    triangles.append([outer_indices[2], outer_indices[3], outer_indices[0]])
    triangles.append([outer_indices[2], outer_indices[0], outer_indices[1]])

    # 外壁侧面
    triangles.append([outer_indices[0], outer_indices[1], outer_indices[5]])
    triangles.append([outer_indices[0], outer_indices[5], outer_indices[4]])
    triangles.append([outer_indices[1], outer_indices[2], outer_indices[6]])
    triangles.append([outer_indices[1], outer_indices[6], outer_indices[5]])
    triangles.append([outer_indices[2], outer_indices[3], outer_indices[7]])
    triangles.append([outer_indices[2], outer_indices[7], outer_indices[6]])
    triangles.append([outer_indices[3], outer_indices[0], outer_indices[4]])
    triangles.append([outer_indices[3], outer_indices[4], outer_indices[7]])

    # 2. 内表面顶点（向内偏移WALL_THICKNESS）
    print("🔧 创建内表面...")

    # 内底面（在底部厚度上方）
    v8 = [WALL_THICKNESS, WALL_THICKNESS, BOTTOM_THICKNESS]
    v9 = [total_outer_width - WALL_THICKNESS, WALL_THICKNESS, BOTTOM_THICKNESS]
    v10 = [total_outer_width - WALL_THICKNESS, total_outer_depth - WALL_THICKNESS, BOTTOM_THICKNESS]
    v11 = [WALL_THICKNESS, total_outer_depth - WALL_THICKNESS, BOTTOM_THICKNESS]

    # 内壁顶部
    v12 = [WALL_THICKNESS, WALL_THICKNESS, total_height]
    v13 = [total_outer_width - WALL_THICKNESS, WALL_THICKNESS, total_height]
    v14 = [total_outer_width - WALL_THICKNESS, total_outer_depth - WALL_THICKNESS, total_height]
    v15 = [WALL_THICKNESS, total_outer_depth - WALL_THICKNESS, total_height]

    inner_indices = []
    for v in [v8, v9, v10, v11, v12, v13, v14, v15]:
        vertices.append(v)
        inner_indices.append(len(vertices) - 1)

    # 内表面三角形（注意法向量方向应该向内）
    # 内底面（朝上）
    triangles.append([inner_indices[0], inner_indices[1], inner_indices[2]])
    triangles.append([inner_indices[0], inner_indices[2], inner_indices[3]])

    # 内壁侧面（从内向外看）
    triangles.append([inner_indices[1], inner_indices[0], inner_indices[4]])
    triangles.append([inner_indices[1], inner_indices[4], inner_indices[5]])
    triangles.append([inner_indices[2], inner_indices[1], inner_indices[5]])
    triangles.append([inner_indices[2], inner_indices[5], inner_indices[6]])
    triangles.append([inner_indices[3], inner_indices[2], inner_indices[6]])
    triangles.append([inner_indices[3], inner_indices[6], inner_indices[7]])
    triangles.append([inner_indices[0], inner_indices[3], inner_indices[7]])
    triangles.append([inner_indices[0], inner_indices[7], inner_indices[4]])

    # 3. 纵向隔板（作为独立的薄壁）
    print(f"🔧 创建{COLS-1}条纵向隔板...")

    for col in range(1, COLS):
        # 隔板位置
        x_pos = WALL_THICKNESS + col * GRID_WIDTH + (col - 1) * PARTITION_THICKNESS

        # 隔板是2mm厚的独立薄壁
        # 前侧顶点
        v16 = [x_pos, WALL_THICKNESS, BOTTOM_THICKNESS]
        v17 = [x_pos + PARTITION_THICKNESS, WALL_THICKNESS, BOTTOM_THICKNESS]
        v18 = [x_pos + PARTITION_THICKNESS, total_outer_depth - WALL_THICKNESS, BOTTOM_THICKNESS]
        v19 = [x_pos, total_outer_depth - WALL_THICKNESS, BOTTOM_THICKNESS]

        # 后侧顶点
        v20 = [x_pos, WALL_THICKNESS, total_height]
        v21 = [x_pos + PARTITION_THICKNESS, WALL_THICKNESS, total_height]
        v22 = [x_pos + PARTITION_THICKNESS, total_outer_depth - WALL_THICKNESS, total_height]
        v23 = [x_pos, total_outer_depth - WALL_THICKNESS, total_height]

        partition_indices = []
        for v in [v16, v17, v18, v19, v20, v21, v22, v23]:
            vertices.append(v)
            partition_indices.append(len(vertices) - 1)

        # 隔板前侧面（面向左侧格子）
        triangles.append([partition_indices[0], partition_indices[1], partition_indices[5]])
        triangles.append([partition_indices[0], partition_indices[5], partition_indices[4]])

        # 隔板后侧面（面向右侧格子）
        triangles.append([partition_indices[2], partition_indices[3], partition_indices[7]])
        triangles.append([partition_indices[2], partition_indices[7], partition_indices[6]])

        # 隔板左侧面（厚度方向）
        triangles.append([partition_indices[3], partition_indices[0], partition_indices[4]])
        triangles.append([partition_indices[3], partition_indices[4], partition_indices[7]])

        # 隔板右侧面（厚度方向）
        triangles.append([partition_indices[1], partition_indices[2], partition_indices[6]])
        triangles.append([partition_indices[1], partition_indices[6], partition_indices[5]])

        # 隔板底面
        triangles.append([partition_indices[0], partition_indices[3], partition_indices[2]])
        triangles.append([partition_indices[0], partition_indices[2], partition_indices[1]])

        # 隔板顶面
        triangles.append([partition_indices[4], partition_indices[5], partition_indices[6]])
        triangles.append([partition_indices[4], partition_indices[6], partition_indices[7]])

    # 4. 横向隔板
    print(f"🔧 创建{ROWS-1}条横向隔板...")

    for row in range(1, ROWS):
        # 隔板位置
        y_pos = WALL_THICKNESS + row * GRID_DEPTH + (row - 1) * PARTITION_THICKNESS

        # 隔板顶点
        v24 = [WALL_THICKNESS, y_pos, BOTTOM_THICKNESS]
        v25 = [total_outer_width - WALL_THICKNESS, y_pos, BOTTOM_THICKNESS]
        v26 = [total_outer_width - WALL_THICKNESS, y_pos + PARTITION_THICKNESS, BOTTOM_THICKNESS]
        v27 = [WALL_THICKNESS, y_pos + PARTITION_THICKNESS, BOTTOM_THICKNESS]

        v28 = [WALL_THICKNESS, y_pos, total_height]
        v29 = [total_outer_width - WALL_THICKNESS, y_pos, total_height]
        v30 = [total_outer_width - WALL_THICKNESS, y_pos + PARTITION_THICKNESS, total_height]
        v31 = [WALL_THICKNESS, y_pos + PARTITION_THICKNESS, total_height]

        partition_indices = []
        for v in [v24, v25, v26, v27, v28, v29, v30, v31]:
            vertices.append(v)
            partition_indices.append(len(vertices) - 1)

        # 隔板前侧面（面向前侧格子）
        triangles.append([partition_indices[0], partition_indices[1], partition_indices[5]])
        triangles.append([partition_indices[0], partition_indices[5], partition_indices[4]])

        # 隔板后侧面（面向后侧格子）
        triangles.append([partition_indices[2], partition_indices[3], partition_indices[7]])
        triangles.append([partition_indices[2], partition_indices[7], partition_indices[6]])

        # 隔板左侧面（厚度方向）
        triangles.append([partition_indices[3], partition_indices[0], partition_indices[4]])
        triangles.append([partition_indices[3], partition_indices[4], partition_indices[7]])

        # 隔板右侧面（厚度方向）
        triangles.append([partition_indices[1], partition_indices[2], partition_indices[6]])
        triangles.append([partition_indices[1], partition_indices[6], partition_indices[5]])

        # 隔板底面
        triangles.append([partition_indices[0], partition_indices[3], partition_indices[2]])
        triangles.append([partition_indices[0], partition_indices[2], partition_indices[1]])

        # 隔板顶面
        triangles.append([partition_indices[4], partition_indices[5], partition_indices[6]])



    print(f"\n✅ 空心盒子创建完成:")
    print(f"  顶点: {len(vertices)}")
    print(f"  三角形: {len(triangles)}")

    return vertices, triangles, (total_outer_width, total_outer_depth, total_height)


def create_simple_hollow_lid(box_width, box_depth):
    """创建简单的空心盖子"""
    print(f"\n正在创建空心盖子...")

    # 盖子参数
    LID_HEIGHT = 40.0
    LID_TOP_THICKNESS = 3.0
    LID_SKIRT_HEIGHT = 15.0
    TOLERANCE = 0.5
    SKIRT_THICKNESS = 2.0

    # 尺寸
    lid_outer_width = box_width + 2 * TOLERANCE
    lid_outer_depth = box_depth + 2 * TOLERANCE

    print(f"🔶 盖子设计:")
    print(f"  外部: {lid_outer_width:.1f}x{lid_outer_depth:.1f}x{LID_HEIGHT}mm")
    print(f"  裙边厚度: {SKIRT_THICKNESS}mm")

    vertices = []
    triangles = []

    # ========== 1. 顶板（3mm厚） ==========
    # 顶板底部
    v0 = [TOLERANCE, TOLERANCE, LID_HEIGHT - LID_TOP_THICKNESS]
    v1 = [TOLERANCE + lid_outer_width, TOLERANCE, LID_HEIGHT - LID_TOP_THICKNESS]
    v2 = [TOLERANCE + lid_outer_width, TOLERANCE + lid_outer_depth, LID_HEIGHT - LID_TOP_THICKNESS]
    v3 = [TOLERANCE, TOLERANCE + lid_outer_depth, LID_HEIGHT - LID_TOP_THICKNESS]

    # 顶板顶部
    v4 = [TOLERANCE, TOLERANCE, LID_HEIGHT]
    v5 = [TOLERANCE + lid_outer_width, TOLERANCE, LID_HEIGHT]
    v6 = [TOLERANCE + lid_outer_width, TOLERANCE + lid_outer_depth, LID_HEIGHT]
    v7 = [TOLERANCE, TOLERANCE + lid_outer_depth, LID_HEIGHT]

    top_indices = []
    for v in [v0, v1, v2, v3, v4, v5, v6, v7]:
        vertices.append(v)
        top_indices.append(len(vertices) - 1)

    # 顶板顶部（外表面）
    triangles.append([top_indices[4], top_indices[5], top_indices[6]])
    triangles.append([top_indices[4], top_indices[6], top_indices[7]])

    # 顶板侧面
    triangles.append([top_indices[0], top_indices[1], top_indices[5]])
    triangles.append([top_indices[0], top_indices[5], top_indices[4]])
    triangles.append([top_indices[1], top_indices[2], top_indices[6]])
    triangles.append([top_indices[1], top_indices[6], top_indices[5]])
    triangles.append([top_indices[2], top_indices[3], top_indices[7]])
    triangles.append([top_indices[2], top_indices[7], top_indices[6]])
    triangles.append([top_indices[3], top_indices[0], top_indices[4]])
    triangles.append([top_indices[3], top_indices[4], top_indices[7]])

    # 顶板底部（内表面）
    triangles.append([top_indices[3], top_indices[2], top_indices[1]])
    triangles.append([top_indices[3], top_indices[1], top_indices[0]])

    # ========== 2. 裙边（2mm厚，底部开放） ==========
    skirt_z_start = LID_HEIGHT - LID_TOP_THICKNESS - LID_SKIRT_HEIGHT

    # 外裙边底部
    v8 = [TOLERANCE, TOLERANCE, skirt_z_start]
    v9 = [TOLERANCE + lid_outer_width, TOLERANCE, skirt_z_start]
    v10 = [TOLERANCE + lid_outer_width, TOLERANCE + lid_outer_depth, skirt_z_start]
    v11 = [TOLERANCE, TOLERANCE + lid_outer_depth, skirt_z_start]

    # 外裙边顶部（连接到顶板底部）
    # 注意：这里应该使用不同的坐标，而不是重用顶部顶点
    # 裙边顶部向外偏移（形成厚度）
    skirt_top_z = LID_HEIGHT - LID_TOP_THICKNESS

    # 外裙边外圈顶部
    v12 = [TOLERANCE, TOLERANCE, skirt_top_z]
    v13 = [TOLERANCE + lid_outer_width, TOLERANCE, skirt_top_z]
    v14 = [TOLERANCE + lid_outer_width, TOLERANCE + lid_outer_depth, skirt_top_z]
    v15 = [TOLERANCE, TOLERANCE + lid_outer_depth, skirt_top_z]

    # 外裙边内圈顶部（内壁）
    v16 = [TOLERANCE + SKIRT_THICKNESS, TOLERANCE + SKIRT_THICKNESS, skirt_top_z]
    v17 = [TOLERANCE + lid_outer_width - SKIRT_THICKNESS, TOLERANCE + SKIRT_THICKNESS, skirt_top_z]
    v18 = [TOLERANCE + lid_outer_width - SKIRT_THICKNESS, TOLERANCE + lid_outer_depth - SKIRT_THICKNESS, skirt_top_z]
    v19 = [TOLERANCE + SKIRT_THICKNESS, TOLERANCE + lid_outer_depth - SKIRT_THICKNESS, skirt_top_z]

    # 外裙边内圈底部
    v20 = [TOLERANCE + SKIRT_THICKNESS, TOLERANCE + SKIRT_THICKNESS, skirt_z_start]
    v21 = [TOLERANCE + lid_outer_width - SKIRT_THICKNESS, TOLERANCE + SKIRT_THICKNESS, skirt_z_start]
    v22 = [TOLERANCE + lid_outer_width - SKIRT_THICKNESS, TOLERANCE + lid_outer_depth - SKIRT_THICKNESS, skirt_z_start]
    v23 = [TOLERANCE + SKIRT_THICKNESS, TOLERANCE + lid_outer_depth - SKIRT_THICKNESS, skirt_z_start]

    # 所有裙边顶点
    skirt_indices = []
    for v in [v8, v9, v10, v11,  # 外圈底部
              v12, v13, v14, v15,  # 外圈顶部
              v16, v17, v18, v19,  # 内圈顶部
              v20, v21, v22, v23]:  # 内圈底部
        vertices.append(v)
        skirt_indices.append(len(vertices) - 1)

    # 外裙边外侧面（使用索引 0-3 为底部，4-7 为顶部）
    # 前侧面（y最小）
    triangles.append([skirt_indices[0], skirt_indices[1], skirt_indices[5]])
    triangles.append([skirt_indices[0], skirt_indices[5], skirt_indices[4]])
    # 右侧面（x最大）
    triangles.append([skirt_indices[1], skirt_indices[2], skirt_indices[6]])
    triangles.append([skirt_indices[1], skirt_indices[6], skirt_indices[5]])
    # 后侧面（y最大）
    triangles.append([skirt_indices[2], skirt_indices[3], skirt_indices[7]])
    triangles.append([skirt_indices[2], skirt_indices[7], skirt_indices[6]])
    # 左侧面（x最小）
    triangles.append([skirt_indices[3], skirt_indices[0], skirt_indices[4]])
    triangles.append([skirt_indices[3], skirt_indices[4], skirt_indices[7]])

    # 外裙边内侧面（使用索引 12-15 为内圈底部，8-11 为内圈顶部）
    # 前侧面
    triangles.append([skirt_indices[13], skirt_indices[12], skirt_indices[8]])
    triangles.append([skirt_indices[13], skirt_indices[8], skirt_indices[9]])
    # 右侧面
    triangles.append([skirt_indices[14], skirt_indices[13], skirt_indices[9]])
    triangles.append([skirt_indices[14], skirt_indices[9], skirt_indices[10]])
    # 后侧面
    triangles.append([skirt_indices[15], skirt_indices[14], skirt_indices[10]])
    triangles.append([skirt_indices[15], skirt_indices[10], skirt_indices[11]])
    # 左侧面
    triangles.append([skirt_indices[12], skirt_indices[15], skirt_indices[11]])
    triangles.append([skirt_indices[12], skirt_indices[11], skirt_indices[8]])

    # 外裙边顶部（连接外圈和内圈顶部）
    # 前侧顶部
    triangles.append([skirt_indices[4], skirt_indices[5], skirt_indices[9]])
    triangles.append([skirt_indices[4], skirt_indices[9], skirt_indices[8]])
    # 右侧顶部
    triangles.append([skirt_indices[5], skirt_indices[6], skirt_indices[10]])
    triangles.append([skirt_indices[5], skirt_indices[10], skirt_indices[9]])
    # 后侧顶部
    triangles.append([skirt_indices[6], skirt_indices[7], skirt_indices[11]])
    triangles.append([skirt_indices[6], skirt_indices[11], skirt_indices[10]])
    # 左侧顶部
    triangles.append([skirt_indices[7], skirt_indices[4], skirt_indices[8]])
    triangles.append([skirt_indices[7], skirt_indices[8], skirt_indices[11]])

    print(f"✅ 空心盖子创建完成:")
    print(f"  顶点: {len(vertices)}")
    print(f"  三角形: {len(triangles)}")

    return vertices, triangles

def create_3mf_true_hollow(box_vertices, box_triangles, lid_vertices, lid_triangles):
    """创建真正的空心3MF文件"""
    print("\n" + "=" * 70)
    print("正在生成真正的空心3MF文件...")

    filename = "storage_box_true_hollow.3mf"

    try:
        with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 必需的文件
            content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="model" ContentType="application/vnd.ms-package.3dmanufacturing-3dmodel+xml"/>
</Types>"""
            zipf.writestr("[Content_Types].xml", content_types)

            rels_content = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/relationships/2006">
  <Relationship Id="rel0" Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel" Target="/3D/3dmodel.model"/>
</Relationships>"""
            zipf.writestr("_rels/.rels", rels_content)

            # 模型文件
            model_xml = create_true_hollow_xml(box_vertices, box_triangles, lid_vertices, lid_triangles)
            zipf.writestr("3D/3dmodel.model", model_xml)

            model_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/relationships/2006">
</Relationships>"""
            zipf.writestr("3D/_rels/3dmodel.model.rels", model_rels)

            file_size = os.path.getsize(filename) if os.path.exists(filename) else 0
            print(f"✅ 3MF文件已创建: {filename}")
            print(f"📦 文件大小: {file_size / 1024:.1f} KB")

            if file_size > 5000:
                print("✅ 文件包含足够的数据")
            else:
                print("⚠️  文件可能过小")

        return True

    except Exception as e:
        print(f"❌ 创建文件失败: {e}")
        return False

def create_true_hollow_xml(box_vertices, box_triangles, lid_vertices, lid_triangles):
    """创建XML内容"""

    xml = []
    xml.append('<?xml version="1.0" encoding="UTF-8"?>')
    xml.append('<model unit="millimeter" xml:lang="en-US" xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02">')

    xml.append('  <metadata name="Application">True Hollow Box Generator</metadata>')
    xml.append(f'  <metadata name="CreationDate">{datetime.now().isoformat()}</metadata>')
    xml.append('  <metadata name="Title">储物盒（真正空心）</metadata>')
    xml.append('  <metadata name="Description">60x60x200mm格子，真正空心结构，壁厚2mm</metadata>')

    xml.append('  <resources>')

    # 盒子
    xml.append('    <object id="1" type="model">')
    xml.append('      <mesh>')
    xml.append('        <vertices>')

    for v in box_vertices:
        xml.append(f'          <vertex x="{v[0]:.6f}" y="{v[1]:.6f}" z="{v[2]:.6f}" />')

    xml.append('        </vertices>')
    xml.append('        <triangles>')

    for t in box_triangles:
        if len(t) == 3:
            xml.append(f'          <triangle v1="{t[0]}" v2="{t[1]}" v3="{t[2]}" />')

    xml.append('        </triangles>')
    xml.append('      </mesh>')
    xml.append('    </object>')

    # 盖子
    xml.append('    <object id="2" type="model">')
    xml.append('      <mesh>')
    xml.append('        <vertices>')

    for v in lid_vertices:
        xml.append(f'          <vertex x="{v[0]:.6f}" y="{v[1]:.6f}" z="{v[2]:.6f}" />')

    xml.append('        </vertices>')
    xml.append('        <triangles>')

    for t in lid_triangles:
        if len(t) == 3:
            xml.append(f'          <triangle v1="{t[0]}" v2="{t[1]}" v3="{t[2]}" />')

    xml.append('        </triangles>')
    xml.append('      </mesh>')
    xml.append('    </object>')

    xml.append('  </resources>')

    xml.append('  <build>')
    xml.append('    <item objectid="1" />')
    xml.append('    <item objectid="2" />')
    xml.append('  </build>')

    xml.append('</model>')

    return '\n'.join(xml)

def main():
    """主函数"""
    print("=" * 70)
    print("真正空心储物盒生成器")
    print("🎯 解决实心问题：创建只有2mm厚壁的空心结构")
    print("=" * 70)

    try:
        # 1. 创建真正的空心盒子
        print("\n1. 创建盒子（真正空心）...")
        box_vertices, box_triangles, box_dimensions = create_true_hollow_box()

        print(f"\n📊 盒子统计:")
        print(f"  尺寸: {box_dimensions[0]:.1f}x{box_dimensions[1]:.1f}x{box_dimensions[2]:.1f}mm")
        print(f"  顶点数: {len(box_vertices)}")
        print(f"  三角形数: {len(box_triangles)}")

        # 估算：外表面 + 内表面 + 隔板
        expected_triangles = 8*2 + 8*2 + 5*12  # 外表面16 + 内表面16 + 5个隔板×12 = 92
        print(f"  预期三角形: ~{expected_triangles}个")

        # 2. 创建空心盖子
        print("\n2. 创建盖子（空心）...")
        lid_vertices, lid_triangles = create_simple_hollow_lid(box_dimensions[0], box_dimensions[1])

        # 3. 生成3MF文件
        print("\n3. 生成3MF文件...")
        success = create_3mf_true_hollow(box_vertices, box_triangles, lid_vertices, lid_triangles)

        if success:
            print("\n" + "=" * 70)
            print("✅ 生成成功!")
            print(f"   文件: storage_box_true_hollow.3mf")

            print("\n🔍 关键特点:")
            print("   1. ✅ 真正空心结构（不是实心块）")
            print("   2. ✅ 壁厚: 2mm")
            print("   3. ✅ 隔板: 3条纵向 + 2条横向")
            print("   4. ✅ 顶部开放（盒子）")
            print("   5. ✅ 底部开放（盖子）")

            print("\n🖨️  切片验证:")
            print("   1. 打开storage_box_true_hollow.3mf")
            print("   2. 切换到X射线视图（Cura: 按TAB）")
            print("   3. 应该看到:")
            print("      - 外壁: 2mm厚")
            print("      - 内壁: 与外壁平行，间隔2mm")
            print("      - 隔板: 连接内外壁")
            print("      - 内部: 完全空心")

            print("\n⚠️  如果仍然显示实心:")
            print("   1. 检查切片软件的显示设置")
            print("   2. 尝试其他切片软件（PrusaSlicer、Simplify3D）")
            print("   3. 文件大小应大于10KB")
        else:
            print("❌ 文件生成失败")

        print("=" * 70)

        # 4. 创建简单的STL作为备用
        print("\n🔧 创建备用STL文件...")
        try:
            import struct

            def write_simple_stl(vertices, triangles, filename):
                with open(filename, 'wb') as f:
                    # 头文件
                    f.write(b'Simple STL Backup' + b' ' * 62)
                    # 三角形数量
                    f.write(struct.pack('<I', len(triangles)))
                    # 三角形数据
                    for t in triangles:
                        if len(t) != 3:
                            continue
                        v0, v1, v2 = vertices[t[0]], vertices[t[1]], vertices[t[2]]
                        # 法向量
                        f.write(struct.pack('<fff', 0, 0, 1))
                        # 顶点
                        for v in [v0, v1, v2]:
                            f.write(struct.pack('<fff', v[0], v[1], v[2]))
                        # 属性
                        f.write(struct.pack('<H', 0))

            write_simple_stl(box_vertices, box_triangles, "box_backup.stl")
            write_simple_stl(lid_vertices, lid_triangles, "lid_backup.stl")
            print("✅ 备用STL文件已创建")

        except:
            print("⚠️  备用STL创建失败")

    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()