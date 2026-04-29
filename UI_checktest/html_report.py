"""
HTML 报告生成器

这个模块生成实时HTML测试报告，包含：
- 测试进度和状态
- 语言切换结果
- UI对比结果
- 图片预览（基准图、测试图、Diff图）
- 统计信息

HTML报告支持实时更新，方便在测试过程中查看进度。
"""

import os
import base64
import json
from datetime import datetime
from .utils import ensure_dir_exists, logger


class HTMLReportGenerator:
    """HTML报告生成器"""

    def __init__(self, report_dir, run_id):
        self.report_dir = report_dir
        self.run_id = run_id
        self.html_path = os.path.join(report_dir, f'UI_Report_{run_id}.html')
        self.data = {
            'run_id': run_id,
            'start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'languages': [],
            'ui_tests': [],
            'statistics': {
                'total_languages': 0,
                'completed_languages': 0,
                'total_ui_tests': 0,
                'passed_ui_tests': 0,
                'failed_ui_tests': 0,
                'warning_ui_tests': 0
            },
            'status': 'running'
        }
        ensure_dir_exists(report_dir)
        self._create_initial_html()

    def _create_initial_html(self):
        """创建初始HTML文件"""
        html_content = self._generate_html()
        with open(self.html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        logger(f'[HTML] 报告已创建: {self.html_path}')

    def _image_to_base64(self, image_path):
        """将图片转换为base64编码，如果失败返回None"""
        if not image_path or not os.path.exists(image_path):
            return None

        try:
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()
                img_ext = os.path.splitext(image_path)[1].lower()
                if img_ext == '.png':
                    img_type = 'png'
                elif img_ext in ['.jpg', '.jpeg']:
                    img_type = 'jpeg'
                else:
                    return None
                return f'data:image/{img_type};base64,{base64.b64encode(img_data).decode()}'
        except Exception as e:
            logger(f'[HTML] 图片编码失败: {image_path} - {e}')
            return None

    def _generate_html(self):
        """生成HTML内容"""
        return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>UI自动化测试报告 - {self.run_id}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }}
        .header p {{
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 1.1em;
        }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }}
        .stat-card {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            text-align: center;
        }}
        .stat-card h3 {{
            margin: 0 0 10px 0;
            color: #495057;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}
        .stat-card .number {{
            font-size: 2em;
            font-weight: bold;
            color: #007bff;
        }}
        .status-running {{
            color: #ffc107;
        }}
        .status-completed {{
            color: #28a745;
        }}
        .content {{
            padding: 30px;
        }}
        .section {{
            margin-bottom: 40px;
        }}
        .section h2 {{
            color: #495057;
            border-bottom: 2px solid #007bff;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        .language-list, .ui-test-list {{
            display: grid;
            gap: 15px;
        }}
        .test-item {{
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: 20px;
            transition: all 0.3s ease;
        }}
        .test-item:hover {{
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            transform: translateY(-2px);
        }}
        .test-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }}
        .test-title {{
            font-size: 1.2em;
            font-weight: 600;
            color: #495057;
            margin-right: 15px;
        }}
        .test-info {{
            display: flex;
            gap: 15px;
            font-size: 0.9em;
            color: #6c757d;
            align-items: center;
        }}
        .test-status {{
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.8em;
            font-weight: 600;
            text-transform: uppercase;
        }}
        .status-pass {{
            background: #d4edda;
            color: #155724;
        }}
        .status-fail {{
            background: #f8d7da;
            color: #721c24;
        }}
        .status-warn {{
            background: #fff3cd;
            color: #856404;
        }}
        .status-unknown {{
            background: #e2e3e5;
            color: #383d41;
        }}
        .test-details {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 15px;
        }}
        .detail-item {{
            background: white;
            padding: 10px;
            border-radius: 4px;
            border: 1px solid #e9ecef;
        }}
        .detail-label {{
            font-weight: 600;
            color: #6c757d;
            font-size: 0.9em;
            margin-bottom: 5px;
        }}
        .detail-value {{
            color: #495057;
        }}
        .image-comparison {{
            margin-top: 20px;
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
        }}
        .image-item {{
            text-align: center;
            background: white;
            padding: 10px;
            border-radius: 4px;
            border: 1px solid #e9ecef;
        }}
        .image-item img {{
            max-width: 100%;
            height: auto;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        .image-placeholder {{
            width: 100%;
            height: 150px;
            background-color: #e9ecef;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #6c757d;
            font-size: 0.9em;
            border-radius: 4px;
        }}
        .image-label {{
            margin-top: 10px;
            font-weight: 600;
            color: #6c757d;
            font-size: 0.9em;
        }}
        .progress-bar {{
            width: 100%;
            height: 20px;
            background: #e9ecef;
            border-radius: 10px;
            overflow: hidden;
            margin: 10px 0;
        }}
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, #007bff, #0056b3);
            transition: width 0.3s ease;
        }}
        .footer {{
            text-align: center;
            padding: 20px;
            background: #f8f9fa;
            color: #6c757d;
            border-top: 1px solid #e9ecef;
        }}
        .refresh-info {{
            background: #e3f2fd;
            border: 1px solid #2196f3;
            border-radius: 4px;
            padding: 10px;
            margin: 20px 0;
            text-align: center;
            color: #1565c0;
        }}
        @media (max-width: 992px) {{
            .image-comparison {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
    <script>
        function autoRefresh() {{
            setTimeout(function() {{
                location.reload();
            }}, 5000); // 5秒自动刷新
        }}

        // 只有在测试进行中时才自动刷新
        {f"window.onload = autoRefresh;" if self.data['status'] == 'running' else ""}
    </script>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>UI自动化测试报告</h1>
            <p>测试运行ID: {self.run_id} | 开始时间: {self.data['start_time']}</p>
        </div>

        {f'<div class="refresh-info">🔄 报告正在实时更新中... (每5秒自动刷新)</div>' if self.data['status'] == 'running' else '<div class="refresh-info">✅ 测试已完成</div>'}

        <div class="stats">
            <div class="stat-card">
                <h3>语言测试</h3>
                <div class="number">{self.data['statistics']['completed_languages']}/{self.data['statistics']['total_languages']}</div>
            </div>
            <div class="stat-card">
                <h3>UI测试</h3>
                <div class="number">{self.data['statistics']['total_ui_tests']}</div>
            </div>
            <div class="stat-card">
                <h3>通过</h3>
                <div class="number status-completed">{self.data['statistics']['passed_ui_tests']}</div>
            </div>
            <div class="stat-card">
                <h3>失败</h3>
                <div class="number" style="color: #dc3545;">{self.data['statistics']['failed_ui_tests']}</div>
            </div>
        </div>

        <div class="content">
            <div class="section">
                <h2>📊 测试进度</h2>
                <div class="progress-bar">
                    <div class="progress-fill" style="width: {self._calculate_progress()}%"></div>
                </div>
                <p>总体进度: {self._calculate_progress():.1f}%</p>
            </div>

            {self._generate_languages_section()}

            {self._generate_ui_tests_section()}
        </div>

        <div class="footer">
            <p>报告生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | UI自动化测试系统</p>
        </div>
    </div>
</body>
</html>'''

    def _calculate_progress(self):
        """计算测试进度百分比"""
        if self.data['statistics']['total_languages'] == 0:
            return 0
        return (self.data['statistics']['completed_languages'] / self.data['statistics']['total_languages']) * 100

    def _generate_languages_section(self):
        """生成语言测试部分"""
        if not self.data['languages']:
            return '<div class="section"><h2>🌐 语言测试</h2><p>暂无语言测试数据</p></div>'

        html = '<div class="section"><h2>🌐 语言测试</h2><div class="language-list">'

        for lang in self.data['languages']:
            status_class = f"status-{lang['status'].lower()}"
            html += f'''
            <div class="test-item">
                <div class="test-header">
                    <div class="test-title">{lang['name']} (ID: {lang['id']})</div>
                    <div class="test-status {status_class}">{lang['status']}</div>
                </div>
                <div class="test-details">
                    <div class="detail-item">
                        <div class="detail-label">差异值</div>
                        <div class="detail-value">{lang['diff_ratio']:.4f}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">测试时间</div>
                        <div class="detail-value">{lang['timestamp']}</div>
                    </div>
                </div>
                {self._generate_language_images(lang)}
            </div>'''

        html += '</div></div>'
        return html

    def _generate_language_images(self, lang):
        """生成语言测试的图片对比"""
        has_img1 = bool(lang.get('image1'))
        has_img2 = bool(lang.get('image2'))
        
        if not has_img1 and not has_img2:
            return ''

        html = '<div class="image-comparison">'
        
        # 图片1
        html += f'''
        <div class="image-item">
            {f'<img src="{lang["image1"]}" alt="切换前图片" loading="lazy">' if has_img1 else '<div class="image-placeholder">图片加载失败</div>'}
            <div class="image-label">切换前</div>
        </div>'''
        
        # 图片2
        html += f'''
        <div class="image-item">
            {f'<img src="{lang["image2"]}" alt="切换后图片" loading="lazy">' if has_img2 else '<div class="image-placeholder">图片加载失败</div>'}
            <div class="image-label">切换后</div>
        </div>'''
        
        html += '</div>'
        return html

    def _generate_ui_tests_section(self):
        """生成UI测试部分"""
        if not self.data['ui_tests']:
            return '<div class="section"><h2>🖥️ UI测试</h2><p>暂无UI测试数据</p></div>'

        html = '<div class="section"><h2>🖥️ UI测试</h2><div class="ui-test-list">'

        for test in self.data['ui_tests']:
            status_class = f"status-{test['status'].lower()}"
            
            # 构建标题信息，使其与 Excel 一致
            title_info = f"{test['language']} - {test['page_type']} - {test['main_page']} - {test['sub_page']}"
            
            # --- 修改点：修复 diff_ratio 的格式化逻辑 ---
            diff_val = test['diff_ratio']
            diff_str = f"{diff_val:.4f}" if diff_val >= 0 else 'N/A'
            
            html += f'''
            <div class="test-item">
                <div class="test-header">
                    <div class="test-title">{title_info}</div>
                    <div class="test-status {status_class}">{test['status']}</div>
                </div>
                <div class="test-details">
                    <div class="detail-item">
                        <div class="detail-label">差异值</div>
                        <div class="detail-value">{diff_str}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">测试时间</div>
                        <div class="detail-value">{test['timestamp']}</div>
                    </div>
                </div>
                {self._generate_ui_images(test)}
            </div>'''

        html += '</div></div>'
        return html

    def _generate_ui_images(self, test):
        """生成UI测试的图片对比 (基准图、测试图、Diff图)"""
        has_base = bool(test.get('base_image'))
        has_test = bool(test.get('test_image'))
        has_diff = bool(test.get('diff_image'))
        
        if not has_base and not has_test and not has_diff:
            return ''

        html = '<div class="image-comparison">'
        
        # 基准图
        html += f'''
        <div class="image-item">
            {f'<img src="{test["base_image"]}" alt="基准图片" loading="lazy">' if has_base else '<div class="image-placeholder">基准图缺失</div>'}
            <div class="image-label">基准图片</div>
        </div>'''
        
        # 测试图
        html += f'''
        <div class="image-item">
            {f'<img src="{test["test_image"]}" alt="测试图片" loading="lazy">' if has_test else '<div class="image-placeholder">测试图缺失</div>'}
            <div class="image-label">测试图片</div>
        </div>'''

        # Diff图
        html += f'''
        <div class="image-item">
            {f'<img src="{test["diff_image"]}" alt="差异图片" loading="lazy">' if has_diff else '<div class="image-placeholder">Diff图缺失</div>'}
            <div class="image-label">差异图片</div>
        </div>'''
        
        html += '</div>'
        return html

    def add_language_test(self, lang_id, lang_name, status, diff_ratio, image1_path=None, image2_path=None):
        """添加语言测试结果"""
        lang_data = {
            'id': lang_id,
            'name': lang_name,
            'status': status,
            'diff_ratio': diff_ratio,
            'timestamp': datetime.now().strftime('%H:%M:%S'),
            'image1': self._image_to_base64(image1_path) if image1_path else None,
            'image2': self._image_to_base64(image2_path) if image2_path else None
        }
        self.data['languages'].append(lang_data)
        self.data['statistics']['total_languages'] = len(set([l['id'] for l in self.data['languages']]))
        self.data['statistics']['completed_languages'] = len([l for l in self.data['languages'] if l['status'] in ['PASS', 'FAIL', 'WARN']])
        self._update_html()

    def add_ui_test(self, lang_name, page_type, main_page, sub_page, status, diff_ratio, base_image_path=None, test_image_path=None, diff_image_path=None):
        """
        添加UI测试结果
        :param lang_name: 语言名称
        :param page_type: 页面类型 (新增，与Excel对齐)
        :param main_page: 一级页面
        :param sub_page: 子页面
        :param status: 状态 (PASS/FAIL/WARN)
        :param diff_ratio: 差异值
        :param base_image_path: 基准图路径
        :param test_image_path: 测试图路径
        :param diff_image_path: Diff图路径 (新增)
        """
        # --- 修改点: 在添加数据前，确保 diff_ratio 是浮点数 ---
        try:
            # 尝试将 diff_ratio 转换为浮点数
            diff_ratio_float = float(diff_ratio)
        except (ValueError, TypeError):
            # 如果转换失败（例如是 None 或非数字字符串），则设为 -1
            diff_ratio_float = -1.0

        ui_data = {
            'language': lang_name,
            'page_type': page_type, # 新增字段
            'main_page': main_page,
            'sub_page': sub_page,
            'status': status,
            'diff_ratio': diff_ratio_float, # 使用转换后的浮点数
            'timestamp': datetime.now().strftime('%H:%M:%S'),
            'base_image': self._image_to_base64(base_image_path) if base_image_path else None,
            'test_image': self._image_to_base64(test_image_path) if test_image_path else None,
            'diff_image': self._image_to_base64(diff_image_path) if diff_image_path else None # 新增字段
        }
        self.data['ui_tests'].append(ui_data)

        # 更新统计
        self.data['statistics']['total_ui_tests'] = len(self.data['ui_tests'])
        if status == 'PASS':
            self.data['statistics']['passed_ui_tests'] += 1
        elif status in ['FAIL', 'FAIL_NO_IMAGE']:
            self.data['statistics']['failed_ui_tests'] += 1
        elif status == 'WARN':
            self.data['statistics']['warning_ui_tests'] += 1

        self._update_html()

    def set_total_languages(self, count):
        """设置总语言数量"""
        self.data['statistics']['total_languages'] = count
        self._update_html()

    def mark_completed(self):
        """标记测试完成"""
        self.data['status'] = 'completed'
        self._update_html()

    def _update_html(self):
        """更新HTML文件"""
        try:
            html_content = self._generate_html()
            with open(self.html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
        except Exception as e:
            logger(f'[HTML] 更新报告失败: {e}')


def create_html_report(report_dir, run_id):
    """创建HTML报告生成器"""
    return HTMLReportGenerator(report_dir, run_id)
