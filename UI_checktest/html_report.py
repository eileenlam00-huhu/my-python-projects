"""
HTML 报告生成器 - 实时刷新版本

这个模块生成实时HTML测试报告，包含：
- 实时数据刷新（通过JSON文件轮询）
- 完整的截图显示和对比
- 与Excel报告一致的数据结构
- 实时进度更新
"""

import os
import json
import base64
from datetime import datetime
from .utils import ensure_dir_exists, logger


class HTMLReportGenerator:
    """HTML报告生成器 - 实时刷新版本"""

    def __init__(self, report_dir, run_id):
        self.report_dir = report_dir
        self.run_id = run_id
        self.html_path = os.path.join(report_dir, f'UI_Report_{run_id}.html')
        self.json_path = os.path.join(report_dir, f'UI_Report_{run_id}.json')
        
        # 数据结构与Excel报告完全一致
        self.data = {
            'run_id': run_id,
            'start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'end_time': None,
            'tests': [],  # 与Excel中的行数据一致
            'statistics': {
                'total_tests': 0,
                'passed_tests': 0,
                'failed_tests': 0,
                'total_languages': 0,
                'completed_languages': 0
            },
            'status': 'running'
        }
        
        ensure_dir_exists(report_dir)
        self._save_json()
        self._create_html_template()

    def _image_to_base64(self, image_path):
        """将图片转换为base64编码"""
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

    def _save_json(self):
        """保存JSON数据文件"""
        try:
            with open(self.json_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger(f'[HTML] 保存JSON失败: {e}')

    def _create_html_template(self):
        """创建HTML模板文件"""
        html_content = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>UI自动化测试报告 - 实时更新</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f5f5;
            color: #333;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            opacity: 0.9;
            font-size: 1em;
        }
        
        .refresh-indicator {
            margin-top: 15px;
            padding: 10px;
            background: rgba(255,255,255,0.2);
            border-radius: 4px;
            font-size: 0.9em;
        }
        
        .refresh-indicator.updating {
            animation: pulse 1s infinite;
        }
        
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.6; }
        }
        
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            text-align: center;
        }
        
        .stat-card h3 {
            color: #666;
            font-size: 0.9em;
            text-transform: uppercase;
            margin-bottom: 10px;
            letter-spacing: 1px;
        }
        
        .stat-card .number {
            font-size: 2.5em;
            font-weight: bold;
            color: #007bff;
        }
        
        .progress-section {
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .progress-bar {
            width: 100%;
            height: 30px;
            background: #e9ecef;
            border-radius: 15px;
            overflow: hidden;
            margin: 10px 0;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #28a745, #20c997);
            transition: width 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 0.9em;
        }
        
        .tests-container {
            display: grid;
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .test-item {
            background: white;
            border-left: 4px solid #007bff;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .test-item.pass {
            border-left-color: #28a745;
        }
        
        .test-item.fail {
            border-left-color: #dc3545;
        }
        
        .test-item.warn {
            border-left-color: #ffc107;
        }
        
        .test-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        .test-title {
            font-size: 1.1em;
            font-weight: 600;
            color: #333;
            flex: 1;
        }
        
        .test-type {
            display: inline-block;
            padding: 4px 10px;
            background: #e9ecef;
            border-radius: 4px;
            font-size: 0.8em;
            font-weight: 600;
            color: #495057;
        }
        
        .test-status {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 4px;
            font-size: 0.8em;
            font-weight: 600;
            text-transform: uppercase;
        }
        
        .test-status.pass {
            background: #d4edda;
            color: #155724;
        }
        
        .test-status.fail {
            background: #f8d7da;
            color: #721c24;
        }
        
        .test-status.warn {
            background: #fff3cd;
            color: #856404;
        }
        
        .test-details {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 12px;
            margin-bottom: 12px;
        }
        
        .detail-item {
            background: #f8f9fa;
            padding: 10px;
            border-radius: 4px;
        }
        
        .detail-label {
            font-size: 0.8em;
            color: #666;
            font-weight: 600;
            text-transform: uppercase;
            margin-bottom: 4px;
        }
        
        .detail-value {
            font-size: 0.95em;
            color: #333;
            word-break: break-all;
        }
        
        .image-comparison {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-top: 12px;
        }
        
        @media (max-width: 768px) {
            .image-comparison {
                grid-template-columns: 1fr;
            }
        }
        
        .image-item {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 4px;
            overflow: hidden;
        }
        
        .image-label {
            padding: 8px 10px;
            font-size: 0.85em;
            font-weight: 600;
            color: #495057;
            border-bottom: 1px solid #dee2e6;
            background: white;
        }
        
        .image-content {
            padding: 8px;
            text-align: center;
        }
        
        .image-content img {
            max-width: 100%;
            height: auto;
            border-radius: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .image-content.no-image {
            padding: 20px;
            color: #999;
            font-size: 0.9em;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>UI自动化测试报告</h1>
            <p>测试运行ID: <span id="runId"></span></p>
            <p>开始时间: <span id="startTime"></span></p>
            <div class="refresh-indicator">
                <span id="refreshStatus">🔄 正在加载数据...</span>
            </div>
        </div>
        
        <div class="stats" id="statsContainer"></div>
        
        <div class="progress-section">
            <h2 style="margin-bottom: 10px; font-size: 1.2em; color: #333;">测试进度</h2>
            <div class="progress-bar">
                <div class="progress-fill" id="progressBar" style="width: 0%">0%</div>
            </div>
            <p id="progressText" style="color: #666; font-size: 0.9em; margin-top: 8px;"></p>
        </div>
        
        <div>
            <h2 style="margin-bottom: 15px; font-size: 1.3em; color: #333;">测试结果</h2>
            <div class="tests-container" id="testsContainer"></div>
        </div>
        
        <div class="footer">
            <p>最后更新: <span id="lastUpdate"></span></p>
            <p>UI自动化测试系统</p>
        </div>
    </div>
    
    <script>
        const runId = '__RUN_ID__';
        const JSON_FILE = '__JSON_FILE__';
        let lastData = null;
        let updateCount = 0;

        async function loadData() {
            try {
                const response = await fetch(JSON_FILE + '?t=' + Date.now());
                if (!response.ok) throw new Error('Failed to load JSON');
                const data = await response.json();
                renderReport(data);
                lastData = data;
                updateCount++;
                updateRefreshStatus();
            } catch (error) {
                console.error('Error loading data:', error);
                document.getElementById('refreshStatus').textContent = '❌ 数据加载失败，重试中...';
            }
        }

        function updateRefreshStatus() {
            const status = document.getElementById('refreshStatus');
            if (lastData && lastData.status === 'completed') {
                status.textContent = '✅ 测试已完成';
            } else {
                status.innerHTML = '🔄 实时更新中 (第 ' + updateCount + ' 次更新)';
            }
        }

        function renderReport(data) {
            document.getElementById('runId').textContent = data.run_id;
            document.getElementById('startTime').textContent = data.start_time;
            document.getElementById('lastUpdate').textContent = new Date().toLocaleString('zh-CN');

            const stats = data.statistics;
            const statsHtml = `
                <div class="stat-card">
                    <h3>总测试数</h3>
                    <div class="number">${stats.total_tests}</div>
                </div>
                <div class="stat-card">
                    <h3>通过</h3>
                    <div class="number" style="color: #28a745;">${stats.passed_tests}</div>
                </div>
                <div class="stat-card">
                    <h3>失败</h3>
                    <div class="number" style="color: #dc3545;">${stats.failed_tests}</div>
                </div>
                <div class="stat-card">
                    <h3>语言进度</h3>
                    <div class="number">${stats.completed_languages}/${stats.total_languages}</div>
                </div>
            `;
            document.getElementById('statsContainer').innerHTML = statsHtml;

            const completedLanguages = stats.completed_languages || 0;
            const totalLanguages = stats.total_languages || 1;
            const progressPercent = Math.round((completedLanguages / totalLanguages) * 100);
            const progressBar = document.getElementById('progressBar');
            progressBar.style.width = progressPercent + '%';
            progressBar.textContent = progressPercent + '%';
            document.getElementById('progressText').textContent = 
                `已完成语言: ${completedLanguages}/${totalLanguages} (${progressPercent}%)`;

            const testsHtml = data.tests.map(test => renderTestItem(test)).join('');
            document.getElementById('testsContainer').innerHTML = testsHtml || '<p style="text-align: center; color: #999;">暂无测试数据</p>';
        }

        function renderTestItem(test) {
            const statusRaw = test.status ? test.status.toUpperCase() : 'UNKNOWN';
            let statusClass = 'unknown';
            if (statusRaw.includes('PASS')) {
                statusClass = 'pass';
            } else if (statusRaw.includes('WARN')) {
                statusClass = 'warn';
            } else if (statusRaw.includes('FAIL')) {
                statusClass = 'fail';
            }
            const diffRatio = test.diff_ratio >= 0 ? test.diff_ratio.toFixed(4) : 'N/A';

            let imagesHtml = '';
            if (test.image1 || test.image2) {
                imagesHtml = '<div class="image-comparison">';

                if (test.image1) {
                    imagesHtml += `
                        <div class="image-item">
                            <div class="image-label">${test.image_label_1 || '基准/切换前'}</div>
                            <div class="image-content">
                                <img src="${test.image1}" alt="Image 1" loading="lazy">
                            </div>
                        </div>
                    `;
                }

                if (test.image2) {
                    imagesHtml += `
                        <div class="image-item">
                            <div class="image-label">${test.image_label_2 || '测试/切换后'}</div>
                            <div class="image-content">
                                <img src="${test.image2}" alt="Image 2" loading="lazy">
                            </div>
                        </div>
                    `;
                }

                imagesHtml += '</div>';
            }

            return `
                <div class="test-item ${statusClass}">
                    <div class="test-header">
                        <div class="test-title">${test.language_id} - ${test.language} - ${test.main_page} - ${test.sub_page}</div>
                        <span class="test-type">${test.type}</span>
                        <span class="test-status ${statusClass}">${test.status}</span>
                    </div>
                    <div class="test-details">
                        <div class="detail-item">
                            <div class="detail-label">时间</div>
                            <div class="detail-value">${test.timestamp || '-'}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">差异值</div>
                            <div class="detail-value">${diffRatio}</div>
                        </div>
                    </div>
                    ${imagesHtml}
                </div>
            `;
        }

        loadData();
        setInterval(loadData, 5000);
    </script>
</body>
</html>
'''


        try:
            html_content = html_content.replace('__RUN_ID__', self.run_id)
            html_content = html_content.replace('__JSON_FILE__', f'UI_Report_{self.run_id}.json')
            with open(self.html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            logger(f'[HTML] HTML模板已创建: {self.html_path}')
        except Exception as e:
            logger(f'[HTML] 创建HTML模板失败: {e}')

    def add_test_result(self, lang_id, lang_name, test_type, main_page, sub_page, 
                       image1_path=None, image2_path=None, diff_ratio=-1, status='UNKNOWN',
                       image_label_1=None, image_label_2=None):
        """添加测试结果 - 与Excel数据结构保持一致"""
        
        test_data = {
            'language_id': lang_id,
            'language': lang_name,
            'type': test_type,  # 'LANG' 或 'UI'
            'main_page': main_page,
            'sub_page': sub_page,
            'image1': self._image_to_base64(image1_path) if image1_path else None,
            'image2': self._image_to_base64(image2_path) if image2_path else None,
            'image_label_1': image_label_1,
            'image_label_2': image_label_2,
            'diff_ratio': diff_ratio,
            'status': status,
            'timestamp': datetime.now().strftime('%H:%M:%S')
        }
        
        self.data['tests'].append(test_data)
        
        # 更新统计信息
        self.data['statistics']['total_tests'] = len(self.data['tests'])
        status_norm = str(status).upper().strip() if status else ''
        if 'PASS' in status_norm:
            self.data['statistics']['passed_tests'] = self.data['statistics'].get('passed_tests', 0) + 1
        elif 'FAIL' in status_norm or 'WARN' in status_norm:
            self.data['statistics']['failed_tests'] = self.data['statistics'].get('failed_tests', 0) + 1
        
        self._save_json()

    def add_language_test(self, lang_id, lang_name, status, diff_ratio=-1, 
                         image1_path=None, image2_path=None):
        """添加语言测试结果"""
        self.add_test_result(
            lang_id=lang_id,
            lang_name=lang_name,
            test_type='LANG',
            main_page='语言设置',
            sub_page='切换验证',
            image1_path=image1_path,
            image2_path=image2_path,
            diff_ratio=diff_ratio,
            status=status,
            image_label_1='切换前',
            image_label_2='切换后'
        )
        # 更新语言完成数
        unique_langs = len(set([t['language_id'] for t in self.data['tests'] if t['type'] == 'LANG']))
        self.data['statistics']['completed_languages'] = unique_langs

    def add_ui_test(self, lang_name, main_page, sub_page, status, diff_ratio=-1, 
                   base_image_path=None, test_image_path=None):
        """添加UI测试结果"""
        lang_id = None
        for test in reversed(self.data['tests']):
            if test['type'] == 'LANG' and test['language'] == lang_name:
                lang_id = test['language_id']
                break
        
        if lang_id is None:
            lang_id = 0
        
        self.add_test_result(
            lang_id=lang_id,
            lang_name=lang_name,
            test_type='UI',
            main_page=main_page,
            sub_page=sub_page,
            image1_path=base_image_path,
            image2_path=test_image_path,
            diff_ratio=diff_ratio,
            status=status,
            image_label_1='基准图',
            image_label_2='测试图'
        )

    def set_total_languages(self, count):
        """设置总语言数量"""
        self.data['statistics']['total_languages'] = count
        self._save_json()

    def mark_completed(self):
        """标记测试完成"""
        self.data['status'] = 'completed'
        self.data['end_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self._save_json()
        logger(f'[HTML] 报告已完成: {self.html_path}')


def create_html_report(report_dir, run_id):
    """创建HTML报告生成器"""
    return HTMLReportGenerator(report_dir, run_id)
