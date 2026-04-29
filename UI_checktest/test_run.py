"""
UI Check Test Suite

这个模块包含 UI 自动化检查的测试套件。
使用 pytest 框架测试各个模块的功能，包括：
- 基本功能测试
- 模块导入验证
- 工具函数测试
- 报告生成测试
- 依赖检查测试
- 环境验证测试

运行方式：
    python -m pytest test_run.py -v
"""

import pytest
import sys
import os
import tempfile
import shutil
import subprocess
import importlib

# 添加项目根目录到路径，以便导入 UI_checktest 包
# 这允许测试脚本导入项目中的其他模块
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# 必需的依赖包列表（带导入名称和显示名称）
REQUIRED_PACKAGES = [
    ('cv2', 'cv2'),         # OpenCV
    ('numpy', 'numpy'),     # NumPy
    ('openpyxl', 'openpyxl'), # Excel 处理
    ('pytest', 'pytest'),   # 测试框架
]

# 可选的依赖包列表（如果可用会增强功能）
OPTIONAL_PACKAGES = [
    ('PIL', 'PIL/Pillow'),  # 图像处理增强 (用于 Excel 图片嵌入)
    ('psutil', 'psutil'),   # 系统监控
]

# 必需的系统工具列表
REQUIRED_SYSTEM_TOOLS = [
    'adb',      # Android Debug Bridge
]

# 项目模块列表
PROJECT_MODULES = [
    'UI_checktest.constants',
    'UI_checktest.utils',
    'UI_checktest.click',
    'UI_checktest.login',
    'UI_checktest.page',
    'UI_checktest.multilanguage',
    'UI_checktest.screenshot',
    'UI_checktest.compare',
    'UI_checktest.report',
    'UI_checktest.run',
]

def test_required_packages_available():
    """测试所有必需的 Python 包是否已安装"""
    missing_packages = []

    for import_name, display_name in REQUIRED_PACKAGES:
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing_packages.append(display_name)

    if missing_packages:
        pytest.fail(f"缺少必需的 Python 包: {', '.join(missing_packages)}")

def test_optional_packages_available():
    """测试可选的 Python 包（如果安装了会增强功能）"""
    available_packages = []
    unavailable_packages = []

    for import_name, display_name in OPTIONAL_PACKAGES:
        try:
            # 特殊处理 PIL (Pillow)
            if import_name == 'PIL':
                try:
                    import PIL
                    available_packages.append(display_name)
                except ImportError:
                    # 尝试 pillow 包
                    try:
                        import pillow
                        available_packages.append(display_name)
                    except ImportError:
                        unavailable_packages.append(display_name)
            else:
                importlib.import_module(import_name)
                available_packages.append(display_name)
        except ImportError:
            unavailable_packages.append(display_name)

    # 只在有包状态变化时输出信息，避免重复日志
    if available_packages:
        print(f"可选包已安装: {', '.join(available_packages)}")
    if unavailable_packages:
        print(f"可选包未安装: {', '.join(unavailable_packages)}")

    # 这个测试总是通过，只是为了信息收集
    assert True

def test_system_tools_available():
    """测试所有必需的系统工具是否可用"""
    missing_tools = []

    for tool in REQUIRED_SYSTEM_TOOLS:
        try:
            result = subprocess.run([tool, '--version'],
                                  capture_output=True,
                                  text=True,
                                  timeout=5)
            if result.returncode != 0:
                missing_tools.append(tool)
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            missing_tools.append(tool)

    if missing_tools:
        pytest.fail(f"缺少必需的系统工具: {', '.join(missing_tools)}")

def test_project_modules_importable():
    """测试所有项目模块是否可以导入"""
    failed_imports = []

    for module in PROJECT_MODULES:
        try:
            importlib.import_module(module)
        except ImportError as e:
            failed_imports.append(f"{module}: {e}")

    if failed_imports:
        pytest.fail(f"模块导入失败: {'; '.join(failed_imports)}")

def test_virtual_environment_check():
    """检查虚拟环境状态（信息性测试，不会失败）"""
    # 检查 sys.prefix 是否指向虚拟环境
    base_prefix = getattr(sys, 'base_prefix', None) or getattr(sys, 'real_prefix', None)
    current_prefix = sys.prefix

    if base_prefix is None:
        # 没有虚拟环境激活
        print("警告: 未检测到虚拟环境，建议在虚拟环境中运行以避免依赖冲突")
    elif current_prefix == base_prefix:
        # 在系统 Python 环境中
        print("警告: 正在使用系统 Python 环境，建议使用虚拟环境")
    else:
        # 在虚拟环境中
        print("✓ 正在虚拟环境中运行")

    # 这个测试总是通过，只是提供信息
    assert True

def test_python_version_compatible():
    """测试 Python 版本兼容性"""
    import sys
    version = sys.version_info

    # 需要 Python 3.8+
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        pytest.fail(f"需要 Python 3.8+，当前版本: {version.major}.{version.minor}.{version.micro}")

    assert True

def test_project_structure():
    """测试项目结构完整性"""
    project_root = os.path.dirname(os.path.dirname(__file__))

    # 检查必需的目录
    required_dirs = [
        'UI_checktest',
    ]

    # 可选的目录（如果存在）
    optional_dirs = [
        '.venv',  # 虚拟环境目录
    ]

    missing_dirs = []
    for dir_name in required_dirs:
        dir_path = os.path.join(project_root, dir_name)
        if not os.path.isdir(dir_path):
            missing_dirs.append(dir_name)

    if missing_dirs:
        pytest.fail(f"缺少必需的目录: {', '.join(missing_dirs)}")

    # 检查可选目录
    missing_optional = []
    for dir_name in optional_dirs:
        dir_path = os.path.join(project_root, dir_name)
        if not os.path.isdir(dir_path):
            missing_optional.append(dir_name)

    if missing_optional:
        print(f"可选目录不存在: {', '.join(missing_optional)}")

    # 检查 python 包的 __init__.py
    init_file = os.path.join(project_root, 'UI_checktest', '__init__.py')
    if not os.path.isfile(init_file):
        pytest.fail("UI_checktest 包缺少 __init__.py 文件")

def test_python_syntax_check():
    """测试 Python 文件语法"""
    import ast
    import glob

    project_root = os.path.dirname(os.path.dirname(__file__))
    python_files = glob.glob(os.path.join(project_root, '**', '*.py'), recursive=True)

    syntax_errors = []
    for py_file in python_files:
        try:
            with open(py_file, 'r', encoding='utf-8') as f:
                source = f.read()
            ast.parse(source, py_file)
        except SyntaxError as e:
            syntax_errors.append(f"{py_file}: {e}")
        except Exception as e:
            syntax_errors.append(f"{py_file}: {e}")

    if syntax_errors:
        pytest.fail(f"语法错误: {'; '.join(syntax_errors)}")

def test_ui_check_basic():
    """基本 UI 检查测试"""
    # 这里可以调用你的 main 函数，但由于它是交互式的，可能需要模拟输入
    # 或者直接测试一些函数
    assert True  # 占位符测试

def test_import_modules():
    """测试模块导入"""
    try:
        import cv2
        import openpyxl
        import numpy as np
        assert True
    except ImportError as e:
        pytest.fail(f"Import failed: {e}")

def test_constants():
    """测试常量定义"""
    # 如果 python 包存在
    try:
        from UI_checktest.constants import BASE_RESULT_DIR, LOCAL_DIR, REMOTE_DIR
        assert BASE_RESULT_DIR is not None
        assert LOCAL_DIR is not None
        assert REMOTE_DIR == "/tmp/ui_screens"
    except ImportError:
        pytest.skip("python package not available")

def test_utils_logger():
    """测试日志函数"""
    try:
        from UI_checktest.utils import logger
        # 测试日志输出（不会实际写文件）
        logger("Test message")
        assert True
    except ImportError:
        pytest.skip("python package not available")

def test_click_safe_imread():
    """测试图片读取函数"""
    try:
        from UI_checktest.screenshot import safe_imread
        # 测试不存在的文件
        result = safe_imread("nonexistent.png")
        assert result is None
        # 测试 None 输入
        result = safe_imread(None)
        assert result is None
    except ImportError:
        pytest.skip("python package not available")

def test_compare_analyze_ui_diff():
    """测试 UI 差异分析"""
    try:
        from UI_checktest.compare import analyze_ui_diff
        # 测试相同路径
        result = analyze_ui_diff("same.png", "same.png")
        assert "Missing Path" in result or "Read Error" in result
    except ImportError:
        pytest.skip("python package not available")

@pytest.fixture
def temp_dir():
    """临时目录 fixture"""
    temp = tempfile.mkdtemp()
    yield temp
    shutil.rmtree(temp)

def test_report_creation(temp_dir):
    """测试报告创建"""
    try:
        from UI_checktest.report import create_report_workbook
        wb, ws = create_report_workbook(os.path.join(temp_dir, "test.xlsx"))
        assert wb is not None
        assert ws.title == "UI对比报告"
        assert ws['A1'].value == "语言ID"
    except ImportError:
        pytest.skip("python package not available")

def test_file_permissions():
    """测试文件权限"""
    project_root = os.path.dirname(os.path.dirname(__file__))

    # 检查关键文件是否可读
    key_files = [
        'UI_checktest/__init__.py',
        'UI_checktest/constants.py',
        'UI_checktest/utils.py',
        'UI_checktest/run.py',
        'UI_checktest/test_run.py',
    ]

    unreadable_files = []
    for file_path in key_files:
        full_path = os.path.join(project_root, file_path)
        if not os.path.isfile(full_path):
            unreadable_files.append(f"{file_path} (文件不存在)")
        elif not os.access(full_path, os.R_OK):
            unreadable_files.append(f"{file_path} (无读取权限)")

    if unreadable_files:
        pytest.fail(f"文件权限问题: {'; '.join(unreadable_files)}")

def test_network_connectivity():
    """测试网络连接（可选，用于检查 ADB 连接）"""
    try:
        # 尝试连接到本地 ADB 服务器
        result = subprocess.run(['adb', 'devices'],
                              capture_output=True,
                              text=True,
                              timeout=10)
        # ADB 工具必须可用，但不强制要求设备连接
        if result.returncode != 0 and 'no devices/emulators found' not in result.stdout.lower():
            pytest.fail(f"ADB 工具异常: {result.stderr}")
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pytest.skip("ADB 不可用或网络连接问题")

def test_memory_usage_check():
    """检查内存使用情况（信息性测试）"""
    try:
        import psutil
        import os

        process = psutil.Process(os.getpid())
        memory_info = process.memory_info()
        memory_mb = memory_info.rss / 1024 / 1024

        print(f"当前内存使用: {memory_mb:.2f} MB")

        # 基本内存检查（给出警告但不失败）
        if memory_mb > 500:
            print(f"警告: 内存使用较高 ({memory_mb:.2f} MB)")
        else:
            print("✓ 内存使用正常")

    except ImportError:
        print("psutil 未安装，跳过内存检查")

    # 这个测试总是通过
    assert True

if __name__ == "__main__":
    pytest.main([__file__, "-v"])

def test_ui_check_basic():
    """基本 UI 检查测试"""
    # 这里可以调用你的 main 函数，但由于它是交互式的，可能需要模拟输入
    # 或者直接测试一些函数
    assert True  # 占位符测试

def test_import_modules():
    """测试模块导入"""
    try:
        import cv2
        import openpyxl
        import numpy as np
        assert True
    except ImportError as e:
        pytest.fail(f"Import failed: {e}")

def test_constants():
    """测试常量定义"""
    # 如果 python 包存在
    try:
        from UI_checktest.constants import BASE_RESULT_DIR, LOCAL_DIR, REMOTE_DIR
        assert BASE_RESULT_DIR is not None
        assert LOCAL_DIR is not None
        assert REMOTE_DIR == "/tmp/ui_screens"
    except ImportError:
        pytest.skip("python package not available")

def test_utils_logger():
    """测试日志函数"""
    try:
        from UI_checktest.utils import logger
        # 测试日志输出（不会实际写文件）
        logger("Test message")
        assert True
    except ImportError:
        pytest.skip("python package not available")

def test_click_safe_imread():
    """测试图片读取函数"""
    try:
        from UI_checktest.screenshot import safe_imread
        # 测试不存在的文件
        result = safe_imread("nonexistent.png")
        assert result is None
        # 测试 None 输入
        result = safe_imread(None)
        assert result is None
    except ImportError:
        pytest.skip("python package not available")

def test_compare_analyze_ui_diff():
    """测试 UI 差异分析"""
    try:
        from UI_checktest.compare import analyze_ui_diff
        # 测试相同路径
        result = analyze_ui_diff("same.png", "same.png")
        assert "Missing Path" in result or "Read Error" in result
    except ImportError:
        pytest.skip("python package not available")

@pytest.fixture
def temp_dir():
    """临时目录 fixture"""
    temp = tempfile.mkdtemp()
    yield temp
    shutil.rmtree(temp)

def test_report_creation(temp_dir):
    """测试报告创建"""
    try:
        from UI_checktest.report import create_report_workbook
        wb, ws = create_report_workbook(os.path.join(temp_dir, "test.xlsx"))
        assert wb is not None
        assert ws.title == "UI对比报告"
        assert ws['A1'].value == "语言ID"
    except ImportError:
        pytest.skip("python package not available")

if __name__ == "__main__":
    pytest.main([__file__, "-v"])