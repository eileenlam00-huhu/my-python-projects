# UI Check Test Project

这是一个UI自动化测试项目，用于测试多语言Android应用的界面一致性。

## 项目结构

```
d:\my-python-projects\
├── UI_checktest\              # 核心测试包
│   ├── __init__.py           # 包初始化
│   ├── constants.py          # 全局常量配置
│   ├── utils.py              # 工具函数
│   ├── login.py              # ADB登录模块
│   ├── click.py              # 点击控制模块
│   ├── page.py               # 页面结构定义
│   ├── multilanguage.py      # 多语言处理
│   ├── screenshot.py         # 屏幕截图
│   ├── compare.py            # 图像对比
│   ├── report.py             # Excel报告生成
│   ├── run.py                # 主程序入口
│   └── test_run.py           # pytest测试套件
├── .venv\                    # Python虚拟环境
├── Muti_Web\                 # Web界面（Flask应用）
└── [其他脚本]
```

## 功能特性

- ✅ 多语言UI界面自动化测试
- ✅ 屏幕截图捕获和对比
- ✅ Excel报告生成（支持图片嵌入）
- ✅ ADB设备控制
- ✅ 完整的pytest测试覆盖
- ✅ 依赖检查和环境验证

## 安装和使用

### 1. 环境准备

```bash
# 激活虚拟环境
& .\.venv\Scripts\activate

# 安装依赖
pip install opencv-python numpy openpyxl pytest
# 可选: pip install pillow psutil
```

### 2. 运行测试

```bash
# 运行完整测试套件
python -m pytest UI_checktest/test_run.py -v

# 运行主程序
python UI_checktest/run.py
```

## 依赖要求

### 必需依赖
- Python 3.8+
- OpenCV (cv2)
- NumPy
- openpyxl
- pytest
- ADB (Android Debug Bridge)

### 可选依赖
- Pillow (用于Excel图片嵌入)
- psutil (用于系统监控)

## 使用说明

1. 连接Android设备并启用USB调试
2. 运行 `python UI_checktest/run.py`
3. 选择要测试的语言
4. 等待自动化测试完成
5. 查看生成的Excel报告

## 测试覆盖

项目包含完整的测试套件：
- 依赖检查测试
- 模块导入测试
- 功能单元测试
- 环境验证测试
- 语法检查测试

## 注意事项

- 确保ADB正常工作且设备已连接
- 测试过程中会自动截图和操作设备
- Excel报告会保存在 `C:\Users\[用户名]\Downloads\ui_test\png\` 目录下</content>
<parameter name="filePath">d:\my-python-projects\README.md