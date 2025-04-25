# PDF转Word在线转换工具

这是一个简单的网页应用，可以将PDF文件转换为Word文档格式。

## 功能特点

- 支持PDF文件拖拽上传
- 简洁美观的用户界面
- 实时转换状态显示
- 自动下载转换后的文件
- 支持大文件转换（最大16MB）

## 安装要求

- Python 3.7+
- pip（Python包管理器）

## 安装步骤

1. 克隆或下载此项目到本地

2. 创建并激活虚拟环境（推荐）：
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. 安装依赖包：
```bash
pip install -r requirements.txt
```

## 运行应用

1. 在终端中运行：
```bash
python app.py
```

2. 打开浏览器访问：
```
http://localhost:5000
```

## 使用说明

1. 点击上传区域或将PDF文件拖放到指定区域
2. 选择要转换的PDF文件
3. 点击"开始转换"按钮
4. 等待转换完成，转换后的Word文件会自动下载

## 注意事项

- 仅支持PDF文件格式
- 文件大小限制为16MB
- 确保有足够的磁盘空间用于临时文件存储

## 联系方式

作者：张睿
邮箱：596085203@qq.com 