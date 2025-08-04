# 安装指南

## 系统要求

- Python 3.7 或更高版本
- Windows 10/11, macOS 10.14+, 或 Linux (Ubuntu 18.04+)

## 安装步骤

### 1. 克隆项目
```bash
git clone https://github.com/gyl2004/Excel-score-query.git
cd Excel-score-query
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 配置文件（可选）
```bash
cp config.example.json config.json
```

## 快速开始

### 1. 生成测试数据
```bash
python generate_large_test_data.py
```

### 2. 运行程序

#### GUI界面（推荐）
```bash
python main.py
```

#### 命令行快速处理
```bash
python fast_processing_main.py "岗位表.xlsx" "面试人员名单.xlsx"
```

## 功能特性

- 🚀 高性能数据处理（支持万级数据）
- 📊 智能数据匹配和分析
- 🎨 友好的图形用户界面
- 📁 统一的文件管理系统
- ⚡ 多种处理引擎选择
- 🔧 灵活的列映射配置

## 更多信息

- 详细使用说明：[使用指南.md](使用指南.md)
- 技术文档：[DOCUMENTATION.md](DOCUMENTATION.md)
- 项目介绍：[README.md](README.md)

## 问题反馈

如有问题，请在 [GitHub Issues](https://github.com/gyl2004/Excel-score-query/issues) 中提交。