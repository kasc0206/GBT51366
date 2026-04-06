# 公共建筑碳排放计算报告系统 V1.0.0

基于 GB/T 51366-2019 和 JS/T 303-2026 的建筑碳排放精细化计算工具。

将建筑碳排放计算拆分为 **39个独立计算单元**，每个单元专注于一个具体的计算任务。

---

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 生成计算表格

```bash
python main.py              # 生成精细化Excel（39个工作表）
python main.py --word       # 生成Word报告模板（12个章节）
```

### 3. 填写数据并同步

```bash
# 填写Excel中的黄色单元格后，同步到Word
python main.py --sync
```

---

## 📁 项目结构

```
碳排放报告模板/
├── main.py                          # 统一入口
├── generate_granular.py             # Excel生成器（39个独立计算单元）
├── generate_word_template.py        # Word模板生成器（12个章节）
├── sync.py                          # OpenXML Excel-Word同步模块
├── requirements.txt                 # Python依赖
├── .gitignore                       # Git忽略规则
├── .markdownlint.json               # Markdown lint配置
│
├── templates/                       # 模板文件
│   ├── 公共建筑碳排放计算表_精细化版.xlsx   # 精细化Excel
│   └── 公共建筑碳排放计算报告.docx          # Word报告模板
│
└── docs/                            # 文档
    ├── README.md                    # 详细使用文档
    └── example_sync.py              # 同步代码示例
```

---

## 📊 核心功能

| 功能 | 说明 |
|------|------|
| **精细化计算** | 39个独立计算单元，单一职责原则 |
| **自动同步** | OpenXML技术，Excel数据自动同步到Word |
| **可扩展** | 新增计算单元只需添加独立工作表 |
| **跨平台** | 纯Python实现，无需VBA，支持macOS/Linux/Windows |

---

## 📖 文档

详细使用指南请查看：[docs/README.md](docs/README.md)

---

## 📚 参考标准

- GB/T 51366-2019《建筑碳排放计算标准》
- JS/T 303-2026《公共机构碳排放核算指南》
- GB/T 24040《环境管理 生命周期评价 原则与框架》
- GB/T 24044《环境管理 生命周期评价 要求与指南》

---

## ⚖️ 免责声明

本工具仅供参考使用，使用者应对其填写的数据和计算结果的准确性负责。

---

**版本**: V1.0.0  
**发布日期**: 2026-04-06  
**许可证**: MIT
