# 广西国控集团公文格式排版技能

根据广西国控资本运营集团有限责任公司公文格式规范（2025 年 7 月版）自动排版 Word 文档的 AI Agent 技能。

## ✨ 功能特性

- 📝 **自动识别**：智能识别公文标题、正文、各级标题层级
- 🎨 **格式规范**：严格按照集团公文格式标准排版
- ⚡ **批量处理**：支持批量格式化多个文档
- 🔧 **灵活配置**：支持自定义字体、字号、行距等参数

## 📋 格式规范

### 页面设置
- 页边距：上 37mm，下 35mm，左 28mm，右 26mm
- 行距：正文 29pt，标题 33pt

### 字体规范
| 元素 | 字体 | 字号 | 特殊要求 |
|------|------|------|---------|
| 公文标题 | 方正小标宋简体 | 22pt | 居中 |
| 正文 | 仿宋_GB2312 | 16pt | 首行缩进 2 字符 |
| 一级标题（一、） | 黑体 | 16pt | 行距 33pt |
| 二级标题（（一）） | 仿宋_GB2312 | 16pt | 加粗，行距 33pt |
| 三级标题（1.） | 仿宋_GB2312 | 16pt | 不加粗，行距 33pt |

## 🚀 安装方法

### 方式一：通过 Skills CLI 安装（推荐）

```bash
npx skills add <your-username>/gxgk-doc-format-skill -g
```

### 方式二：手动安装

1. 克隆或下载本仓库
2. 将 `gxgk-doc-format` 文件夹复制到以下位置之一：
   - 全局技能目录：`~/.workbuddy/skills/`
   - 项目技能目录：`{项目路径}/.workbuddy/skills/`

## 💡 使用方法

### 在 AI Agent 中使用

当您需要排版公文时，只需说：
- "帮我按广西国控集团格式排版这份公文"
- "格式化这个 Word 文档为集团标准"
- "发送前按集团规范排版"

### 命令行使用

```bash
python format_gxgk.py 输入文档.docx 输出文档.docx
```

### Python 代码调用

```python
from format_gxgk import format_document

format_document('input.docx', 'output.docx')
```

## 📁 文件结构

```
gxgk-doc-format/
├── SKILL.md          # 技能定义文件
├── format_gxgk.py    # Python 格式化脚本
└── README.md         # 使用说明文档
```

## ⚙️ 依赖要求

- Python 3.x
- python-docx 库

安装依赖：
```bash
pip install python-docx
```

## 🎯 适用场景

- 广西国控集团内部公文排版
- 党政机关公文格式标准化
- 企业 Word 文档格式统一
- 批量文档格式化任务

## 📝 版本历史

- **v1.0.0** (2025-07) - 初始版本，支持基础公文格式排版
  - 自动识别标题层级
  - 支持短标题智能识别
  - 保留用户已有序号，不自动添加

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request 来改进这个技能！

## 📄 许可证

MIT License

## 🔗 相关资源

- [Agent Skills 规范](https://agentskills.io)
- [Skills CLI 文档](https://skills.sh/docs/cli)
- [广西国控集团](https://www.gxgkjt.com)

## 💬 技术支持

如有问题或建议，请通过以下方式联系：
- 提交 Issue
- 邮件：huangsiqi9125@example.com

---

**注意**：使用本技能前，请确保系统中已安装所需字体（方正小标宋简体、仿宋_GB2312、楷体_GB2312 等）。
