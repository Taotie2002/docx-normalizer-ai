# Docx DFGP 知识库

> 文档格式生成模式（Document Format Generation Pattern）持久化存储，用于自动化文档规格化
>
> 参考 GB/T 9704-2012《党政机关公文格式》SRS设计

---

## ⚠️ 最高原则

**GB/T 9704-2012《党政机关公文格式》是强制规范**

凡模板与该规范冲突的，**以该规范为准**，这是处理公文格式的**最高原则**！

---

## 项目状态

| 项目 | 状态 | 说明 |
|------|------|------|
| 知识库框架 | ✅ 完成 | 15种公文类型分类 |
| 公文DFGP模板 | ✅ 已验证 | 4份请示+政府模板 |
| 灌注工具 | ✅ 可用 | scripts/dfgp_tool.py |
| CLI脚本 | ✅ 骨架 | extract.py/apply.py |
| 多Agent协同 | 🚧 规划中 | 由agent-main仓库提供 |
| 原子备份 | 🚧 规划中 | - |

---

## 目录结构

```
docx-normalizer-ai/
├── README.md              # 本文件
├── 分类索引.md            # 公文类型分类索引
├── LICENSE               # MIT许可证
├── .gitignore            # Git忽略配置
├── .env.example          # 环境变量示例
├── requirements.txt      # Python依赖
├── scripts/
│   ├── __init__.py
│   ├── dfgp_tool.py       # 公文格式灌注工具
│   ├── extract.py        # 格式提取CLI (骨架)
│   └── apply.py          # 格式应用CLI (骨架)
└── templates/
    ├── government/        # 政府公文模板（含请示、报告等子类型）
    │   └── dfgp-gov-001.md
    ├── student/           # 学生作业模板
    │   ├── student-001-template.docx
    │   └── dfgp-student-001.md
    ├── academic/          # 🚧 待填充
    ├── announcement/      # 🚧 待填充
    ├── circular/          # 🚧 待填充
    └── ...                # 其他类型（待扩展）
```

---

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 使用灌注工具

```bash
# 基本用法
python scripts/dfgp_tool.py <输入文档> <输出文档> [页码模板]

# 示例
python scripts/dfgp_tool.py input.docx output.docx templates/government/gov-001-template.docx
```

### 3. 使用CLI脚本

```bash
# 提取模板格式（骨架）
python scripts/extract.py templates/government/gov-001-template.docx -o format.json

# 应用格式到目标文档（骨架）
python scripts/apply.py input.docx template.docx output.docx
```

---

## 灌注规则（GB/T 9704-2012）

| 段落类型 | 字体 | 字号 | 对齐 | 首行缩进 |
|---------|------|------|------|---------|
| 主标题 | 方正小标宋简体 | 22pt | 居中 | 无 |
| 一级标题 | 黑体 | 16pt | 左对齐 | 无 |
| 二级标题 | 楷体_GB2312 | 16pt | 左对齐 | 无 |
| 正文 | 仿宋_GB2312 | 16pt | 两端对齐 | 32pt |
| 称谓行 | 仿宋_GB2312 | 16pt | 左对齐 | 无 |

**页面设置**：上3.7cm，下3.5cm，左2.8cm，右2.6cm

**行间距**：固定值28pt

**页码**：奇数页居右，偶数页居左，首页不显示

---

## 已验证的段落识别函数

```python
def is_main_title(text, index):
    """主标题：第一个居中的短标题"""
    if index == 0 and '关于' in text and ('请示' in text or '报告' in text):
        return True
    if index == 0 and len(text) <= 30 and len(text) > 2:
        return True
    return False

def is_title_l1(text):
    """一级标题：X、格式"""
    return bool(text.startswith(('一','二','三','四','五','六','七','八','九','十')) and '、' in text[:5])

def is_title_l2(text):
    """二级标题：（一）格式"""
    return text.startswith('（') and '）' in text[:6]

def is_salutation(text):
    """称谓行：尊敬的...、市政府：等"""
    if text.startswith('尊敬的') and '领导' in text:
        return True
    if text.endswith('：') and len(text) < 20:
        return True
    return False
```

---

## 审计评分体系

| 维度 | 权重 | 要求 |
|------|------|------|
| 字体使用 | 30% | 正文=仿宋_GB2312，标题按级别 |
| 对齐方式 | 20% | 正文两端对齐，标题居中/左对齐 |
| 行间距 | 15% | 固定值28pt或1.5倍 |
| 首行缩进 | 15% | 正文32pt，标题0pt |
| 页面设置 | 10% | 3.7/3.5/2.8/2.6cm |
| 内容完整性 | 10% | 字符偏差<5% |

**阈值**：≥90%自动通过，<80%自动返工

---

## 公文类型分类（15种）

| 类型 | 目录 | 识别依据 | 状态 |
|------|------|---------|------|
| 命令(令) | order/ | 标题含"命令"/"令"字 | 🚧 待填充 |
| 决定 | decision/ | 正文含"经研究决定" | 🚧 待填充 |
| 公告 | announcement/ | 发布范围广泛、无主送机关 | 🚧 待填充 |
| 通告 | notice/ | 事项性质面向社会 | 🚧 待填充 |
| 通知 | notification/ | 标题含"通知"、有主送机关 | 🚧 待填充 |
| 通报 | circular/ | 含"通报批评"/"表彰" | 🚧 待填充 |
| 报告 | report/ | 上行文、含"请审阅"/"请批示" | 🚧 待填充 |
| 请示 | government/ | 上行文、含"妥否，请批示"，属于government子类型 | ✅ 已合并 |
| 批复 | reply/ | 回应请示、含"同意"/"不同意" | 🚧 待填充 |
| 议案 | proposal/ | 提交人大审议 | 🚧 待填充 |
| 函 | letter/ | 平行文、无上下级隶属关系 | 🚧 待填充 |
| 决议 | resolution/ | 含"会议决定" | 🚧 待填充 |
| 指示 | directive/ | 含安排部署性质 | 🚧 待填充 |
| 条例 | regulation/ | 法规性质、标题含"条例" | 🚧 待填充 |
| 纪要 | minutes/ | 标题含"纪要"、有出席人员 | 🚧 待填充 |
| 政府公文 | government/ | 通用公文格式 | ✅ 已验证 |
| 技术文档 | technical/ | API文档、技术手册 | 🚧 待填充 |
| 学术论文 | academic/ | 论文、期刊文章 | 🚧 待填充 |
| 学生作业 | student/ | 练习题、试卷 | ✅ 已验证 |

---

## 版本记录

| 版本 | 日期 | 内容 |
|------|------|------|
| 1.0 | 2026-04-09 | 初始：2类模板 |
| 2.0 | 2026-04-09 | 扩展15种公文类型 |
| 2.1 | 2026-04-10 | 审计评分权重调整，新增行间距审计 |
| 2.2 | 2026-04-12 | 修复仓库完整性，添加CLI骨架 |

---

## 相关仓库

- [docx-normalizer-ai](https://github.com/Taotie2002/docx-normalizer-ai) - 本仓库，DFGP知识库
- [agent-main](https://github.com/Taotie2002/agent-main) - OpenClaw主Agent工作区（含多Agent协同实现）
