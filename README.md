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

| 项目 | 状态 |
|------|------|
| 知识库框架 | ✅ 完成 |
| 公文DFGP模板 | ✅ 已验证 |
| 灌注工具 | ✅ 可用 |
| CLI命令 | 🚧 规划中 |
| 多Agent协同 | 🚧 规划中 |

---

## 目录结构

```
docx-dfgp-knowledge/
├── README.md              # 本文件
├── 分类索引.md            # 公文类型分类索引
├── requirements.txt       # Python依赖
├── scripts/
│   └── dfp_tool.py       # 公文格式灌注工具 (已验证)
└── templates/
    ├── government/        # 政府公文模板
    │   ├── gov-001-template.docx
    │   └── dfp-gov-001.md
    ├── request/           # 请示公文模板
    │   ├── 请示1-4.docx   # 4份真实请示文档
    │   └── dfp-request-001.md
    ├── student/           # 学生作业模板
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
python scripts/dfp_tool.py <输入文档> <输出文档> [页码模板]

# 示例
python scripts/dfp_tool.py input.docx output.docx templates/government/gov-001-template.docx
```

### 3. 灌注规则

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

| 类型 | 识别依据 |
|------|---------|
| 命令(令) | 标题含"命令"/"令"字 |
| 决定 | 正文含"经研究决定" |
| 公告 | 发布范围广泛、无主送机关 |
| 通告 | 事项性质面向社会 |
| 通知 | 标题含"通知"、有主送机关 |
| 通报 | 含"通报批评"/"表彰" |
| 报告 | 上行文、含"请审阅"/"请批示" |
| 请示 | 上行文、含"妥否，请批示" |
| 批复 | 回应请示、含"同意"/"不同意" |
| 议案 | 提交人大审议 |
| 函 | 平行文、无上下级隶属关系 |
| 决议 | 含"会议决定" |
| 指示 | 含安排部署性质 |
| 条例 | 法规性质、标题含"条例" |
| 纪要 | 标题含"纪要"、有出席人员 |

---

## 版本记录

| 版本 | 日期 | 内容 |
|------|------|------|
| 1.0 | 2026-04-09 | 初始：2类模板 |
| 2.0 | 2026-04-09 | 扩展15种公文类型 |
| 2.1 | 2026-04-10 | 审计评分权重调整，新增行间距审计 |

---

## 规划中功能

- [ ] CLI命令（extract/apply/diff/batch）
- [ ] 原子备份机制
- [ ] 多Agent协同（钱串子/代码驴/审计狗）
- [ ] YAML Schema解析器
- [ ] 批量处理能力

---

## 相关仓库

- [docx-normalizer-ai](https://github.com/Taotie2002/docx-normalizer-ai) - 本仓库，DFGP知识库
- [agent-main](https://github.com/Taotie2002/agent-main) - OpenClaw主Agent工作区
