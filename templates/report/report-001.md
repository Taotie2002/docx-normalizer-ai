# 报告 DFGP - report-001

## 基本信息

| 项目 | 值 |
|------|------|
| 模板ID | report-001 |
| 文档类型 | 报告 |
| 来源文件 | report-001-template.docx |
| 创建时间 | 2026-04-12 |
| 状态 | **已验证** |

---

## 字体规范

| 元素 | 字体 | 字号 | 说明 |
|------|------|------|------|
| 主标题 | 方正小标宋简体 | 22pt | 居中，无缩进 |
| 收文机关 | 仿宋_GB2312 | 16pt | 无缩进，左对齐 |
| 一级标题 | 黑体 | 16pt | X、格式，无缩进 |
| 二级标题 | 楷体_GB2312 | 16pt | （一）格式，无缩进 |
| 正文 | 仿宋_GB2312 | 16pt | 首行缩进32pt，两端对齐 |
| 结语（妥否请批示） | 仿宋_GB2312 | 16pt | 居中，无缩进 |
| 落款 | 仿宋_GB2312 | 16pt | 右对齐 |

---

## 段落格式

| 项目 | 值 | 说明 |
|------|------|------|
| 行距 | 固定值 28pt | |
| 段前距 | 0pt | |
| 段后距 | 0pt | |
| 一级标题段前距 | 12pt | |
| 一级标题段后距 | 6pt | |
| 结语段前距 | 12pt | |

---

## 首行缩进规则

| 段落类型 | 首行缩进 | 对齐 |
|---------|---------|------|
| 主标题 | 无 | 居中 |
| 收文机关（如"省政府："） | 无 | 左对齐 |
| 一级标题（一、） | 无 | 左对齐 |
| 二级标题（（一）） | 无 | 左对齐 |
| 正文 | 32pt (2字符) | 两端对齐 |
| 结语（如"妥否，请批示。"） | 无 | **居中** |
| 落款 | 无 | 右对齐 |

---

## 页面设置

| 项目 | 值 |
|------|------|
| 页边距 | 上3.7cm, 下3.5cm, 左2.8cm, 右2.6cm |
| 纸张 | A4 |
| 首页不同 | True |

---

## 识别特征

- **标题特征**：主标题以"关于...的报告"格式，居中方正小标宋22pt
- **收文机关**：以"省政府："、"市委："等格式开头，以冒号结尾
- **结语**：包含"妥否，请批示。"或"以上报告，请审阅。"等，**居中对齐**
- **正文**：仿宋_GB2312，首行缩进32pt

---

## 段落识别函数

```python
def is_report_title(text, index):
    """报告主标题：'关于...的报告'格式"""
    if index == 0 and '关于' in text and '报告' in text:
        return True
    if index == 0 and len(text) <= 30 and len(text) > 2:
        return True
    return False

def is_receiving_org(text):
    """收文机关：冒号结尾，短于20字"""
    if text.endswith('：') and len(text) < 20:
        return True
    return False

def is_report_conclusion(text):
    """报告结语：'妥否'/'以上报告'等"""
    if '妥否' in text and '批示' in text:
        return True
    if '以上' in text and ('报告' in text or '审阅' in text):
        return True
    return False

def is_title_l1(text):
    """一级标题：X、格式"""
    return bool(text.startswith(('一','二','三','四','五')) and '、' in text[:5])

def is_title_l2(text):
    """二级标题：（一）格式"""
    return text.startswith('（') and '）' in text[:6]

def is_report_signature(text):
    """报告落款"""
    if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text):
        return True
    return False
```

---

## 验证记录

| 日期 | 状态 | 说明 |
|------|------|------|
| 2026-04-12 | 已创建 | 初始版本 |

---

## 参考模板

- `report-001-template.docx` - 政务公开工作情况报告示例
