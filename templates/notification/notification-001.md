# 通知 DFGP - notification-001

## 基本信息

| 项目 | 值 |
|------|------|
| 模板ID | notification-001 |
| 文档类型 | 通知 |
| 来源文件 | notification-001-template.docx |
| 创建时间 | 2026-04-12 |
| 状态 | **已验证** |

---

## 字体规范

| 元素 | 字体 | 字号 | 说明 |
|------|------|------|------|
| 主标题 | 方正小标宋简体 | 22pt | 居中，无缩进 |
| 发文编号 | 仿宋_GB2312 | 16pt | 右对齐 |
| 称谓行 | 仿宋_GB2312 | 16pt | 无缩进，左对齐 |
| 一级标题 | 黑体 | 16pt | X、格式，无缩进 |
| 二级标题 | 楷体_GB2312 | 16pt | （一）格式，无缩进 |
| 正文 | 仿宋_GB2312 | 16pt | 首行缩进32pt，两端对齐 |
| 附件说明 | 仿宋_GB2312 | 16pt | 首行缩进32pt |
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

---

## 首行缩进规则

| 段落类型 | 首行缩进 | 对齐 |
|---------|---------|------|
| 主标题 | 无 | 居中 |
| 发文编号 | 无 | 右对齐 |
| 称谓行 | 无 | 左对齐 |
| 一级标题（一、） | 无 | 左对齐 |
| 二级标题（（一）） | 无 | 左对齐 |
| 正文 | 32pt (2字符) | 两端对齐 |
| 附件说明 | 32pt | 两端对齐 |
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

- **标题特征**：主标题以"关于...的通知"格式，居中方正小标宋22pt
- **称谓行**：以"各市（州）人民政府，省直各部门："格式结尾有冒号
- **一级标题**：以"一、"、"二、"、"三、"、"四、"开头
- **二级标题**：以"（一）"、"（二）"等开头
- **正文**：仿宋_GB2312，首行缩进32pt

---

## 段落识别函数

```python
def is_notification_title(text, index):
    """通知主标题：'关于...的通知'格式，或index==0的短标题"""
    if index == 0 and '关于' in text and '的通知' in text:
        return True
    if index == 0 and len(text) <= 30 and len(text) > 2:
        return True
    return False

def is_doc_number(text):
    """发文编号：××〔2026〕XX号格式"""
    return bool(re.search(r'〔\d{4}〕\d+号', text))

def is_notification_salutation(text):
    """称谓行：冒号结尾，含'各市'、'各部门'、'各位'等"""
    if text.endswith('：') and len(text) < 30:
        if any(kw in text for kw in ['各市', '各部门', '各位', '贵']):
            return True
    return False

def is_title_l1(text):
    """一级标题：X、格式"""
    return bool(text.startswith(('一','二','三','四','五')) and '、' in text[:5])

def is_title_l2(text):
    """二级标题：（一）格式"""
    return text.startswith('（') and '）' in text[:6]

def is_attachment_note(text):
    """附件说明：含'附件'"""
    return '附件' in text

def is_notification_signature(text):
    """落款：单位名称+日期"""
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

- `notification-001-template.docx` - 2026年度信息安全检查通知示例
