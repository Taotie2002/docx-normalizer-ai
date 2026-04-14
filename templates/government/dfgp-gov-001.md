# Government DFGP - gov-001

## 基本信息

| 项目 | 值 |
|------|------|
| 模板ID | gov-001 |
| 文档类型 | 政府公文（通用） |
| 包含子类型 | 请示、报告、批复等上行文 |
| 来源文件 | 测试目标公文等 |
| 创建时间 | 2026-04-09 |
| 更新时间 | 2026-04-14 |
| 状态 | **已验证** |

---

## 字体规范

| 元素 | 字体 | 字号 | 说明 |
|------|------|------|------|
| 主标题 | 方正小标宋简体 | 22pt | 居中 |
| 一级标题 | 黑体 | 16pt | X、格式 |
| 二级标题 | 楷体_GB2312 | 16pt | （一）格式 |
| 正文 | 仿宋_GB2312 | 16pt | 首行缩进32pt |
| 称谓行 | 仿宋_GB2312 | 16pt | 无缩进 |
| 结语 | 仿宋_GB2312 | 16pt | 居中 |
| 落款 | 仿宋_GB2312 | 16pt | 居右 |

---

## 段落格式

| 项目 | 值 | 说明 |
|------|------|------|
| 行距 | 固定值 28pt | EXACTLY |
| 段前距 | 0pt | |
| 段后距 | 0pt | |

---

## 首行缩进规则

> ⚠️ **重要教训**：除称谓行外，所有正文段落都必须首行缩进

| 段落类型 | 首行缩进 | 对齐 |
|---------|---------|------|
| 主标题 | 无 | 居中 |
| 一级标题 | 无 | 左对齐 |
| 二级标题 | 无 | 左对齐 |
| 称谓行（如"市政府："） | **无** | 左对齐 |
| 正文 | 32pt (2字符) | 两端对齐 |
| 结语（如"妥否，请批示。"） | 无 | 居中 |
| 落款 | 无 | 居右 |

---

## 页面设置

| 项目 | 值 |
|------|------|
| 页边距 | 上3.7cm, 下3.5cm, 左2.8cm, 右2.6cm |
| 纸张 | A4 |
| 首页不同 | True |

---

## 页码规范

| 项目 | 值 | 说明 |
|------|------|------|
| 格式 | —{页码}— | em dash + 数字 + em dash |
| 字体 | 仿宋_GB2312 | |
| 字号 | 14pt (twips=28) | |
| 奇数页位置 | 页脚居右 | |
| 偶数页位置 | 页脚居左 | |
| 首页 | 不显示页码 | |

---

## 段落识别函数（已验证）

> ⚠️ **核心经验**：必须先准确识别段落类型，再应用格式

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
    if text.endswith('：') and len(text) < 20:  # 如"市政府："、"各位领导："
        return True
    return False

def is_conclusion(text):
    """结语：妥否，请批示"""
    return '妥否' in text and '批示' in text

def is_signature(text):
    """落款：单位名称或日期"""
    if any(kw in text for kw in ['管理局', '办公室', '委员会', '办公厅']):
        return True
    if re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text):
        return True
    return False
```

---

## 焦土策略（已验证）

> ⚠️ **核心经验**：必须彻底清除原始格式后再灌注

```python
def scorched_earth(para):
    """彻底清除段落和run的所有格式"""
    # 1. 清除阴影效果
    for run in para.runs:
        for shd in run._element.findall('.//{...}shd'):
            run._element.remove(shd)
    
    # 2. 清除段落对齐
    pPr = para._element.find('{...}pPr')
    if pPr is not None:
        for jc in pPr.findall('.//{...}jc'):
            pPr.remove(jc)
    
    # 3. 清除run格式（保留纯文本）
    for run in para.runs:
        rPr = run._element.find('{...}rPr')
        if rPr is not None:
            for child in list(rPr):
                tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
                if tag not in ['rFonts', 'sz', 'szCs']:
                    rPr.remove(child)
    
    # 4. 重建run（保留文本内容）
    full_text = para.text
    for run in para.runs:
        run.text = ''
    if para.runs:
        para.runs[0].text = full_text
    else:
        para.add_run(full_text)
```

---

## 灌注流程

1. **识别段落类型** → 使用已验证的识别函数
2. **焦土处理** → 彻底清除原始格式
3. **应用格式** → 按段落类型设置字体/对齐/缩进/行距
4. **添加页码** → 奇偶页不同，首页不显示

---

## 验证记录

| 日期 | 置信度 | 状态 | 说明 |
|------|--------|------|------|
| 2026-04-09 | 97.5% | PASS | 初始版本 |
| 2026-04-10 | 97.8% | PASS | 测试目标4 |
| 2026-04-10 | 97.8% | PASS | 测试公文v2 |
| 2026-04-10 | - | 更新 | 固化经验教训 |

---

## 常见错误及教训

| 错误 | 教训 |
|------|------|
| 标题识别错误导致字体混用 | 必须先识别段落类型再应用格式 |
| 称谓行被加了首行缩进 | 称谓行必须无缩进 |
| 主标题被误判为一级标题 | 主标题只判断index==0的第一个短标题 |
| 页码添加到首页 | 必须设置titlePg=1 |
| 奇偶页脚相同 | 必须两个footer：footer1(奇数右)、footer2(偶数左) |
| subagent脚本质量差 | 关键处理由main执行 |

---

## 参考文件

- `/home/zyu/.openclaw/media/inbound/请示2---159f9e22-4054-4c15-bb9d-2169d31ef81f.docx` - 请示模板（含正确页脚）
