"""中国增值税发票字段解析模块 — v6: ¥三级回退 + 竖排兜底"""
import re


def parse_invoice(text):
    text = _normalize(text)
    result = {}
    result["items"] = _parse_items(text)
    result.update(_parse_amounts(text))
    result["invoice_date"] = _extract_field(text, r'([0-9]{4}[ ]*年[ ]*[0-9]{1,2}[ ]*月[ ]*[0-9]{1,2}[ ]*日)')
    result["remarks"] = _parse_remarks(text)
    result["invoice_number"] = _parse_invoice_number(text)
    return result


def _normalize(text):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text.strip()


def _extract_field(text, pattern):
    match = re.search(pattern, text, re.DOTALL)
    return match.group(1).strip() if match else ""


def _parse_items(text):
    """两种明细格式: A 完整跨行, B 简单星号"""
    items = []
    # A: 完整跨行明细
    pattern_a = re.compile(
        r'\*\s*(.+?)\s*\*\s*(.+?)\s*\n\s*(.+?)\s*\n\s*([\d.]+)\s*\n'
        r'\s*([\d,.]+)\s*\n\s*([\d,.]+)\s*\n\s*(\d+%?)\s*\n\s*([\d,.]+)',
        re.DOTALL
    )
    matched = set()
    for m in pattern_a.finditer(text):
        matched.add((m.start(), m.end()))
        items.append({
            "货物或服务名称": m.group(2).replace('\n', '').strip(),
            "规格型号": m.group(1).replace('\n', '').strip(),
            "数量": m.group(4).strip(), "单价": m.group(5).strip(),
            "金额": m.group(6).strip(), "税率": m.group(7).strip(),
        })
    # B: 只有 *分类*名称
    for m in re.finditer(r'\*\s*(.+?)\s*\*\s*(.+?)(?:\s|$)', text):
        if not any(ms <= m.start() and m.end() <= me for ms, me in matched):
            items.append({
                "货物或服务名称": m.group(2).replace('\n', '').strip(),
                "规格型号": m.group(1).replace('\n', '').strip(),
                "数量": "", "单价": "", "金额": "", "税率": "",
            })
    return items


def _parse_amounts(text):
    """价税合计 = max(¥金额) — 三级回退策略"""
    amounts = {"amount_excluding_tax": "", "total_tax": "", "total_amount": ""}

    # 策略1: ¥同行数字
    yen_pairs = [(float(m.group(1).replace(',', '')), m.group(1))
                 for m in re.finditer(r'¥[ \t]*([\d,.]+)', text)]
    # 策略2: ¥跨行数字
    if not yen_pairs:
        yen_pairs = [(float(m.group(1).replace(',', '')), m.group(1))
                     for m in re.finditer(r'¥\s*\n\s*([\d,.]+)', text)]

    if yen_pairs:
        yen_pairs.sort(key=lambda x: x[0], reverse=True)
        amounts["total_amount"] = yen_pairs[0][1]  # 最大 = 价税合计
        if len(yen_pairs) >= 3:
            amounts["total_tax"] = yen_pairs[-1][1]
            amounts["amount_excluding_tax"] = yen_pairs[1][1]
        elif len(yen_pairs) == 2:
            amounts["total_tax"] = yen_pairs[1][1]
            amounts["amount_excluding_tax"] = yen_pairs[1][1]
        return amounts

    # 策略3: ¥独立成行且跨行也失败 → (小写) 附近
    m = re.search(r'[（(]小写[)）]\s*¥?\s*([\d,.]+)', text)
    if m:
        amounts["total_amount"] = m.group(1).strip()
        return amounts

    # 策略4: 最后兜底 — 取最大金额数字
    all_nums = re.findall(r'([\d,]{2,}\.\d{2})', text)
    if all_nums:
        parsed = [(float(n.replace(',', '')), n) for n in all_nums]
        parsed.sort(key=lambda x: x[0], reverse=True)
        amounts["total_amount"] = parsed[0][1]

    return amounts


def _parse_remarks(text):
    m = re.search(r'¥[\d,.]+.*?\n(.+)$', text, re.DOTALL)
    return m.group(1).strip() if m else ""


def _parse_invoice_number(text):
    m = re.search(r'发票号码[：:\s]*(\d{8,30})', text)
    if m:
        return m.group(1).strip()
    for m in re.finditer(r'^(\d{15,30})\s*$', text, re.MULTILINE):
        return m.group(1).strip()
    return ""
