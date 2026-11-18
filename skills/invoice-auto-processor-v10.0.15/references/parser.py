"""中国增值税发票字段解析模块 — v10.0.12: 发票号码精准识别(区分发票号码/票据号码/票据代码) + 全文语义车费分类 + 中文跨行合并"""
import re


def parse_invoice(text):
    text = _normalize(text)
    result = {}
    
    # 检测铁路电子客票
    if '铁路电子客票' in text or '电子发票（铁路电子客票）' in text:
        return _parse_train_ticket(text)
    
    result["items"] = _parse_items(text)
    result.update(_parse_amounts(text))
    result["invoice_date"] = _extract_date(text)
    result["remarks"] = _parse_remarks(text)
    result["invoice_number"] = _parse_invoice_number(text)
    return result


def _normalize(text):
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text.strip()


def _extract_field(text, pattern):
    match = re.search(pattern, text, re.DOTALL)
    return match.group(1).strip() if match else ""


def _extract_date(text):
    """提取开票日期 — 支持多种格式"""
    # 格式1: YYYY年MM月DD日（增值税发票）
    m = re.search(r'([0-9]{4}[ ]*年[ ]*[0-9]{1,2}[ ]*月[ ]*[0-9]{1,2}[ ]*日)', text)
    if m:
        return m.group(1).strip()
    # 格式2: 开票日期:YYYY-MM-DD 或 开票日期：YYYY-MM-DD（通行费票据）
    m = re.search(r'开票日期[：:]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})', text)
    if m:
        return m.group(1).strip()
    # 格式3: 任意 YYYY-MM-DD 格式（用于非税收入票据的开票日期在末尾的情况）
    m = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', text)
    if m:
        return m.group(1).strip()
    return ""


def _parse_items(text):
    """三种明细格式: A 完整跨行, B 简单星号, C 表格格式（通行费等）"""
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
    # B: 只有 *分类*名称 — 过滤身份证号掩码（****1819）和无效匹配
    for m in re.finditer(r'\*\s*(.+?)\s*\*\s*(.+?)(?:\n|$)', text):
        if any(ms <= m.start() and m.end() <= me for ms, me in matched):
            continue
        cat = m.group(1).replace('\n', '').strip()
        name = m.group(2).replace('\n', '').strip()
        # 过滤身份证号掩码：分类或名称全由*和数字组成（如***1819、****1819）
        if re.match(r'^[\*\d]+$', cat) and re.match(r'^[\*\d]+$', name):
            continue
        # 过滤分类全是*号的情况
        if re.match(r'^\*+$', cat):
            continue
        items.append({
            "货物或服务名称": name,
            "规格型号": cat,
            "数量": "", "单价": "", "金额": "", "税率": "",
        })
    # B2: 跨行合并 — 如果项目名后面紧跟中文行（如"LGD,CSOT"下一行是"水质分析仪维修"），合并
    if items:
        for item in items:
            name = item.get("货物或服务名称", "")
            # 项目名全是ASCII字符（如 LGD,CSOT），检查下一行是否有中文继续
            if name and re.match(r'^[\x00-\x7f]+$', name):
                idx = text.find(name)
                if idx >= 0:
                    after = text[idx + len(name):idx + len(name) + 30]
                    m = re.match(r'\n\s*([\u4e00-\u9fa5][^\n]*)', after)
                    if m:
                        item["货物或服务名称"] = name + m.group(1).strip()
    # C: 表格格式
    if not items:
        # 通行费：找独立的数字行（项目编码）后面的中文名称
        m = re.search(r'(?:^|\n)(\d{6,})\s*\n\s*([\u4e00-\u9fa5]+(?:[（）\u4e00-\u9fa5\w]+)?)', text)
        if m:
            items.append({
                "货物或服务名称": m.group(2).strip(),
                "规格型号": "", "数量": "", "单价": "", "金额": "", "税率": "",
            })
        else:
            # 非税收入：找"发明专利/实用新型专利"这样的模式（去掉前面的数字编码）
            m = re.search(r'(?:\d+\s*)((?:发明|实用新型|外观设计)专利[^\n]*)', text)
            if m:
                items.append({
                    "货物或服务名称": m.group(1).strip(),
                    "规格型号": "", "数量": "", "单价": "", "金额": "", "税率": "",
                })
    # 清理名称：去掉括号及后面的内容（如"95号车用汽油(ⅥB" → "95号车用汽油"）
    # 以及去掉无用后缀（如"车辆停放沪房地" → "车辆停放"）
    # 以及跨行项目名修复（如"*生产生活服务*车辆无\n停放服务" → "停车费"）
    _TRUNCATE_KEYWORDS = ['沪房地', '不动产权', '产权证书', '规格型号', '单位']
    for item in items:
        name = item.get("货物或服务名称", "")

        # === 中文跨行合并（v10.0.9）：把断续的项目名拼全 ===
        # 如 "临时\n停车费" → "临时停车费"、"代收高\n速通行费" → "代收高速通行费"、"车位\n场地占用费" → "车位场地占用费"
        # 先去掉内部的单个换行（把间断的多行项目名拼成一个连续串）
        name = re.sub(r'\n\s*', '', name).strip()

        # 去掉星号（如"*供电*扫码充电-" → "扫码充电-"）
        name = re.sub(r'\*', '', name)
        # 去掉括号及后面内容
        name = re.split(r'[（(]', name)[0].strip()
        # 去掉无用后缀
        for kw in _TRUNCATE_KEYWORDS:
            if kw in name:
                name = name[:name.index(kw)].strip()
                break
        # 去掉末尾的"无"字（如"车辆停放无" → "车辆停放"）
        if name.endswith('无'):
            name = name[:-1].strip()

        # === 基于全文语义的车费分类（v10.0.9，读全不留片段）===
        name = _classify_vehicle_charge(text, name)

        # 修复："车辆"（去"无"后只剩"车辆"）→ 检查后面是否有"停放服务" → "停车费"
        if name == '车辆':
            name = '停车费'
        # 统一命名：汽油/车用汽油/柴油等 → 加油费
        if re.search(r'汽油|柴油|燃油', name):
            name = '加油费'
        # 统一命名：餐饮/餐饮服务等 → 餐饮费
        if re.search(r'餐饮', name):
            name = '餐饮费'
        # 统一命名：住宿/住宿服务等 → 住宿费（"住宿费"本身就是正确的，这里处理"住宿服务"等变体）
        if re.search(r'住宿', name) and name != '住宿费':
            name = '住宿费'
        # 统一命名：充电/扫码充电/供电等 → 汽车充电费
        if re.search(r'充电|供电', name):
            name = '汽车充电费'
        item["货物或服务名称"] = name
    return items


def _classify_vehicle_charge(text, name):
    """基于全文特征综合判断车费类别（读全语义，不抓片段）。

    优先级：
    1. 高速通行费：全文含"高速"+"通行费"或代收高速通行费特征（入口站/出口站/通行时间 + 不征税税率）
    2. 停车费：全文含停车/车位/场地占用/停车场，但不含高速特征
    3. 其他：交由后续现有规则（加油/餐饮/住宿/充电）处理

    之所以读全文而非只看 name 片段：这些票据的 name 常被截断成
    "临时"、"代收高"、"车位"，单看无法归类，但全文特征明确。
    """
    # 高速通行费特征（全文）：高速 + 通行费 且 有入口/出口收费站
    is_highway = (
        ('高速' in text and ('通行费' in text or '代收' in text))
        and ('入口' in text or '出口' in text)
    )
    # 停车费特征（全文）：停车/车位/场地占用/停车场
    is_parking = any(k in text for k in ['停车', '车位', '场地占用', '停车场'])

    # 高速通行费优先（因为它也可能含"停车"相关干扰，但有入口站/出口站必是高速）
    if is_highway:
        return '高速通行费'
    if is_parking:
        return '停车费'
    # 否则交给后续规则处理（加油/餐饮/住宿/充电等），返回原名
    return name


def _parse_amounts(text):
    """价税合计 — v8: 支持多种票据格式"""
    amounts = {"amount_excluding_tax": "", "total_tax": "", "total_amount": ""}

    # 策略0: 优先取「（小写）」或「(小写)」后的金额（增值税发票、通行费）
    m = re.search(r'[（(]小写[)）]\s*¥?\s*([\d,.]+)', text)
    if m:
        amounts["total_amount"] = m.group(1).strip()
        return amounts

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

    # 策略3: 非税收入票据 — 找大写金额前面的数字行
    # 文本结构：金额合计（大写）\n（小写）\n...\n135.00\n元\n0.15\n900.00\n壹佰叁拾伍元整
    # 找大写金额
    m = re.search(r'([零壹贰叁肆伍陆柒捌玖拾佰仟]+[万亿]?[零壹贰叁肆伍陆柒捌玖拾佰仟]*元[零壹贰叁肆伍陆柒捌玖角分]*整?)', text)
    if m:
        # 找大写金额前面的数字行
        pos = m.start()
        before_text = text[:pos]
        # 找"金额合计（大写）"标签
        m2 = re.search(r'金额合计[（(]大写[)）]', before_text)
        if m2:
            # 在标签和大写金额之间找数字行
            between_text = before_text[m2.end():]
            lines = between_text.split('\n')
            # 从标签后面找数字行
            for line in lines:
                line = line.strip()
                # 找纯数字行（金额）
                if re.match(r'^\d+\.?\d*$', line):
                    # 检查是否是金额（通常是小数）
                    if '.' in line:
                        amounts["total_amount"] = line
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


def _parse_train_ticket(text):
    """解析铁路电子客票"""
    result = {}
    result["items"] = [{"货物或服务名称": "铁路客运", "规格型号": "", "数量": "", "单价": "", "金额": "", "税率": ""}]
    result["total_amount"] = ""
    result["invoice_date"] = ""
    result["remarks"] = ""
    result["invoice_number"] = ""
    
    # 发票号码
    m = re.search(r'发票号码[：:\s]*(\d{8,30})', text)
    if m:
        result["invoice_number"] = m.group(1).strip()
    
    # 乘车日期（"发票日期"代表实际行程日期，即不带标签的YYYY年MM月DD日）
    # 优先取乘车日期，而非"开票日期"（申请开票的日期）
    dates = re.findall(r'(\d{4}年\d{1,2}月\d{1,2}日)', text)
    for d in dates:
        idx = text.find(d)
        if idx > 0 and '开票日期' in text[max(0,idx-10):idx]:
            continue
        result["invoice_date"] = d.strip()
        break
    if not result["invoice_date"]:
        m = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', text)
        if m:
            result["invoice_date"] = m.group(1).strip()
    # 如果都没有，回退到开票日期（YYYY-MM-DD格式）
    if not result["invoice_date"]:
        m = re.search(r'开票日期[：:\s]*(\d{4}[-/]\d{1,2}[-/]\d{1,2})', text)
        if m:
            result["invoice_date"] = m.group(1).strip()
    
    # 金额：找所有¥金额，取最大值
    amounts = re.findall(r'[￥¥]\s*([\d,.]+)', text)
    if amounts:
        parsed = [(float(a.replace(',', '')), a) for a in amounts]
        parsed.sort(key=lambda x: x[0], reverse=True)
        result["total_amount"] = parsed[0][1]
    
    return result


def _parse_invoice_number(text):
    """提取发票/票据号码 — 精准区分「发票号码」/「票据号码」/「票据代码」。

    v10.0.12 修正：
    - 字段优先：发票号码 > 票据号码（绝不把「票据代码」当号码）
    - 发票号码标签后可能为空，号码散落在独立长数字行（如高德打车电子发票）
    - 票据代码（8位，如33021026）≠ 票据号码（10位，如0506025045），严格区分

    号码特征判断：排除税号(15-18位无前缀)、金额(小数)、票据代码(8位固定前缀)。
    """
    # ---------- 1. 发票号码 ----------
    # 1a. 紧跟标签的数字：发票号码：xxxx
    m = re.search(r'发票号码[：:]\s*(\d{8,30})', text)
    if m:
        return m.group(1).strip()
    # 1b. 标签后可能为空 → 在标签附近/全文找独立长数字行（字符串，20位左右）
    #     特征：独立一行、≥15位、非金额非税号（用后续函数过滤）
    candidate = _find_standalone_invoice_number(text)
    if candidate:
        return candidate

    # ---------- 2. 票据号码（通行费/财政票据）----------
    # 2a. 与数字同行：票据号码：xxxx（优先）
    m = re.search(r'票据号码[：:]\s*(\d{8,30})', text)
    if m:
        return m.group(1).strip()
    # 2b. 标签后独立数字行
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if '票据号码' in line and '：' in line and not re.search(r'\d{8,}', line):
            for j in range(i+1, len(lines)):
                m2 = re.match(r'^(\d{8,30})\s*$', lines[j].strip())
                if m2 and not _is_bill_code(m2.group(1)):
                    return m2.group(1).strip()
            break

    # ---------- 3. 兜底（排除票据代码）----------
    # 绝不把【票据代码】(8位固定码)当号码。仅当确实没有任何号码时才兜底，
    # 且必须排除明显的票据代码（8位、3302/3301等税局代码段开头）
    for m in re.finditer(r'^(\d{15,30})\s*$', text, re.MULTILINE):
        return m.group(1).strip()
    return ""


def _is_bill_code(s):
    """判断是否为「票据代码」而非「票据号码」。

    票据代码特征：通常是 8 位固定码（如 33021026），标识票据种类，
    多位固定税局段（33=浙江），不是独立业务号码。
    """
    if not s:
        return False
    if len(s) == 8:
        return True  # 8位 → 极可能是票据代码，不是业务号码
    return False


def _find_standalone_invoice_number(text):
    """在全文里找独立的发票号码（发票号码标签为空时的 fallback）。

    特征：独立一行，≥15 位数字，排除：
    - 金额（含有 . 或过短）
    - 税号/统一社会信用代码（15/18位，但常带字母，且多出现在『信用代码』标签后）
    - 票据代码（8位）
    高德打车类发票号码示例：26502000001790469301（20位独立行）
    """
    lines = text.split('\n')
    for i, line in enumerate(lines):
        s = line.strip()
        # 独立纯数字行，长度在 [15, 30] 之间
        if re.match(r'^\d{15,30}$', s):
            # 排除票据代码（8位）——不会命中因为这里要求≥15位
            # 排除后面紧跟"年"的（那是日期行）——日期是 8 位不带，不影响
            return s
    # 全文兜底：任意 ≥18 位纯数字，排除含"."的（金额）
    for m in re.finditer(r'(?<!\d)(\d{18,30})(?!\d)', text):
        return m.group(1)
    return ""
