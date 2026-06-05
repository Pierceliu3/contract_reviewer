import json
import re
import streamlit as st
from docx import Document
import pypdf
from datetime import datetime

# ============================================
# 文本提取
# ============================================
def extract_text_from_docx(file_obj):
    doc = Document(file_obj)
    lines = []
    for para in doc.paragraphs:
        if para.text.strip():
            lines.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            if any(cells):
                lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)

def extract_text_fallback(uploaded_file):
    file_type = uploaded_file.name.split('.')[-1].lower()
    if file_type == "pdf":
        reader = pypdf.PdfReader(uploaded_file)
        return "\n".join(page.extract_text() for page in reader.pages if page.extract_text())
    elif file_type == "docx":
        return extract_text_from_docx(uploaded_file)
    else:
        return uploaded_file.read().decode("utf-8")

def extract_contract_text(uploaded_file):
    return extract_text_fallback(uploaded_file)

# ============================================
# 辅助函数：归一化文本（去除标点、空白、统一小写）
# ============================================
def normalize_text(s):
    if not isinstance(s, str):
        return ""
    # 去除标点符号（保留字母数字中文空格）
    s = re.sub(r'[^\w\u4e00-\u9fff\s]', '', s)
    # 去除多余空格，转小写
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s

def extract_numbers(s):
    """提取字符串中的数字（整数或小数）"""
    match = re.search(r'(\d+(?:\.\d+)?)', s)
    return match.group(1) if match else None

# ============================================
# 智能比较函数
# ============================================
def is_numeric_equal(num_str1, num_str2):
    """
    判断两个数值字符串是否代表相同的数值（忽略尾随零，但保留非零小数精度）。
    例如：
        "123" 和 "123.00"  -> True
        "123" 和 "123.12"  -> False
        "123.0" 和 "123"   -> True
        "123.12" 和 "123.120" -> True (尾随零)
        "123.12" 和 "123.1200" -> True
    """
    if num_str1 is None or num_str2 is None:
        return False
    try:
        f1 = float(num_str1)
        f2 = float(num_str2)
        if f1 != f2:
            return False
        # 值相等，检查小数部分的“有效非零位数”
        def has_nonzero_decimal(s):
            if '.' not in s:
                return False
            # 去除尾随零后的小数部分
            decimal_part = s.split('.')[1].rstrip('0')
            return len(decimal_part) > 0 and any(c != '0' for c in decimal_part)
        has1 = has_nonzero_decimal(num_str1)
        has2 = has_nonzero_decimal(num_str2)
        return has1 == has2
    except:
        return False

def smart_compare(front_val, table_val, field_name):
    if front_val == "未找到" or table_val == "未找到":
        return False, front_val, table_val, "缺失数据"
    
    # 需要精确匹配的字段（保留原始标点符号）
    exact_fields = ["提交地点", "采购项目名称", "项目地点"]
    if field_name in exact_fields:
        front_clean = front_val.strip()
        table_clean = table_val.strip()
        if front_clean == table_clean:
            return True, front_clean, table_clean, "精确匹配"
        else:
            return False, front_clean, table_clean, "内容不一致"
    
    # 收件联系人特殊处理（姓名+手机号）
    if field_name == "收件联系人":
        front_phone = re.search(r'\d{11}', front_val)
        table_phone = re.search(r'\d{11}', table_val)
        front_name = re.search(r'[\u4e00-\u9fff]{2,4}', front_val)
        table_name = re.search(r'[\u4e00-\u9fff]{2,4}', table_val)
        phone_ok = (front_phone and table_phone and front_phone.group(0) == table_phone.group(0))
        name_ok = (front_name and table_name and front_name.group(0) == table_name.group(0))
        if phone_ok and name_ok:
            return True, f"{front_name.group(0)} {front_phone.group(0)}", f"{table_name.group(0)} {table_phone.group(0)}", "姓名+手机号匹配"
        else:
            return False, front_val, table_val, "姓名或手机号不一致"

    # 建设规模：数值+单位比较（应用通用数值比较）
        # 建设规模：数值+单位比较（单位归一化）
        # 建设规模：数值+单位比较（单位归一化）
    if field_name == "建设规模":
        front_num_str = extract_numbers(front_val)
        table_num_str = extract_numbers(table_val)
        
        # 归一化单位：将常见等价单位映射为统一字符串
        def normalize_unit(unit):
            if not unit:
                return ""
            unit_lower = unit.lower()
            if unit_lower in ['平方米', '㎡', 'm2', 'm²', '平方']:
                return "平方米"
            return unit_lower
        
        def extract_unit(s):
            # 优先匹配常见单位（长单位优先）
            match = re.search(r'(\d+(?:\.\d+)?)\s*(平方米|㎡|m2|m²|平方|m)', s)
            if match:
                return match.group(2)
            # 兜底：取数字后第一个非数字非空白序列（最多3字符）
            match = re.search(r'\d+(?:\.\d+)?\s*([^\d\s]{1,3})', s)
            if match:
                return match.group(1)
            return ""
        
        front_unit_raw = extract_unit(front_val)
        table_unit_raw = extract_unit(table_val)
        front_unit = normalize_unit(front_unit_raw)
        table_unit = normalize_unit(table_unit_raw)
        
        if front_num_str and table_num_str:
            if is_numeric_equal(front_num_str, table_num_str):
                if front_unit == table_unit:
                    return True, front_val, table_val, "数值+单位一致"
                else:
                    return False, front_val, table_val, f"数值相同但单位不同 (前部单位: {front_unit_raw}, 须知单位: {table_unit_raw})"
            else:
                return False, front_val, table_val, f"数值不一致 (前部: {front_num_str}, 须知: {table_num_str})"
        else:
            return False, front_val, table_val, "无法提取数值"
    
        # 计划工期 / 工期要求（比较天数、开工日期）
        # 计划工期 / 工期要求（同时比较天数和开工日期）
    if field_name in ["计划工期", "工期要求"]:
        # 提取日历天数（支持“90天”或“90日历天”）
        front_days_match = re.search(r'(\d+)\s*(?:日历天|天)', front_val)
        table_days_match = re.search(r'(\d+)\s*(?:日历天|天)', table_val)
        front_days = front_days_match.group(1) if front_days_match else None
        table_days = table_days_match.group(1) if table_days_match else None
        days_ok = is_numeric_equal(front_days, table_days) if front_days and table_days else False
        
        # 提取开工日期（支持空格，例如“2026 年 6 月 10 日”）
        def extract_start_date(s):
            # 匹配 “2026年6月10日” 或 “2026 年 6 月 10 日” 等
            match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]', s)
            if match:
                y, m, d = match.groups()
                return f"{y}年{int(m):02d}月{int(d):02d}日"
            return None
        
        front_start = extract_start_date(front_val)
        table_start = extract_start_date(table_val)
        start_ok = (front_start is not None and table_start is not None and front_start == table_start)
        
        # 判断：如果双方都有天数，必须天数相等；如果双方都有开工日期，必须开工日期相等
        # 如果一方缺失某个字段，则视为不一致（可调整为宽松，但用户要求严格）
        has_days = (front_days is not None and table_days is not None)
        has_start = (front_start is not None and table_start is not None)
        
        if has_days and has_start:
            if days_ok and start_ok:
                return True, front_val, table_val, "天数和开工日期均一致"
            else:
                return False, front_val, table_val, "天数或开工日期不一致"
        elif has_days:
            # 只有天数，没有开工日期
            if days_ok:
                # 如果用户允许只比较天数，则通过；否则失败。这里按用户要求需要开工日期也一致，所以失败
                return False, front_val, table_val, "缺少开工日期信息"
            else:
                return False, front_val, table_val, "天数不一致"
        elif has_start:
            # 只有开工日期，没有天数
            if start_ok:
                return False, front_val, table_val, "缺少工期天数信息"
            else:
                return False, front_val, table_val, "开工日期不一致"
        else:
            # 双方都缺失天数和开工日期，无法比较
            return False, front_val, table_val, "缺少工期和开工日期信息"
    # 报价及单价总价计价方式：关键词匹配
    if field_name == "报价及单价总价计价方式":
        keywords = ["固定综合单价", "固定单价", "综合单价", "可调价", "成本加酬金"]
        front_key = next((kw for kw in keywords if kw in front_val), None)
        table_key = next((kw for kw in keywords if kw in table_val), None)
        if front_key and table_key and front_key == table_key:
            return True, front_key, table_key, "关键词匹配"
        norm_front = normalize_text(front_val)
        norm_table = normalize_text(table_val)
        if norm_front in norm_table or norm_table in norm_front:
            return True, norm_front, norm_table, "归一化包含匹配"
        return False, front_val, table_val, "计价方式不一致"

    if field_name == "截止时间":
        front_clean = re.sub(r'\s+', '', front_val)
        table_clean = re.sub(r'\s+', '', table_val)
        if front_clean == table_clean:
            return True, front_clean, table_clean, "忽略空格匹配"
        else:
            return False, front_val, table_val, "内容不一致"

    # 采购范围：宽松包含匹配
        # 采购范围：宽松相似度匹配
    if field_name == "采购范围":
        import difflib
        # 清洗：移除常见修饰词、括号内容、标点空格
        def clean_scope(s):
            s = re.sub(r'（[^）]*）', '', s)          # 去除中文括号内容
            s = re.sub(r'\([^)]*\)', '', s)          # 去除英文括号内容
            s = re.sub(r'以[^。]*为准', '', s)
            s = re.sub(r'根据[^，,。]*[,，]?', '', s)
            s = re.sub(r'详见[^，,。]*[,，]?', '', s)
            s = re.sub(r'具体[^，,。]*[,，]?', '', s)
            s = re.sub(r'包含但不限于', '', s)
            s = re.sub(r'[，,。.、；;:：\s]+', '', s) # 去除标点空格
            # 提取连续的中文（也可保留英文数字，但采购范围主要是中文）
            s = re.sub(r'[^a-zA-Z\u4e00-\u9fff]', '', s)  # 只留字母和中文（可选）
            return s.strip()
        front_clean = clean_scope(front_val)
        table_clean = clean_scope(table_val)
        # 如果清理后一方为空，则直接判定不一致
        if not front_clean or not table_clean:
            return False, front_val, table_val, "内容为空"
        # 1. 优先包含匹配
        if front_clean in table_clean or table_clean in front_clean:
            return True, front_val, table_val, "内容包含匹配"
        # 2. 相似度匹配（SequenceMatcher）
        ratio = difflib.SequenceMatcher(None, front_clean, table_clean).ratio()
        if ratio >= 0.6:
            return True, front_val, table_val, f"相似度匹配 ({ratio:.2f})"
        else:
            return False, front_val, table_val, f"内容不匹配 (相似度 {ratio:.2f})"
 

    # 质量要求等字段：归一化后精确比较（忽略标点、空格、大小写）
    norm_front = normalize_text(front_val)
    norm_table = normalize_text(table_val)
    if norm_front == norm_table:
        return True, norm_front, norm_table, "归一化匹配"
    # 最后尝试数值比较（如纯数字）
    front_num = extract_numbers(front_val)
    table_num = extract_numbers(table_val)
    if front_num and table_num and is_numeric_equal(front_num, table_num):
        return True, front_num, table_num, "数值匹配"
    return False, front_val, table_val, "不匹配"
# ============================================
# 解析前部信息（保留原始字段名，修正建设规模提取）
# ============================================
def parse_front_info(text):
    info = {}
    # 采购项目名称
    match = re.search(r'采购项目名称[：:]\s*([^\n]+)', text)
    info["采购项目名称"] = match.group(1).strip() if match else "未找到"
    
    match = re.search(r'采购文件发布时间[：:]\s*([^\n]+)', text)
    if match:
        raw_date = match.group(1).strip()
        info["发文时间"] = re.sub(r'\s+', '', raw_date)  # 去除空格
    else:
        info["发文时间"] = "未找到"

    # 项目地点
    match = re.search(r'项目地点[：:]\s*([^\n]+)', text) or re.search(r'建设地点[：:]\s*([^\n]+)', text)
    info["项目地点"] = match.group(1).strip() if match else "未找到"
    
    # 建设规模
    pattern = r'采购项目概况[：:]\s*(.*?)(?=\n\s*\n|采购范围|$)'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        overview = match.group(1)
        size_match = re.search(r'总建筑面积[^0-9]*(\d+(?:\.\d+)?)\s*([^\d\s]{1,5})', overview)
        if size_match:
            num = size_match.group(1)
            unit = size_match.group(2).strip()
            unit = re.split(r'[，,。\s]', unit)[0]
            info["建设规模"] = f"{num}{unit}" if unit else num
        else:
            info["建设规模"] = "未找到"
    else:
        size_match = re.search(r'总建筑面积[^0-9]*(\d+(?:\.\d+)?)\s*([^\d\s]{1,5})', text)
        if size_match:
            num = size_match.group(1)
            unit = size_match.group(2).strip()
            unit = re.split(r'[，,。\s]', unit)[0]
            info["建设规模"] = f"{num}{unit}" if unit else num
        else:
            info["建设规模"] = "未找到"
    
    # 报价方式
    match = re.search(r'报价方式[：:]\s*([^\n]+)', text) or re.search(r'承包方式[：:]\s*([^\n]+)', text)
    info["报价方式"] = match.group(1).strip() if match else "未找到"
    
    # 质量要求
    match = re.search(r'质量要求[：:]\s*([^\n]+)', text) or re.search(r'质量标准[：:]\s*([^\n]+)', text)
    info["质量要求"] = match.group(1).strip() if match else "未找到"
    
    # 安全文明施工要求
    match = re.search(r'安全文明施工要求[：:]\s*([^\n]+)', text)
    if match:
        info["安全文明施工要求"] = match.group(1).strip()
    else:
        info["安全文明施工要求"] = "未找到"

    # 采购范围
    match = re.search(r'采购范围[：:]\s*([^\n]+(?:[^\n]*\n[^\n]*){0,5}?)(?=\n[^\n]*[：:]|\Z)', text, re.DOTALL)
    info["采购范围"] = match.group(1).strip().replace('\n', ' ')[:200] if match else "未找到"

    
    # 计划工期
    plan_line_match = re.search(r'计划工期[：:]\s*([^\n]+)', text)
    plan_text = plan_line_match.group(1).strip() if plan_line_match else ""
    
    # 定位“2.3”的位置（如果有），限定搜索范围
    end_pos = len(text)
    section_23 = re.search(r'\n\s*2\.3\s', text)
    if section_23:
        end_pos = section_23.start()
    # 在限定范围内搜索“数字+天”（如“90天”、“90日历天”）
    days_match = re.search(r'(\d+)\s*天', text[:end_pos])
    if days_match:
        days = days_match.group(1)
        info["计划工期"] = f"{plan_text}；合同工期总日历天数：{days}天" if plan_text else f"合同工期总日历天数：{days}天"
    else:
        info["计划工期"] = plan_text if plan_text else "未找到"
    
   # 报名时间（例如：4.4 报名时间：2026年05月23日至2026年05月28日）
    match = re.search(r'报名时间[：:]\s*([^\n]+)', text)
    info["报名时间"] = match.group(1).strip() if match else "未找到"
    # 分供商资格要求
    start_pattern = r'3\.1\s+分供商应依法设立且满足如下要求[：:]\s*'
    match_start = re.search(start_pattern, text)
    if match_start:
        start_pos = match_start.end()
        end_match = re.search(r'\n\s*3\.2\s+分供商不得存在', text[start_pos:])
        if end_match:
            end_pos = start_pos + end_match.start()
        else:
            end_pos = len(text)
        full_text = text[start_pos:end_pos].strip()
        full_text = re.sub(r'\n\s*\n', '\n', full_text)
        info["分供商资格要求"] = full_text
    else:
        info["分供商资格要求"] = "未找到"
    
    info["报价及单价总价计价方式"] = info["报价方式"]
    
    # ---------- 辅助函数：在给定文本片段中提取联系人+电话 ----------
    def extract_contact_from_section(section):
        patterns = [
            r'收件联系人[：:]\s*([^\n]+?)\s*电话[：:]\s*(\d{11})',
            r'联系人[：:]\s*([^\n]+?)\s*电话[：:]\s*(\d{11})',
            r'收件联系人[：:]\s*(\d{11})\s*联系电话[：:]\s*([^\n]+)',
            r'([\u4e00-\u9fff]{2,4})\s*[：:]?\s*(1[3-9]\d{9})',
        ]
        for pat in patterns:
            m = re.search(pat, section, re.DOTALL)
            if m:
                groups = m.groups()
                if len(groups) == 2:
                    g1, g2 = groups
                    if re.match(r'1[3-9]\d{9}', g1):
                        phone, name = g1, g2
                    elif re.match(r'1[3-9]\d{9}', g2):
                        phone, name = g2, g1
                    else:
                        name, phone = g1, g2
                    name = re.sub(r'[：:]*', '', name).strip()
                    name_match = re.search(r'[\u4e00-\u9fff]{2,4}', name)
                    if name_match:
                        name = name_match.group(0)
                    return f"{name} {phone}"
        return "未找到"
    
    # ---------- 截止时间、提交地点（从“响应文件的递交”段落提取）----------
    delivery_match = re.search(r'(响应文件的递交\s*\n?)', text)
    if delivery_match:
        start_pos = delivery_match.end()
        snippet = text[start_pos:start_pos+500]
        
        # 截止时间：支持“开启时间”或“开标时间”
        line_match = re.search(
            r'响应文件递交的截止时间[（(](?:开启时间|开标时间)[）)]?[为]?\s*(.+?)(?:，|$)', 
            snippet, 
            re.DOTALL
        )
        if line_match:
            date_str = line_match.group(1).strip()
            dt_match = re.search(r'(\d{4}年\s*\d{1,2}月\s*\d{1,2}日\s*\d{1,2}时\d{1,2}分)', date_str)
            if dt_match:
                dt = re.sub(r'\s+', '', dt_match.group(1))
                info["截止时间"] = dt
            else:
                info["截止时间"] = date_str
        else:
            fallback = re.search(r'截止时间[：:]\s*([^\n]+)', snippet)
            info["截止时间"] = fallback.group(1).strip() if fallback else "未找到"
        
        # 提交地点
        loc_match = re.search(r'地点[为:：]?\s*([\u4e00-\u9fff][^。\n]{10,})', snippet)
        if loc_match:
            info["提交地点"] = loc_match.group(1).strip()
        else:
            addr_match = re.search(r'([\u4e00-\u9fff]{2,}[省市][^。\n]{10,})', snippet)
            info["提交地点"] = addr_match.group(1).strip() if addr_match else "未找到"
    else:
        # 没有“响应文件的递交”标题时的备用
        line_match = re.search(
            r'响应文件递交的截止时间[（(](?:开启时间|开标时间)[）)]?[为]?\s*(.+?)(?:，|$)', 
            text, 
            re.DOTALL
        )
        if line_match:
            date_str = line_match.group(1).strip()
            dt_match = re.search(r'(\d{4}年\s*\d{1,2}月\s*\d{1,2}日\s*\d{1,2}时\d{1,2}分)', date_str)
            if dt_match:
                dt = re.sub(r'\s+', '', dt_match.group(1))
                info["截止时间"] = dt
            else:
                info["截止时间"] = date_str
        else:
            time_match = re.search(r'截止时间[：:]\s*([^\n]+)', text)
            info["截止时间"] = time_match.group(1).strip() if time_match else "未找到"
        
        loc_match = re.search(r'响应文件递交的地点[：:]\s*([^\n]+)', text) or re.search(r'响应文件提交地点[：:]\s*([^\n]+)', text)
        if loc_match:
            info["提交地点"] = loc_match.group(1).strip()
        else:
            info["提交地点"] = "未找到"
    
    # ---------- 收件联系人：独立提取，从“8 联系方式”章节中寻找 ----------
    contact_heading_match = re.search(r'(?:^|\n)\s*(?:\d+\s*)?联系方式\s*\n', text, re.IGNORECASE)
    if contact_heading_match:
        start_pos = contact_heading_match.end()
        contact_section = text[start_pos:start_pos+500]
        info["收件联系人"] = extract_contact_from_section(contact_section)
    else:
        info["收件联系人"] = "未找到"
    
    return info
# ============================================
# 解析手动粘贴的表格（修复变量作用域错误）
# ============================================
def parse_table_from_text(table_text):
    # 清理 Markdown 标记
    def clean_markdown(s):
        s = re.sub(r'\[\[(.*?)\]\](?:\.\{[^\}]+\})?', r'\1', s)
        s = re.sub(r'\[(.*?)\](?:\.\{[^\}]+\})?', r'\1', s)
        s = re.sub(r'\{[^\}]+\}', '', s)
        return s

    raw_text = table_text
    table_text = clean_markdown(table_text)
    lines = table_text.strip().split('\n')
    rows = []
    for line in lines:
        line = line.strip()
        if line.startswith('|'):
            if re.match(r'^[\|\s\-:]+$', line):
                continue
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if len(cells) >= 3:
                rows.append((cells[0], cells[2]))
        elif re.match(r'^\d+\s+', line):
            parts = line.split(None, 2)
            if len(parts) >= 3:
                rows.append((parts[1], parts[2]))
    
    table_dict = {}
    for k, v in rows:
        k_clean = re.sub(r'^\d+\.?\s*', '', k).strip()
        if k_clean:
            table_dict[k_clean] = v

    # 拆分复合字段
    target_key = None
    for key in table_dict.keys():
        if "响应文件提交地点及截止时间" in key:
            target_key = key
            break
    if target_key:
        value = table_dict[target_key]
        extra = split_composite_field(value)
        del table_dict[target_key]
        table_dict.update(extra)

    # 降级提取提交地点和截止时间
    if "提交地点" not in table_dict or table_dict["提交地点"] == "未找到":
        loc_match = re.search(r'提交地点[：:]\s*([^。\n]+)', raw_text)
        if loc_match:
            table_dict["提交地点"] = loc_match.group(1).strip()
    if "截止时间" not in table_dict or table_dict["截止时间"] == "未找到":
        time_match = re.search(r'截止时间[：:]\s*([^。\n]+)', raw_text)
        if time_match:
            table_dict["截止时间"] = time_match.group(1).strip()

    # 统一位置截取“响应人资格要求”
    keywords = ["响应人资格要求", "供应商资格要求"]
    full_text = None
    for kw in keywords:
        start_pos = raw_text.find(kw)
        if start_pos != -1:
            sub_start = start_pos + len(kw)
            if sub_start < len(raw_text) and raw_text[sub_start] in '：:':
                sub_start += 1
            while sub_start < len(raw_text) and raw_text[sub_start] in ' \t':
                sub_start += 1
            match_seq = re.search(r'\n\s*(\d+)[\s\t]', raw_text[sub_start:])
            if match_seq:
                end_pos = sub_start + match_seq.start()
            else:
                end_pos = len(raw_text)
            full_text = raw_text[sub_start:end_pos].strip()
            if full_text:
                table_dict["响应人资格要求"] = full_text  # 统一键名
                break

        # 提取“工期要求”的完整内容（跨行）
    target_key = None
    for key in table_dict.keys():
        if "工期要求" in key:
            target_key = key
            break
    if target_key:
        current_val = table_dict[target_key]
        # 如果当前内容不完整（不包含“合同工期”或长度小于50），则重新提取
        if "合同工期" not in current_val:
            # 在原始文本中查找“工期要求”后的内容，直到下一个数字序号
            start_keyword = "工期要求"
            start_pos = raw_text.find(start_keyword)
            if start_pos != -1:
                sub_start = start_pos + len(start_keyword)
                if sub_start < len(raw_text) and raw_text[sub_start] in '：:':
                    sub_start += 1
                while sub_start < len(raw_text) and raw_text[sub_start] in ' \t':
                    sub_start += 1
                match_seq = re.search(r'\n\s*(\d+)[\s\t]', raw_text[sub_start:])
                if match_seq:
                    end_pos = sub_start + match_seq.start()
                else:
                    end_pos = len(raw_text)
                full_text = raw_text[sub_start:end_pos].strip()
                if full_text and len(full_text) > len(current_val):
                    table_dict[target_key] = full_text

    if "报价以及单价和总价计算方式" not in table_dict:
        for k in table_dict.keys():
            # 匹配“报价方式”或包含“报价”和“单价”的键（如“报价以及单价和总价计算方式”）
            if k == "报价方式" or ("报价" in k and "单价" in k):
                table_dict["报价以及单价和总价计算方式"] = table_dict[k]
                break

    return table_dict
def split_composite_field(value):
    # 您之前写的拆分函数，保持不变
    result = {}
    # 收件联系人
    if '收件人' in value:
        # 先找“收件人”后的内容直到遇到“提交地点”或“截止时间”或结尾
        start = value.find('收件人')
        end = len(value)
        if '提交地点' in value:
            end = min(end, value.find('提交地点'))
        if '截止时间' in value:
            end = min(end, value.find('截止时间'))
        part = value[start:end]
        # 提取手机号（11位数字）
        phone_match = re.search(r'(\d{11})', part)
        # 提取姓名：在“收件人”之后、手机号之前的中文（2-4个汉字）
        name_match = re.search(r'收件人[：:]\s*([\u4e00-\u9fff]{2,4})', part)
        if phone_match and name_match:
            result['收件联系人'] = f"{name_match.group(1)} {phone_match.group(1)}"
        elif phone_match:
            result['收件联系人'] = phone_match.group(1)
        elif name_match:
            result['收件联系人'] = name_match.group(1)
        else:
            # 保底：取“收件人”后到下一个关键词的内容
            m = re.search(r'收件人[：:]\s*(.+?)(?=提交地点|截止时间|$)', part)
            if m:
                result['收件联系人'] = m.group(1).strip()
    # 提交地点
    if '提交地点' in value:
        start = value.find('提交地点')
        end = len(value)
        if '截止时间' in value:
            end = min(end, value.find('截止时间'))
        part = value[start:end]
        m = re.search(r'提交地点[：:]\s*(.+?)(?=截止时间|$)', part)
        if m:
            result['提交地点'] = m.group(1).strip()
    # 截止时间
    if '截止时间' in value:
        start = value.find('截止时间')
        part = value[start:]
        m = re.search(r'截止时间[：:]\s*(.+?)(?=$)', part)
        if m:
            result['截止时间'] = m.group(1).strip()
    return result
# ============================================
# 比对映射（使用智能比较）
# ============================================
def compare_with_rules(front_info, table_dict):
    mapping = [
        ("采购项目名称", "工程名称"),
        ("项目地点", "建设地点"),
        ("建设规模", "建设规模"),
        # ("报价方式", "承包方式"),   # 删除该项，不再比对
        ("质量要求", "质量标准"),
        ("采购范围", "采购范围"),
        ("计划工期", "工期要求"),
        ("分供商资格要求", "响应人资格要求"),
        ("报价及单价总价计价方式", "报价以及单价和总价计算方式"),
        ("提交地点", "提交地点"),
        ("截止时间", "截止时间"),
        ("收件联系人", "收件联系人"),
    ]
    results = []
    for front_key, table_key in mapping:
        front_val = front_info.get(front_key, "")
        # 特殊处理：报价方式（前部）与须知表的报价方式或报价以及单价和总价计算方式比较
        if front_key == "报价方式":
            # 先从 table_dict 中查找“报价方式”，如果没有则查找“报价以及单价和总价计算方式”
            table_val = table_dict.get("报价方式") or table_dict.get("报价以及单价和总价计算方式")
            if table_val is None:
                # 如果都没有，则标记缺失
                results.append({
                    "项目": front_key,
                    "文件前部内容": front_val,
                    "须知样表内容": "（未找到对应项）",
                    "状态": "❌ 缺失",
                    "显示值": f"前部: {front_val}\n须知: 未找到"
                })
                continue
        else:
            # 其他字段正常匹配
            table_val = None
            for tk, tv in table_dict.items():
                if tk == table_key or table_key in tk:
                    table_val = tv
                    break
            if table_val is None:
                results.append({
                    "项目": front_key,
                    "文件前部内容": front_val,
                    "须知样表内容": "（未找到对应项）",
                    "状态": "❌ 缺失",
                    "显示值": f"前部: {front_val}\n须知: 未找到"
                })
                continue
        
        # 进行智能比较
        is_match, norm_front, norm_table, reason = smart_compare(front_val, table_val, front_key)
        status = "✅ 一致" if is_match else "❌ 不一致"
        results.append({
            "项目": front_key,
            "文件前部内容": front_val,
            "须知样表内容": table_val,
            "状态": status,
            "显示值": f"前部: {front_val}\n须知: {table_val}\n比较依据: {reason}"
        })
    return results

def check_internal_rules(front_info, table_dict):
    issues = []
    
    # ---------- 规则1：履约担保比例是否 ≥ 10% ----------
    guarantee_text = table_dict.get("履约担保", "")
    if guarantee_text and guarantee_text != "未找到":
        match = re.search(r'(\d+(?:\.\d+)?)\s*%', guarantee_text)
        if match:
            percent = float(match.group(1))
            if percent >= 10:
                issues.append({
                    "规则": "履约担保比例",
                    "通过": True,
                    "原文": guarantee_text,
                    "详情": f"当前比例 {percent}%，符合 ≥10% 的要求"
                })
            else:
                issues.append({
                    "规则": "履约担保比例",
                    "通过": False,
                    "原文": guarantee_text,
                    "详情": f"当前比例 {percent}%，小于 10%，不符合要求"
                })
        else:
            issues.append({
                "规则": "履约担保比例",
                "通过": False,
                "原文": guarantee_text,
                "详情": "未在履约担保条款中找到明确的百分比数值"
            })
    else:
        issues.append({
            "规则": "履约担保比例",
            "通过": False,
            "原文": "未找到",
            "详情": "未找到履约担保条款"
        })
    
    # ---------- 规则2：资格要求中是否存在禁止性条款（灵活匹配）----------
    flexible_rules = [
        # 模式1：营业执照/经营范围、必须/包含、制作/销售 至少两类同时出现
        {
            "groups": [
                r'营业执照|经营范围',
                r'必须|须|应当|要求包含|限定',
                r'制作|销售|产品|服务'
            ],
            "min_matches": 2,
            "desc": "限定营业执照经营范围（如要求涵盖特定产品）"
        },
        # 模式2：限定/指定 + 专利/商标/品牌等 至少两类（实际两组）
        {
            "groups": [
                r'限定|指定|有|持有|拥有',
                r'专利|商标|品牌|原产地|供应商'
            ],
            "min_matches": 2,
            "desc": "限定或指定特定专利、商标、品牌、原产地或供应商"
        },
        # 模式3a：注册资本词 + 金额数字
        {
            "groups": [
                r'注册资本金|注册资本|注册资金',
                r'\d+(?:\.\d+)?\s*(?:万|元|万元以上)'
            ],
            "min_matches": 2,
            "desc": "要求投标人注册资本金在XXX元以上"
        },
        # 模式3b：注册资本词 + 强制性词语
        {
            "groups": [
                r'注册资本金|注册资本|注册资金',
                r'必须|须|至少|不低于'
            ],
            "min_matches": 2,
            "desc": "要求投标人注册资本金（强制性表述）"
        },
    ]
    
    def check_forbidden(text, source_name):
        if not text or text == "未找到":
            return None
        # 按句子分割
        sentences = re.split(r'[。！；\n]+', text)
        violations = []
        for rule in flexible_rules:
            groups = rule["groups"]
            min_matches = rule.get("min_matches", len(groups))
            desc = rule["desc"]
            for sentence in sentences:
                matched_count = 0
                for pattern in groups:
                    if re.search(pattern, sentence, re.IGNORECASE):
                        matched_count += 1
                if matched_count >= min_matches:
                    violations.append(f"{desc} (原文片段: {sentence[:100]})")
                    break  # 一条规则只记录一次
        if violations:
            return f"{source_name}中存在禁止性条款：{'；'.join(violations)}"
        return None
    
    front_qual = front_info.get("分供商资格要求", "")
    table_qual = table_dict.get("响应人资格要求", "")
    front_violation = check_forbidden(front_qual, "前部-分供商资格要求")
    table_violation = check_forbidden(table_qual, "须知表-响应人资格要求")
    
    if front_violation or table_violation:
        details = []
        if front_violation:
            details.append(front_violation)
        if table_violation:
            details.append(table_violation)
        issues.append({
            "规则": "资格要求禁止性条款",
            "通过": False,
            "原文": f"前部资格要求: {front_qual[:200]}...\n须知资格要求: {table_qual[:200]}...",
            "详情": "；".join(details)
        })
    else:
        issues.append({
            "规则": "资格要求禁止性条款",
            "通过": True,
            "原文": f"前部资格要求: {front_qual[:200]}...\n须知资格要求: {table_qual[:200]}...",
            "详情": "未发现明显的禁止性条款"
        })
        # ---------- 规则3：采购方式与报名时间（仅对特定方式检查）----------
        # ---------- 规则3：采购方式与报名时间（仅对特定方式检查）----------
    procurement_method = table_dict.get("采购方式", "")
    register_time = front_info.get("报名时间", "")
    
    # 定义需要检查的采购方式及其最低报名天数
    method_days = {
        "公开询价": 2,
        "竞争性谈判（公开）": 3,
        "邀请招标": 5,
        "公开招标（依法必招）": 7,
        "公开招标（非依法必招）": 5,
    }
    # 定向询价、单一来源、竞争性谈判（邀请）不检查报名时间
    
    if procurement_method and procurement_method != "未找到":
        matched = None
        for method, need_days in method_days.items():
            if method in procurement_method:
                matched = need_days
                break
        if matched is not None:
            # 需要检查报名时间
            if register_time and register_time != "未找到":
                date_pattern = r'(\d{4}年\d{1,2}月\d{1,2}日)'
                dates = re.findall(date_pattern, register_time)
                if len(dates) == 2:
                    start_str, end_str = dates
                    try:
                        start = datetime.strptime(start_str, "%Y年%m月%d日")
                        end = datetime.strptime(end_str, "%Y年%m月%d日")
                        days = (end - start).days
                        if days >= matched:
                            issues.append({
                                "规则": "报名时间",   # 改标题
                                "通过": True,
                                "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                                "详情": f"报名时长为 {days} 天，满足最低要求 {matched} 天"
                            })
                        else:
                            issues.append({
                                "规则": "报名时间",   # 改标题
                                "通过": False,
                                "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                                "详情": f"报名时长仅 {days} 天，低于最低要求 {matched} 天"
                            })
                    except:
                        issues.append({
                            "规则": "报名时间",   # 改标题
                            "通过": False,
                            "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                            "详情": "报名时间格式解析失败"
                        })
                else:
                    issues.append({
                        "规则": "报名时间",   # 改标题
                        "通过": False,
                        "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                        "详情": "报名时间不是有效的起止日期区间"
                    })
            else:
                issues.append({
                    "规则": "报名时间",   # 改标题
                    "通过": False,
                    "原文": f"采购方式：{procurement_method}",
                    "详情": "未找到报名时间"
                })
        else:
            # 采购方式不需要检查报名时间（定向询价、单一来源、竞争性谈判（邀请）等），跳过
            pass
    else:
        issues.append({
            "规则": "报名时间",   # 改标题
            "通过": False,
            "原文": "未找到采购方式",
            "详情": "未在须知表中找到采购方式"
        })
    
    
        # ---------- 规则4：发文时间到响应截止时间的投标时间 ----------
    issue_date = front_info.get("发文时间", "")
    deadline_str = front_info.get("截止时间", "")
    procurement_method = table_dict.get("采购方式", "")
    
    if issue_date != "未找到" and deadline_str != "未找到" and procurement_method and procurement_method != "未找到":
        try:
            # 提取发文时间中的年月日（去除空格）
            issue_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', issue_date)
            if not issue_match:
                raise ValueError(f"发文时间格式错误: {issue_date}")
            iy, im, iday = issue_match.groups()
            issue_clean = f"{iy}年{im}月{iday}日"
            
            # 提取截止时间中的年月日（忽略空格，忽略后面的时、分）
            deadline_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', deadline_str)
            if not deadline_match:
                raise ValueError(f"截止时间未找到日期部分: {deadline_str}")
            dy, dm, dday = deadline_match.groups()
            deadline_clean = f"{dy}年{dm}月{dday}日"
            
            start = datetime.strptime(issue_clean, "%Y年%m月%d日")
            end = datetime.strptime(deadline_clean, "%Y年%m月%d日")
            days = (end - start).days
            
            # 采购方式对应的最低投标时间（天）
            method_min_days = {
                "公开询价": 1,
                "竞争性谈判（公开）": 2,
                "邀请招标": 2,
                "定向询价": 3,
                "竞争性谈判（邀请）": 5,
                "公开招标（依法必招）": 13,
                "公开招标（非依法必招）": 2,
                "单一来源": 1,
            }
            need_days = None
            for method, days_needed in method_min_days.items():
                if method in procurement_method:
                    need_days = days_needed
                    break
            if need_days is not None:
                if days >= need_days:
                    issues.append({
                        "规则": "投标时间",   # 改标题
                        "通过": True,
                        "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
                        "详情": f"投标时间 {days} 天，满足最低要求 {need_days} 天"
                    })
                else:
                    issues.append({
                        "规则": "投标时间",   # 改标题
                        "通过": False,
                        "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
                        "详情": f"投标时间仅 {days} 天，低于最低要求 {need_days} 天"
                    })
            else:
                issues.append({
                    "规则": "投标时间",   # 改标题
                    "通过": False,
                    "原文": f"采购方式：{procurement_method}",
                    "详情": "未知的采购方式，无法判断"
                })
        except Exception as e:
            issues.append({
                "规则": "投标时间",   # 改标题
                "通过": False,
                "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}",
                "详情": f"日期解析失败：{str(e)}"
            })
    else:
        missing = []
        if issue_date == "未找到":
            missing.append("发文时间")
        if deadline_str == "未找到":
            missing.append("截止时间")
        if not procurement_method or procurement_method == "未找到":
            missing.append("采购方式")
        issues.append({
            "规则": "投标时间",   # 改标题
            "通过": False,
            "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
            "详情": f"缺少必要字段：{', '.join(missing)}"
        })
    safety_text = front_info.get("安全文明施工要求", "")
    if safety_text and safety_text != "未找到":
        # 必须包含的零事故表述
        zero_required = [
            (r'死亡率为零|死亡率.*?0|零死亡', "未明确死亡率为零"),
            (r'重伤率为零|重伤率.*?0|零重伤', "未明确重伤率为零"),
            (r'职业病为零|职业病.*?0|零职业病', "未明确职业病为零"),
        ]
        # 必须包含的安全文明管理规定
        rules_required = [
            (r'安全生产.*?管理|安全.*?规定', "未提及符合安全生产管理规定"),
            (r'文明施工.*?管理|文明.*?规定', "未提及符合文明施工管理规定"),
        ]
        missing = []
        for pattern, msg in zero_required:
            if not re.search(pattern, safety_text, re.IGNORECASE):
                missing.append(msg)
        for pattern, msg in rules_required:
            if not re.search(pattern, safety_text, re.IGNORECASE):
                missing.append(msg)
        if missing:
            issues.append({
                "规则": "安全文明施工要求",
                "通过": False,
                "原文": safety_text,
                "详情": "内容不合理：" + "；".join(missing)
            })
        else:
            issues.append({
                "规则": "安全文明施工要求",
                "通过": True,
                "原文": safety_text,
                "详情": "内容合理，包含零事故指标及安全文明管理规定"
            })
    else:
        issues.append({
            "规则": "安全文明施工要求",
            "通过": False,
            "原文": "未找到",
            "详情": "未找到安全文明施工要求条款"
        })

         # ---------- 规则6：评标方式（评委人数检查）----------
    judge_key = None
    for k in table_dict.keys():
        if ('评委' in k or '评标' in k or '评审' in k or '委员会' in k) and ('人数' in k or '成员' in k or '组成' in k):
            judge_key = k
            break
    if judge_key:
        judge_text = table_dict[judge_key]
        # 检查是否包含“单数”或“奇数”
        has_odd = re.search(r'单数|奇数', judge_text)
        # 提取数字
        match = re.search(r'(\d+)', judge_text)
        if match and has_odd:
            num = int(match.group(1))
            if num >= 5 and num % 2 == 1:
                issues.append({
                    "规则": "评标方式",   # 改标题
                    "通过": True,
                    "原文": judge_text,
                    "详情": f"评委人数为 {num} 人，符合5人及以上单数的要求"
                })
            else:
                issues.append({
                    "规则": "评标方式",   # 改标题
                    "通过": False,
                    "原文": judge_text,
                    "详情": f"评委人数为 {num} 人，不符合5人及以上单数的要求"
                })
        elif match and not has_odd:
            issues.append({
                "规则": "评标方式",   # 改标题
                "通过": False,
                "原文": judge_text,
                "详情": "未明确写明“单数”或“奇数”"
            })
        else:
            issues.append({
                "规则": "评标方式",   # 改标题
                "通过": False,
                "原文": judge_text,
                "详情": "未找到明确的人数数字"
            })
    else:
        issues.append({
            "规则": "评标方式",   # 改标题
            "通过": False,
            "原文": "未找到",
            "详情": "须知表中未找到评标方式相关条款"
        })
    return issues
# ============================================
# ============================================
# ============================================
# Streamlit UI
# ============================================
st.set_page_config(page_title="招标文件审查助手", layout="wide")
st.title("📑 招标文件审查助手")

# 自定义CSS：仅针对两个特定expander放大标题字体
st.markdown("""
<style>
/* 针对标题为 "🔍 内容合理性检查" 的 expander */
.streamlit-expanderHeader[aria-label="🔍 内容合理性检查"] {
    font-size: 36px !important;
    font-weight: bold !important;
}
/* 针对标题为 "🔎 上下文一致性检查" 的 expander */
.streamlit-expanderHeader[aria-label="🔎 上下文一致性检查"] {
    font-size: 36px !important;
    font-weight: bold !important;
}
</style>
""", unsafe_allow_html=True)

# 文件类型选择（默认为空，不报错但阻止后续操作）
file_type = st.selectbox("请选择文件类型", ["", "询价文件", "响应文件"], index=0, help="请选择文件类型，招标文件将进行完整审查，询价文件功能待开发")

if file_type == "":
    st.warning("请先选择文件类型")
    st.stop()

if file_type == "询价文件":
    st.markdown("上传招标采购文件，系统将进行内容合理性检查和上下文一致性检查。")
    uploaded_file = st.file_uploader("📄 上传文件 (.docx / .pdf / .txt)", type=["docx", "pdf", "txt"])

    # 初始化 session_state
    if "front_info" not in st.session_state:
        st.session_state.front_info = None
    if "table_dict" not in st.session_state:
        st.session_state.table_dict = None
    if "results" not in st.session_state:
        st.session_state.results = None
    if "issues" not in st.session_state:
        st.session_state.issues = None
    if "last_file" not in st.session_state:
        st.session_state.last_file = None

    # 当上传新文件时，重新提取公告信息
    if uploaded_file is not None and st.session_state.last_file != uploaded_file.name:
        with st.spinner("正在提取公告信息..."):
            full_text = extract_contract_text(uploaded_file)
            st.session_state.front_info = parse_front_info(full_text)
            # 清空之前的表格和检查结果
            st.session_state.table_dict = None
            st.session_state.results = None
            st.session_state.issues = None
            st.session_state.last_file = uploaded_file.name

    # 显示公告信息（默认折叠）
    if st.session_state.front_info:
        with st.expander("📌 提取的公告信息", expanded=False):
            st.json(st.session_state.front_info)

    # 手动粘贴须知表内容
    if uploaded_file is not None:
        st.subheader("📋 获取“分供商须知样表”内容")
        use_auto = st.radio("表格来源", ["手动粘贴", "自动提取（仅限原生表格）"], horizontal=True)
        
        if use_auto == "手动粘贴":
            manual_text = st.text_area("请将“分供商须知样表”的表格内容（包括表头及所有数据行）完整粘贴到下方", height=300,
                                       help="从 Word 文档中复制表格，粘贴到这里。请确保包含项目、内容、说明与要求等列。")
            if st.button("📥 解析手动粘贴的表格", key="parse_manual"):
                if manual_text.strip():
                    table_dict = parse_table_from_text(manual_text)
                    if table_dict:
                        st.session_state.table_dict = table_dict
                        st.success(f"成功解析 {len(table_dict)} 项")
                    else:
                        st.error("解析失败，请检查粘贴内容是否包含表格数据。")
                else:
                    st.warning("请先粘贴表格内容")
        else:
            with st.spinner("尝试自动提取表格..."):
                full_text = extract_contract_text(uploaded_file)
                lines = full_text.split('\n')
                start_idx = None
                for i, line in enumerate(lines):
                    if '分供商须知样表' in line:
                        start_idx = i
                        break
                if start_idx is not None:
                    table_lines = []
                    for i in range(start_idx + 1, len(lines)):
                        line = lines[i].strip()
                        if not line:
                            break
                        if line.startswith('|') or re.match(r'^\d+\s+', line):
                            table_lines.append(line)
                        else:
                            break
                    if table_lines:
                        table_dict = parse_table_from_text("\n".join(table_lines))
                        if table_dict:
                            st.session_state.table_dict = table_dict
                            st.success("自动提取成功")
                        else:
                            st.error("自动提取失败，请切换为手动粘贴模式")
                    else:
                        st.error("未找到表格数据，请切换为手动粘贴模式")
                else:
                    st.error("未找到“分供商须知样表”标题，请切换为手动粘贴模式")

    # 显示须知表内容（默认折叠）
    if st.session_state.table_dict:
        with st.expander("📋 须知样表内容", expanded=False):
            st.json(st.session_state.table_dict)

        # 开始检查按钮
        if st.button("🚀 开始检查", type="primary"):
            if st.session_state.front_info:
                issues = check_internal_rules(st.session_state.front_info, st.session_state.table_dict)
                st.session_state.issues = issues
            results = compare_with_rules(st.session_state.front_info, st.session_state.table_dict)
            st.session_state.results = results

    # 内容合理性检查（折叠，使用原生expander，通过CSS放大标题）
        # 内容合理性检查（自定义折叠，箭头固定）
    if st.session_state.get("issues"):
        # 初始化状态
        if "show_content" not in st.session_state:
            st.session_state.show_content = False
        
        col1, col2 = st.columns([0.1, 5])
        with col1:
            arrow = "▼" if st.session_state.show_content else "▶"
            if st.button(arrow, key="content_arrow", help="折叠/展开"):
                st.session_state.show_content = not st.session_state.show_content
        with col2:
            st.markdown("<h2 style='font-size: 28px; margin: 0;'>🔍 内容合理性检查</h2>", unsafe_allow_html=True)
        
        if st.session_state.show_content:
            for issue in st.session_state.issues:
                if issue["通过"]:
                    st.success(f"✅ **{issue['规则']}**：{issue['详情']}")
                else:
                    st.error(f"❌ **{issue['规则']}**：{issue['详情']}")
                with st.expander("查看详情"):
                    st.write(f"**原文**：{issue['原文']}")

    # 上下文一致性检查（自定义折叠，箭头固定）
    if st.session_state.get("results"):
        if "show_consistency" not in st.session_state:
            st.session_state.show_consistency = False
        
        col1, col2 = st.columns([0.1, 5])
        with col1:
            arrow = "▼" if st.session_state.show_consistency else "▶"
            if st.button(arrow, key="consistency_arrow", help="折叠/展开"):
                st.session_state.show_consistency = not st.session_state.show_consistency
        with col2:
            st.markdown("<h2 style='font-size: 28px; margin: 0;'>🔎 上下文一致性检查</h2>", unsafe_allow_html=True)
        
        if st.session_state.show_consistency:
            for res in st.session_state.results:
                if "✅" in res["状态"]:
                    st.success(f"**{res['项目']}**：{res['状态']}")
                else:
                    st.error(f"**{res['项目']}**：{res['状态']}")
                with st.expander("查看详情"):
                    st.text(res["显示值"])

else:  # 响应文件
    st.info("响应文件功能正在开发中，敬请期待...")