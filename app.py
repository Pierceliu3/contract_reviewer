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
# 辅助函数
# ============================================
def normalize_text(s):
    if not isinstance(s, str):
        return ""
    s = re.sub(r'[^\w\u4e00-\u9fff\s]', '', s)
    s = re.sub(r'\s+', ' ', s).strip().lower()
    return s

def extract_numbers(s):
    match = re.search(r'(\d+(?:\.\d+)?)', s)
    return match.group(1) if match else None

def is_numeric_equal(num_str1, num_str2):
    if num_str1 is None or num_str2 is None:
        return False
    try:
        f1 = float(num_str1)
        f2 = float(num_str2)
        if f1 != f2:
            return False
        def has_nonzero_decimal(s):
            if '.' not in s:
                return False
            decimal_part = s.split('.')[1].rstrip('0')
            return len(decimal_part) > 0 and any(c != '0' for c in decimal_part)
        has1 = has_nonzero_decimal(num_str1)
        has2 = has_nonzero_decimal(num_str2)
        return has1 == has2
    except:
        return False

# ============================================
# 智能比较函数
# ============================================
def smart_compare(front_val, table_val, field_name):
    if front_val == "未找到" or table_val == "未找到":
        return False, front_val, table_val, "缺失数据"
    
    exact_fields = ["提交地点","采购项目名称"]
    if field_name in exact_fields:
        front_clean = front_val.strip()
        table_clean = table_val.strip()
        if front_clean == table_clean:
            return True, front_clean, table_clean, "精确匹配"
        else:
            return False, front_clean, table_clean, "内容不一致"
    
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

    if field_name == "建设规模":
        front_num_str = extract_numbers(front_val)
        table_num_str = extract_numbers(table_val)
        def normalize_unit(unit):
            if not unit:
                return ""
            unit_lower = unit.lower()
            if unit_lower in ['平方米', '㎡', 'm2', 'm²', '平方']:
                return "平方米"
            return unit_lower
        def extract_unit(s):
            match = re.search(r'(\d+(?:\.\d+)?)\s*(平方米|㎡|m2|m²|平方|m)', s)
            if match:
                return match.group(2)
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
    
    if field_name in ["计划工期", "工期要求"]:
        front_days_match = re.search(r'(\d+)\s*(?:日历天|天)', front_val)
        table_days_match = re.search(r'(\d+)\s*(?:日历天|天)', table_val)
        front_days = front_days_match.group(1) if front_days_match else None
        table_days = table_days_match.group(1) if table_days_match else None
        days_ok = is_numeric_equal(front_days, table_days) if front_days and table_days else False
        
        def extract_start_date(s):
            match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]', s)
            if match:
                y, m, d = match.groups()
                return f"{y}年{int(m):02d}月{int(d):02d}日"
            return None
        
        front_start = extract_start_date(front_val)
        table_start = extract_start_date(table_val)
        start_ok = (front_start is not None and table_start is not None and front_start == table_start)
        
        has_days = (front_days is not None and table_days is not None)
        has_start = (front_start is not None and table_start is not None)
        
        if has_days and has_start:
            if days_ok and start_ok:
                return True, front_val, table_val, "天数和开工日期均一致"
            else:
                return False, front_val, table_val, "天数或开工日期不一致"
        elif has_days:
            if days_ok:
                return False, front_val, table_val, "缺少开工日期信息"
            else:
                return False, front_val, table_val, "天数不一致"
        elif has_start:
            if start_ok:
                return False, front_val, table_val, "缺少工期天数信息"
            else:
                return False, front_val, table_val, "开工日期不一致"
        else:
            return False, front_val, table_val, "缺少工期和开工日期信息"

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

    if field_name == "采购范围":
        import difflib
        def clean_scope(s):
            s = re.sub(r'（[^）]*）', '', s)
            s = re.sub(r'\([^)]*\)', '', s)
            s = re.sub(r'以[^。]*为准', '', s)
            s = re.sub(r'根据[^，,。]*[,，]?', '', s)
            s = re.sub(r'详见[^，,。]*[,，]?', '', s)
            s = re.sub(r'具体[^，,。]*[,，]?', '', s)
            s = re.sub(r'包含但不限于', '', s)
            s = re.sub(r'[，,。.、；;:：\s]+', '', s)
            s = re.sub(r'[^a-zA-Z\u4e00-\u9fff]', '', s)
            return s.strip()
        front_clean = clean_scope(front_val)
        table_clean = clean_scope(table_val)
        if not front_clean or not table_clean:
            return False, front_val, table_val, "内容为空"
        if front_clean in table_clean or table_clean in front_clean:
            return True, front_val, table_val, "内容包含匹配"
        ratio = difflib.SequenceMatcher(None, front_clean, table_clean).ratio()
        if ratio >= 0.6:
            return True, front_val, table_val, f"相似度匹配 ({ratio:.2f})"
        else:
            return False, front_val, table_val, f"内容不匹配 (相似度 {ratio:.2f})"

    norm_front = normalize_text(front_val)
    norm_table = normalize_text(table_val)
    if norm_front == norm_table:
        return True, norm_front, norm_table, "归一化匹配"
    front_num = extract_numbers(front_val)
    table_num = extract_numbers(table_val)
    if front_num and table_num and is_numeric_equal(front_num, table_num):
        return True, front_num, table_num, "数值匹配"
    return False, front_val, table_val, "不匹配"

# ============================================
# 解析前部信息（公告）
# ============================================
def parse_front_info(text):
    start_marker = "询价公告"
    end_marker = "目录"
    start_pos = text.find(start_marker)
    end_pos = text.find(end_marker, start_pos) if start_pos != -1 else -1
    if start_pos != -1 and end_pos != -1:
        text = text[start_pos:end_pos]
    info = {}
    # 采购项目名称（优先找“采购项目名称”，找不到则找“工程名称”）
    match = re.search(r'采购项目名称[：:]\s*([^\n]+)', text)
    if match:
        info["采购项目名称"] = match.group(1).strip()
    else:
        match = re.search(r'工程名称[：:]\s*([^\n]+)', text)
        info["采购项目名称"] = match.group(1).strip() if match else "未找到"
    
    # 发文时间：优先取“采购文件发布时间”，没有则取“公告发布时间”
    match = re.search(r'采购文件发布时间[：:]\s*([^\n]+)', text)
    if match:
        raw_date = match.group(1).strip()
        info["发文时间"] = re.sub(r'\s+', '', raw_date)
    else:
        match = re.search(r'公告发布时间[：:]\s*([^\n]+)', text)
        if match:
            raw_date = match.group(1).strip()
            info["发文时间"] = re.sub(r'\s+', '', raw_date)
        else:
            info["发文时间"] = "未找到"

    # 项目地点
    match = re.search(r'项目地点[：:]\s*([^\n]+)', text) or re.search(r'建设地点[：:]\s*([^\n]+)', text)
    info["项目地点"] = match.group(1).strip() if match else "未找到"
    
    # 建设规模：同时匹配“建筑面积”或“总建筑面积”
        # 建设规模：匹配“建筑面积”、“总建筑面积”或“总面积”
    pattern = r'采购项目概况[：:]\s*(.*?)(?=\n\s*\n|采购范围|$)'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        overview = match.group(1)
        size_match = re.search(r'(?:(?:总)?建筑面积|总面积)[^0-9]*(\d+(?:\.\d+)?)\s*([^\d\s]{1,5})', overview)
        if size_match:
            num = size_match.group(1)
            unit = size_match.group(2).strip()
            unit = re.split(r'[，,。\s]', unit)[0]
            info["建设规模"] = f"{num}{unit}" if unit else num
        else:
            info["建设规模"] = "未找到"
    else:
        size_match = re.search(r'(?:(?:总)?建筑面积|总面积)[^0-9]*(\d+(?:\.\d+)?)\s*([^\d\s]{1,5})', text)
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
        # 质量要求：匹配“质量、技术要求”、“质量要求”或“质量标准”等，跨行提取到下一个标题
    match = re.search(r'(?:质量[、，]技术|质量要求|质量标准)[：:]\s*(.*?)(?=\n[^\n]*[：:]|\Z)', text, re.DOTALL)
    if match:
        info["质量要求"] = match.group(1).strip().replace('\n', ' ')[:500]
    else:
        info["质量要求"] = "未找到"
    
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
    end_pos = len(text)
    section_23 = re.search(r'\n\s*2\.3\s', text)
    if section_23:
        end_pos = section_23.start()
    days_match = re.search(r'(\d+)\s*天', text[:end_pos])
    if days_match:
        days = days_match.group(1)
        info["计划工期"] = f"{plan_text}；合同工期总日历天数：{days}天" if plan_text else f"合同工期总日历天数：{days}天"
    else:
        info["计划工期"] = plan_text if plan_text else "未找到"
    
    # 报名时间
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
    
    # 辅助函数提取联系人
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
    
    # 截止时间、提交地点
    delivery_match = re.search(r'(响应文件的递交\s*\n?)', text)
    if delivery_match:
        start_pos = delivery_match.end()
        snippet = text[start_pos:start_pos+500]
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
        
        loc_match = re.search(r'地点[为:：]?\s*([\u4e00-\u9fff][^。\n]{10,})', snippet)
        if loc_match:
            info["提交地点"] = loc_match.group(1).strip()
        else:
            addr_match = re.search(r'([\u4e00-\u9fff]{2,}[省市][^。\n]{10,})', snippet)
            info["提交地点"] = addr_match.group(1).strip() if addr_match else "未找到"
    else:
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
    
    # 收件联系人
    contact_heading_match = re.search(r'(?:^|\n)\s*(?:\d+\s*)?联系方式\s*\n', text, re.IGNORECASE)
    if contact_heading_match:
        start_pos = contact_heading_match.end()
        contact_section = text[start_pos:start_pos+500]
        info["收件联系人"] = extract_contact_from_section(contact_section)
    else:
        info["收件联系人"] = "未找到"
    
    return info

# ============================================
# 解析手动粘贴的表格
# ============================================
def parse_table_from_text(table_text):
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

    # 提取响应人资格要求
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
                table_dict["响应人资格要求"] = full_text
                break

    # 提取工期要求的完整内容
    target_key = None
    for key in table_dict.keys():
        if "工期要求" in key:
            target_key = key
            break
    if target_key:
        current_val = table_dict[target_key]
        if "合同工期" not in current_val:
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

    # 确保报价方式键存在
    if "报价以及单价和总价计算方式" not in table_dict:
        for k in table_dict.keys():
            if k == "报价方式" or ("报价" in k and "单价" in k):
                table_dict["报价以及单价和总价计算方式"] = table_dict[k]
                break

    return table_dict

def split_composite_field(value):
    result = {}
    if '收件人' in value:
        start = value.find('收件人')
        end = len(value)
        if '提交地点' in value:
            end = min(end, value.find('提交地点'))
        if '截止时间' in value:
            end = min(end, value.find('截止时间'))
        part = value[start:end]
        phone_match = re.search(r'(\d{11})', part)
        name_match = re.search(r'收件人[：:]\s*([\u4e00-\u9fff]{2,4})', part)
        if phone_match and name_match:
            result['收件联系人'] = f"{name_match.group(1)} {phone_match.group(1)}"
        elif phone_match:
            result['收件联系人'] = phone_match.group(1)
        elif name_match:
            result['收件联系人'] = name_match.group(1)
        else:
            m = re.search(r'收件人[：:]\s*(.+?)(?=提交地点|截止时间|$)', part)
            if m:
                result['收件联系人'] = m.group(1).strip()
    if '提交地点' in value:
        start = value.find('提交地点')
        end = len(value)
        if '截止时间' in value:
            end = min(end, value.find('截止时间'))
        part = value[start:end]
        m = re.search(r'提交地点[：:]\s*(.+?)(?=截止时间|$)', part)
        if m:
            result['提交地点'] = m.group(1).strip()
    if '截止时间' in value:
        start = value.find('截止时间')
        part = value[start:]
        m = re.search(r'截止时间[：:]\s*(.+?)(?=$)', part)
        if m:
            result['截止时间'] = m.group(1).strip()
    return result

# ============================================
# 比对映射（公告 vs 须知表）
# ============================================
def compare_with_rules(front_info, table_dict):
    mapping = [
        ("采购项目名称", "工程名称"),
        ("项目地点", "建设地点"),
        ("建设规模", "建设规模"),
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
        # 特殊处理报价方式
        if front_key == "报价方式":
            table_val = table_dict.get("报价方式") or table_dict.get("报价以及单价和总价计算方式")
            if table_val is None:
                results.append({
                    "项目": front_key,
                    "文件前部内容": front_val,
                    "须知样表内容": "（未找到对应项）",
                    "状态": "❌ 缺失",
                    "显示值": f"公告: {front_val}\n须知: 未找到"
                })
                continue
        else:
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
                    "显示值": f"公告: {front_val}\n须知: 未找到"
                })
                continue
        
        is_match, norm_front, norm_table, reason = smart_compare(front_val, table_val, front_key)
        status = "✅ 一致" if is_match else "❌ 不一致"
        results.append({
            "项目": front_key,
            "文件前部内容": front_val,
            "须知样表内容": table_val,
            "状态": status,
            "显示值": f"公告: {front_val}\n须知: {table_val}\n比较依据: {reason}"
        })
    return results

# ============================================
# 内容合理性检查
# ============================================
def check_internal_rules(front_info, table_dict):
    issues = []
    
    # 规则：采购方式与报名时间
    procurement_method = table_dict.get("采购方式", "")
    register_time = front_info.get("报名时间", "")
    method_days = {
        "公开询价": 2,
        "竞争性谈判（公开）": 3,
        "邀请招标": 5,
        "公开招标（依法必招）": 7,
        "公开招标（非依法必招）": 5,
    }
    if procurement_method and procurement_method != "未找到":
        matched = None
        for method, need_days in method_days.items():
            if method in procurement_method:
                matched = need_days
                break
        if matched is not None:
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
                                "规则": "报名时间",
                                "通过": True,
                                "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                                "详情": f"报名时长为 {days} 天，满足最低要求 {matched} 天"
                            })
                        else:
                            issues.append({
                                "规则": "报名时间",
                                "通过": False,
                                "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                                "详情": f"报名时长仅 {days} 天，低于最低要求 {matched} 天"
                            })
                    except:
                        issues.append({
                            "规则": "报名时间",
                            "通过": False,
                            "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                            "详情": "报名时间格式解析失败"
                        })
                else:
                    issues.append({
                        "规则": "报名时间",
                        "通过": False,
                        "原文": f"采购方式：{procurement_method}；报名时间：{register_time}",
                        "详情": "报名时间不是有效的起止日期区间"
                    })
            else:
                issues.append({
                    "规则": "报名时间",
                    "通过": False,
                    "原文": f"采购方式：{procurement_method}",
                    "详情": "未找到报名时间"
                })
    else:
        issues.append({
            "规则": "报名时间",
            "通过": False,
            "原文": "未找到采购方式",
            "详情": "未在须知表中找到采购方式"
        })
    
    # 规则：投标时间
    issue_date = front_info.get("发文时间", "")
    deadline_str = front_info.get("截止时间", "")
    procurement_method = table_dict.get("采购方式", "")
    if issue_date != "未找到" and deadline_str != "未找到" and procurement_method and procurement_method != "未找到":
        try:
            issue_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', issue_date)
            if not issue_match:
                raise ValueError(f"发文时间格式错误: {issue_date}")
            iy, im, iday = issue_match.groups()
            issue_clean = f"{iy}年{im}月{iday}日"
            deadline_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', deadline_str)
            if not deadline_match:
                raise ValueError(f"截止时间未找到日期部分: {deadline_str}")
            dy, dm, dday = deadline_match.groups()
            deadline_clean = f"{dy}年{dm}月{dday}日"
            start = datetime.strptime(issue_clean, "%Y年%m月%d日")
            end = datetime.strptime(deadline_clean, "%Y年%m月%d日")
            days = (end - start).days
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
                        "规则": "投标时间",
                        "通过": True,
                        "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
                        "详情": f"投标时间 {days} 天，满足最低要求 {need_days} 天"
                    })
                else:
                    issues.append({
                        "规则": "投标时间",
                        "通过": False,
                        "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
                        "详情": f"投标时间仅 {days} 天，低于最低要求 {need_days} 天"
                    })
            else:
                issues.append({
                    "规则": "投标时间",
                    "通过": False,
                    "原文": f"采购方式：{procurement_method}",
                    "详情": "未知的采购方式，无法判断"
                })
        except Exception as e:
            issues.append({
                "规则": "投标时间",
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
            "规则": "投标时间",
            "通过": False,
            "原文": f"发文时间：{issue_date}，截止时间：{deadline_str}，采购方式：{procurement_method}",
            "详情": f"缺少必要字段：{', '.join(missing)}"
        })
    
    # 规则：截止时间格式完整性检查
    deadline_front = front_info.get("截止时间", "")
    deadline_table = table_dict.get("截止时间", "")
    errors = []
    pattern = r'\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日\s*\d{1,2}\s*时\s*\d{1,2}\s*分'
    if deadline_front != "未找到":
        if not re.search(pattern, deadline_front):
            errors.append("公告中的截止时间缺少年月日时分中的部分成分")
    else:
        errors.append("公告中未找到截止时间")
    if deadline_table != "未找到":
        if not re.search(pattern, deadline_table):
            errors.append("须知表中的截止时间缺少年月日时分中的部分成分")
    else:
        errors.append("须知表中未找到截止时间")
    if errors:
        issues.append({
            "规则": "截止时间格式完整性",
            "通过": False,
            "原文": f"公告: {deadline_front}\n须知: {deadline_table}",
            "详情": "；".join(errors)
        })
    else:
        issues.append({
            "规则": "截止时间格式完整性",
            "通过": True,
            "原文": f"公告: {deadline_front}\n须知: {deadline_table}",
            "详情": "截止时间均包含完整的年月日时分"
        })

    # 规则：安全文明施工要求
    safety_text = front_info.get("安全文明施工要求", "")
    if safety_text and safety_text != "未找到":
        zero_required = [
            (r'死亡率为零|死亡率.*?0|零死亡', "未明确死亡率为零"),
            (r'重伤率为零|重伤率.*?0|零重伤', "未明确重伤率为零"),
            (r'职业病为零|职业病.*?0|零职业病', "未明确职业病为零"),
        ]
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

    # 规则：评标方式
    judge_key = None
    for k in table_dict.keys():
        if ('评委' in k or '评标' in k or '评审' in k or '委员会' in k) and ('人数' in k or '成员' in k or '组成' in k):
            judge_key = k
            break
    if judge_key:
        judge_text = table_dict[judge_key]
        has_odd = re.search(r'单数|奇数', judge_text)
        match = re.search(r'(\d+)', judge_text)
        if match and has_odd:
            num = int(match.group(1))
            if num >= 5 and num % 2 == 1:
                issues.append({
                    "规则": "评标方式",
                    "通过": True,
                    "原文": judge_text,
                    "详情": f"评委人数为 {num} 人，符合5人及以上单数的要求"
                })
            else:
                issues.append({
                    "规则": "评标方式",
                    "通过": False,
                    "原文": judge_text,
                    "详情": f"评委人数为 {num} 人，不符合5人及以上单数的要求"
                })
        elif match and not has_odd:
            issues.append({
                "规则": "评标方式",
                "通过": False,
                "原文": judge_text,
                "详情": "未明确写明“单数”或“奇数”"
            })
        else:
            issues.append({
                "规则": "评标方式",
                "通过": False,
                "原文": judge_text,
                "详情": "未找到明确的人数数字"
            })
    else:
        issues.append({
            "规则": "评标方式",
            "通过": False,
            "原文": "未找到",
            "详情": "须知表中未找到评标方式相关条款"
        })
    
    return issues

# ============================================
# 负面清单检查
# ============================================
def check_negative_list(front_info, table_dict):
    issues = []
    skip_table = (table_dict == {})

    # ---------- 1. 履约担保比例（仅在未跳过表格时检查）----------
    if not skip_table:
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

    # ---------- 2. 资格要求禁止性条款（全文扫描，始终执行）----------
    flexible_rules = [
        {
            "groups": [
                r'营业执照|经营范围',
                r'必须|须|应当|要求包含|限定',
                r'制作|销售|产品|服务'
            ],
            "min_matches": 3,
            "desc": "限定营业执照经营范围（如要求涵盖特定产品）",
            "keywords": ['营业执照', '经营范围', '必须', '须', '应当', '要求包含', '限定', '制作', '销售', '产品', '服务']
        },
        {
            "groups": [
                r'限定|指定|使用|持有|拥有',
                r'专利|商标|品牌|原产地|供应商'
            ],
            "min_matches": 2,
            "desc": "限定或指定特定专利、商标、品牌、原产地或供应商",
            "keywords": ['限定', '指定', '使用', '持有', '拥有', '专利', '商标', '品牌', '原产地', '供应商']
        },
        {
            "groups": [
                r'注册资本金|注册资本|注册资金',
                r'\d+(?:\.\d+)?\s*(?:万|元|万元以上)'
            ],
            "min_matches": 2,
            "desc": "要求投标人注册资本金在XXX元以上",
            "keywords": ['注册资本金', '注册资本', '注册资金', '万', '元']
        },
        {
            "groups": [
                r'注册资本金|注册资本|注册资金',
                r'必须|须|至少|不低于'
            ],
            "min_matches": 2,
            "desc": "要求投标人注册资本金（强制性表述）",
            "keywords": ['注册资本金', '注册资本', '注册资金', '必须', '须', '至少', '不低于']
        },
        {
            "groups": [
                r'(?:必须|须|应当|要求|仅限于|限定)\s*(?:在)?\s*(?:本地|当地|本市|本省|本区|注册地)',
                r'投标人.*?(?:本地|当地|本市|本省|本区|注册地)',
                r'注册地\s*(?:必须|须|应当|要求)'
            ],
            "min_matches": 1,
            "desc": "存在地域限制或注册地要求（可能排斥其他地区潜在投标人）",
            "keywords": ['必须在','须在','本地', '当地', '本市', '本省', '本区', '注册地']
        },
    ]
    
    def collect_all_violations(text):
        if not text:
            return []
        sentences = re.split(r'[。！；\n]+', text)
        violations = []
        for rule in flexible_rules:
            desc = rule["desc"]
            groups = rule["groups"]
            keywords = rule.get("keywords", [])
            min_matches = rule.get("min_matches", len(groups))
            for sentence in sentences:
                matched_count = 0
                for pattern in groups:
                    if re.search(pattern, sentence, re.IGNORECASE):
                        matched_count += 1
                if matched_count >= min_matches:
                    violations.append((desc, sentence[:200], keywords))
        return violations
    
    full_text = st.session_state.get("full_text", "")
    if not full_text:
        issues.append({
            "规则": "资格要求禁止性条款",
            "通过": False,
            "原文": "未获取到全文",
            "详情": "无法进行全文扫描，请重新上传文件"
        })
    else:
        violations_list = collect_all_violations(full_text)
        if violations_list:
            details = []
            full_original = []
            for idx, (desc, snippet, keywords) in enumerate(violations_list, 1):
                details.append(f"{idx}. {desc}")
                highlighted = snippet
                for kw in keywords:
                    pattern = re.compile(r'(' + re.escape(kw) + r')', re.IGNORECASE)
                    highlighted = pattern.sub(r'<span style="color:red; font-weight:bold;">\1</span>', highlighted)
                full_original.append(f"【违规{idx} - {desc}】<br>原文：{highlighted}")
            issues.append({
                "规则": "资格要求禁止性条款",
                "通过": False,
                "原文": "<br><br>".join(full_original),
                "详情": "<br>".join(details)
            })
        else:
            issues.append({
                "规则": "资格要求禁止性条款",
                "通过": True,
                "原文": "全文未发现明显的禁止性条款",
                "详情": "未发现禁止性条款"
            })

    # ---------- 3. 质量保证金比例检查（全文扫描，始终执行）----------
    if full_text:
        pattern = r'(?:工程质量保证金|质量保证金)\s*[为:：]?\s*(\d+(?:\.\d+)?)\s*%'
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            percent = float(match.group(1))
            start = max(0, match.start() - 50)
            end = min(len(full_text), match.end() + 200)
            context = full_text[start:end].replace('\n', ' ')
            if percent > 3:
                issues.append({
                    "规则": "质量保证金比例",
                    "通过": False,
                    "原文": context[:300],
                    "详情": f"质量保证金比例为 {percent}%，大于 3%，不符合要求"
                })
            else:
                issues.append({
                    "规则": "质量保证金比例",
                    "通过": True,
                    "原文": context[:300],
                    "详情": f"质量保证金比例为 {percent}%，符合 ≤3% 的要求"
                })
        else:
            sentences = re.split(r'[。！；\n]+', full_text)
            found = False
            for sentence in sentences:
                if ('质保金' in sentence or '质量保证金' in sentence) and '%' in sentence:
                    found = True
                    percent_match = re.search(r'(\d+(?:\.\d+)?)\s*%', sentence)
                    if percent_match:
                        percent = float(percent_match.group(1))
                        if percent > 3:
                            issues.append({
                                "规则": "质量保证金比例",
                                "通过": False,
                                "原文": sentence[:200],
                                "详情": f"质量保证金比例为 {percent}%，大于 3%，不符合要求"
                            })
                        else:
                            issues.append({
                                "规则": "质量保证金比例",
                                "通过": True,
                                "原文": sentence[:200],
                                "详情": f"质量保证金比例为 {percent}%，符合 ≤3% 的要求"
                            })
                    else:
                        issues.append({
                            "规则": "质量保证金比例",
                            "通过": False,
                            "原文": sentence[:200],
                            "详情": "未找到明确的百分比数值"
                        })
                    break
            if not found:
                issues.append({
                    "规则": "质量保证金比例",
                    "通过": False,
                    "原文": "未找到",
                    "详情": "未找到质量保证金或质保金条款"
                })
    else:
        issues.append({
            "规则": "质量保证金比例",
            "通过": False,
            "原文": "未获取到全文",
            "详情": "无法进行全文扫描，请重新上传文件"
        })

    # ---------- 4. 投标保证金比例/金额限制检查（全文扫描，始终执行）----------
    if full_text:
        sentences = re.split(r'[。！；\n]+', full_text)
        found = False
        for sentence in sentences:
            if '投标保证金' in sentence:
                found = True
                percent_match = re.search(r'(\d+(?:\.\d+)?)\s*%', sentence)
                if percent_match:
                    percent = float(percent_match.group(1))
                    if percent > 2:
                        issues.append({
                            "规则": "投标保证金比例",
                            "通过": False,
                            "原文": sentence[:200],
                            "详情": f"投标保证金比例 {percent}%，超过法定上限 2%"
                        })
                    else:
                        issues.append({
                            "规则": "投标保证金比例",
                            "通过": True,
                            "原文": sentence[:200],
                            "详情": f"投标保证金比例 {percent}%，符合 ≤2% 的要求"
                        })
                amount_match = re.search(r'(\d+(?:\.\d+)?)\s*(万?元?|万元)', sentence)
                if amount_match:
                    num = float(amount_match.group(1))
                    unit = amount_match.group(2)
                    if '万' in unit:
                        amount = num * 10000
                    else:
                        amount = num
                    if amount > 800000:
                        issues.append({
                            "规则": "投标保证金金额",
                            "通过": False,
                            "原文": sentence[:200],
                            "详情": f"投标保证金金额 {amount/10000:.0f}万元，超过法定上限 80万元"
                        })
                    else:
                        issues.append({
                            "规则": "投标保证金金额",
                            "通过": True,
                            "原文": sentence[:200],
                            "详情": f"投标保证金金额 {amount/10000:.0f}万元，符合 ≤80万元 的要求"
                        })
                if percent_match is None and amount_match is None:
                    issues.append({
                        "规则": "投标保证金条款",
                        "通过": False,
                        "原文": sentence[:200],
                        "详情": "未找到明确的百分比或金额数字"
                    })
                break
        if not found:
            issues.append({
                "规则": "投标保证金条款",
                "通过": True,
                "原文": "未找到",
                "详情": "未找到投标保证金相关条款"
            })
    else:
        issues.append({
            "规则": "投标保证金条款",
            "通过": False,
            "原文": "未获取到全文",
            "详情": "无法进行全文扫描，请重新上传文件"
        })

    # ---------- 5. 评标方式（仅在未跳过表格时检查）----------
    if not skip_table:
        judge_key = None
        for k in table_dict.keys():
            if ('评委' in k or '评标' in k or '评审' in k or '委员会' in k) and ('人数' in k or '成员' in k or '组成' in k):
                judge_key = k
                break
        if judge_key:
            judge_text = table_dict[judge_key]
            has_odd = re.search(r'单数|奇数', judge_text)
            match = re.search(r'(\d+)', judge_text)
            if match and has_odd:
                num = int(match.group(1))
                if num >= 5 and num % 2 == 1:
                    issues.append({
                        "规则": "评标方式",
                        "通过": True,
                        "原文": judge_text,
                        "详情": f"评委人数为 {num} 人，符合5人及以上单数的要求"
                    })
                else:
                    issues.append({
                        "规则": "评标方式",
                        "通过": False,
                        "原文": judge_text,
                        "详情": f"评委人数为 {num} 人，不符合5人及以上单数的要求"
                    })
            elif match and not has_odd:
                issues.append({
                    "规则": "评标方式",
                    "通过": False,
                    "原文": judge_text,
                    "详情": "未明确写明“单数”或“奇数”"
                })
            else:
                issues.append({
                    "规则": "评标方式",
                    "通过": False,
                    "原文": judge_text,
                    "详情": "未找到明确的人数数字"
                })
        else:
            issues.append({
                "规则": "评标方式",
                "通过": False,
                "原文": "未找到",
                "详情": "须知表中未找到评标方式相关条款"
            })
    else:
        pass  # 跳过评标方式检查

    return issues

# ============================================
# Streamlit UI
# ============================================
st.set_page_config(page_title="招标文件审查助手", layout="wide")
st.title("📑 招标文件审查助手")

st.markdown("""
<style>
.streamlit-expanderHeader[aria-label="🔍 内容合理性检查"] {
    font-size: 36px !important;
    font-weight: bold !important;
}
.streamlit-expanderHeader[aria-label="🔎 上下文一致性检查"] {
    font-size: 36px !important;
    font-weight: bold !important;
}
</style>
""", unsafe_allow_html=True)

file_type = st.selectbox("请选择文件类型", ["", "询价文件", "响应文件"], index=0, help="请选择文件类型，招标文件将进行完整审查，询价文件功能待开发")

if file_type == "":
    st.warning("请先选择文件类型")
    st.stop()

if file_type == "询价文件":
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
    if "full_text" not in st.session_state:
        st.session_state.full_text = None
    if "negative_issues" not in st.session_state:
        st.session_state.negative_issues = None
    if "skip_table" not in st.session_state:
        st.session_state.skip_table = False

    # 上传文件处理
    if uploaded_file is not None and st.session_state.last_file != uploaded_file.name:
        with st.spinner("正在提取公告信息..."):
            full_text = extract_contract_text(uploaded_file)
            st.session_state.full_text = full_text
            cutoff = full_text.find("分供商须知样表")
            if cutoff != -1:
                front_text = full_text[:cutoff]
            else:
                front_text = full_text
            st.session_state.front_info = parse_front_info(front_text)
            # 清空之前的表格和检查结果
            st.session_state.table_dict = None
            st.session_state.results = None
            st.session_state.issues = None
            st.session_state.negative_issues = None
            st.session_state.last_file = uploaded_file.name

    # 显示公告信息
    if st.session_state.front_info:
        with st.expander("📌 提取的公告信息", expanded=False):
            st.json(st.session_state.front_info)

    # 须知表获取
    if uploaded_file is not None:
        st.subheader("📋 获取“分供商须知样表”内容")
        use_auto = st.radio("表格来源", ["手动粘贴", "自动提取（仅限原生表格）"], horizontal=True)
        
        if use_auto == "手动粘贴":
            if st.session_state.table_dict is None:
                manual_text = st.text_area(
                    "请将“分供商须知样表”的表格内容（包括表头及所有数据行）完整粘贴到下方",
                    height=300,
                    help="从 Word 文档中复制表格，粘贴到这里。请确保包含项目、内容、说明与要求等列。"
                )
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📥 解析表格", key="parse_manual"):
                        if manual_text.strip():
                            table_dict = parse_table_from_text(manual_text)
                            if table_dict:
                                st.session_state.table_dict = table_dict
                                st.success(f"成功解析 {len(table_dict)} 项")
                            else:
                                st.error("解析失败，请检查粘贴内容是否包含表格数据。")
                        else:
                            st.warning("请先粘贴表格内容")
                with col2:
                    if st.button("⏩ 跳过表格", key="skip_table"):
                        st.session_state.table_dict = {}
                        st.success("已跳过须知表，将不进行依赖表格的检查")
            else:
                if st.session_state.table_dict == {}:
                    st.info("已跳过须知表，未提供表格内容")
                else:
                    with st.expander("📋 须知样表内容", expanded=False):
                        st.json(st.session_state.table_dict)
        else:  # 自动提取
            if st.session_state.table_dict is None:
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
                                st.session_state.table_dict = None
                        else:
                            st.error("未找到表格数据，请切换为手动粘贴模式")
                            st.session_state.table_dict = None
                    else:
                        st.error("未找到“分供商须知样表”标题，请切换为手动粘贴模式")
                        st.session_state.table_dict = None
                if st.session_state.table_dict is None:
                    st.warning("自动提取失败，您也可以手动粘贴表格内容：")
                    manual_fallback = st.text_area("请粘贴表格", height=200)
                    if st.button("手动解析", key="parse_fallback"):
                        if manual_fallback.strip():
                            table_dict = parse_table_from_text(manual_fallback)
                            if table_dict:
                                st.session_state.table_dict = table_dict
                                st.success("手动解析成功")
                            else:
                                st.error("解析失败")
            else:
                if st.session_state.table_dict == {}:
                    st.info("已跳过须知表，未提供表格内容")
                else:
                    with st.expander("📋 须知样表内容", expanded=False):
                        st.json(st.session_state.table_dict)

    # 开始检查按钮
    if st.button("🚀 开始检查", type="primary"):
        if st.session_state.front_info:
            skip_table = (st.session_state.table_dict == {})
            st.session_state.skip_table = skip_table
            
            if skip_table:
                st.warning("已跳过须知表，上下文一致性检查将不执行，其他检查照常进行。")
                # 负面清单检查（内部已兼容空字典，会跳过履约担保和评标方式，但全文扫描继续）
                st.session_state.negative_issues = check_negative_list(st.session_state.front_info, st.session_state.table_dict)
                # 内容合理性检查（继续执行，依赖表格的规则会因缺失数据而显示“未找到”）
                issues = check_internal_rules(st.session_state.front_info, st.session_state.table_dict)
                st.session_state.issues = issues
                # 上下文一致性检查跳过，清空结果
                st.session_state.results = []
            else:
                st.session_state.negative_issues = check_negative_list(st.session_state.front_info, st.session_state.table_dict)
                issues = check_internal_rules(st.session_state.front_info, st.session_state.table_dict)
                st.session_state.issues = issues
                st.session_state.results = compare_with_rules(st.session_state.front_info, st.session_state.table_dict)

    # 负面清单检查显示
    if st.session_state.get("negative_issues"):
        if "show_negative" not in st.session_state:
            st.session_state.show_negative = False
        col1, col2 = st.columns([0.1, 5])
        with col1:
            arrow = "▼" if st.session_state.show_negative else "▶"
            if st.button(arrow, key="negative_arrow", help="折叠/展开"):
                st.session_state.show_negative = not st.session_state.show_negative
        with col2:
            st.markdown("<h2 style='font-size: 28px; margin: 0;'>📋 负面清单检查</h2>", unsafe_allow_html=True)
        if st.session_state.show_negative:
            for issue in st.session_state.negative_issues:
                if issue["通过"]:
                    st.success(f"✅ **{issue['规则']}**：{issue['详情']}")
                else:
                    st.error(f"❌ **{issue['规则']}**：{issue['详情']}")
                with st.expander("查看详情"):
                    st.markdown(f"**原文**：<br>{issue['原文']}", unsafe_allow_html=True)

    # 内容合理性检查显示
        # 内容合理性检查显示
    # 内容合理性检查显示
    if st.session_state.get("issues"):
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

    # 上下文一致性检查显示
        # 上下文一致性检查显示
    if st.session_state.get("skip_table"):
        st.info("已跳过须知表，不进行上下文一致性检查")
    elif st.session_state.get("results"):
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