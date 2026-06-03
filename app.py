import json
import tempfile
import os
import streamlit as st
import requests
from langchain_openai import ChatOpenAI
from langchain_ollama import ChatOllama

# ============================================================
# 板块 2：内置负面清单规则库
# ============================================================

RULES_COMPILATION = [
    {"id": "编_1", "category": "资格限定", "prohibited_text": "禁止限定营业执照经营范围，如：营业执照范围必须涵盖招标产品的制作或销售。", "legal_basis": "《关于进一步规范招标投标过程中企业经营资质资格审查工作的通知》第一条", "rule_type": "content"},
    {"id": "编_2", "category": "资格限定", "prohibited_text": "禁止限定/指定特定专利、商标、品牌、原产地或者供应商。", "legal_basis": "《招标投标法》第十八条、第二十条；《实施条例》第三十二条", "rule_type": "content"},
    {"id": "编_3", "category": "资格限定", "prohibited_text": "禁止投标人资格中要求投标人注册资本金在XXX元以上。", "legal_basis": "《全国深化“放管服”改革...》清理注册资本金等不合理条件", "rule_type": "content"},
    {"id": "编_4", "category": "资格限定", "prohibited_text": "禁止招标文件中有度身定向招标/不合理条件限制、排斥其他潜在投标人公平竞争。", "legal_basis": "《招标投标法实施条例》第三十二条、第三十三条", "rule_type": "content"},
    {"id": "编_5", "category": "资格限定", "prohibited_text": "禁止劳务分包招标要求劳务分包资质具备级别。", "legal_basis": "《建筑业企业资质标准》施工劳务序列不分类别和等级", "rule_type": "content"},
    {"id": "编_6", "category": "资格限定", "prohibited_text": "禁止分包招标要求的总承包、专业承包资质等级以及资质种类超出《建筑业企业资质标准》2014版的范围要求。", "legal_basis": "《建筑业企业资质标准》业务范围规定", "rule_type": "content"},
    {"id": "编_7", "category": "保证金约定", "prohibited_text": "禁止设立除依法依规设立的投标保证金、履约保证金、工程质量保证金、农民工工资保证金以外的其他保证金。", "legal_basis": "《国务院办公厅关于清理规范工程建设领域保证金的通知》", "rule_type": "content"},
    {"id": "编_8", "category": "保证金约定", "prohibited_text": "禁止出现无息退还投标保证金及退还投标保证金时间不合规。", "legal_basis": "《招标投标法实施条例》第五十七条、第六十六条", "rule_type": "content"},
    {"id": "编_9", "category": "保证金约定", "prohibited_text": "禁止以不合理的理由没收投标保证金。", "legal_basis": "《招标投标法实施条例》关于可以不退还保证金的情形", "rule_type": "content"},
    {"id": "编_10", "category": "保证金约定", "prohibited_text": "禁止投标保证金超限额（超过项目估算价的2%或最高80万元）。", "legal_basis": "《招标投标法实施条例》第二十六条；《工程建设项目货物招标投标办法》第二十七条", "rule_type": "quantitative"},
    {"id": "编_11", "category": "保证金约定", "prohibited_text": "禁止出现质量保证金超3%的条款。", "legal_basis": "《建设工程质量保证金管理办法》第七条", "rule_type": "quantitative"},
    {"id": "编_12", "category": "招投标期限", "prohibited_text": "禁止招标文件发售期少于5日。", "legal_basis": "《招标投标法实施条例》第十六条", "rule_type": "quantitative"},
    {"id": "编_13", "category": "招投标期限", "prohibited_text": "禁止招标投标截止日期少于7日，专业分包和设备招标采购原则上最短不应少于10日。", "legal_basis": "《招标投标法》第二十四条；中国机械工业建设集团内部细则", "rule_type": "quantitative"},
    {"id": "编_14", "category": "招投标期限", "prohibited_text": "禁止依法应招必招项目招标投标截止日期少于20日。", "legal_basis": "《招标投标法》第二十四条", "rule_type": "quantitative"},
    {"id": "编_15", "category": "招投标期限", "prohibited_text": "禁止投标截止日期与开标日期不一致，项目投标文件递交截止时间与开标时间存在时间差。", "legal_basis": "《招标投标法》第三十四条", "rule_type": "content"},
    {"id": "编_16", "category": "招标清单", "prohibited_text": "禁止劳务分包招标文件中包含主材、建筑材料款、机械费、周转材料等。", "legal_basis": "《建筑工程施工发包与承包违法行为认定查处管理办法》第十二条", "rule_type": "content"},
    {"id": "编_17", "category": "招标清单", "prohibited_text": "禁止劳务分包招标文件原封不动引用主合同清单。", "legal_basis": "", "rule_type": "content"},
    {"id": "编_18", "category": "招标清单", "prohibited_text": "禁止分包招标文件编制中包含项目主体结构部分的清单。", "legal_basis": "《建筑法》第二十九条", "rule_type": "content"},
    {"id": "编_19", "category": "商务价格", "prohibited_text": "禁止招标文件中规定当单价与数量的乘积与合价、合价与金额累加等不一致时，按就低不就高原则修正，且选择修正报价与投标报价中低价作为中标价条款。", "legal_basis": "《招标投标法实施条例》第五十二条；《评标委员会和评标办法暂行规定》第十九条", "rule_type": "content"},
    {"id": "编_20", "category": "招标程序", "prohibited_text": "禁止招标文件中存在设置开标谈判程序。", "legal_basis": "《招标投标法》第四十三条；《实施条例》第五十七条", "rule_type": "content"},
    {"id": "编_21", "category": "招标程序", "prohibited_text": "禁止在开标后，出现第二轮竞价或第二次谈判或进一步谈判条款。", "legal_basis": "《招标投标法》第四十三条；《实施条例》第五十七条；《工程建设项目施工招标投标办法》第五十九条", "rule_type": "content"},
    {"id": "编_22", "category": "招标程序", "prohibited_text": "禁止无正当理由，随意暂停或终止招标的条款。", "legal_basis": "《工程建设项目施工招标投标办法》第十五条", "rule_type": "content"},
    {"id": "编_23", "category": "其他", "prohibited_text": "禁止要求不合理的签订合同时限（超过中标通知书发出后30日）。", "legal_basis": "《招标投标法》第四十六条", "rule_type": "quantitative"},
    {"id": "编_24", "category": "其他", "prohibited_text": "禁止依法应招必招项目评标委员会组成不合规（非5人以上单数、专家不足2/3等）。", "legal_basis": "《评标委员会和评标方法暂行规定》第八至十二条", "rule_type": "content"},
    {"id": "编_25", "category": "其他", "prohibited_text": "禁止招标文件内容条款前后约定不一致。", "legal_basis": "", "rule_type": "content"},
]

RULES_PROCESS = [
    {"id": "全_1", "category": "招标准备", "prohibited_text": "招标采购与审批、核准、备案的手续倒置", "legal_basis": "", "rule_type": "process"},
    {"id": "全_2", "category": "招标准备", "prohibited_text": "招标采购的标的与审批、核准、备案文件不一致", "legal_basis": "", "rule_type": "process"},
    {"id": "全_3", "category": "招标准备", "prohibited_text": "招标采购公告发布时间晚于投标文件编制和提交时间", "legal_basis": "", "rule_type": "process"},
    {"id": "全_4", "category": "招标环节", "prohibited_text": "在招标公告发布前提前联系潜在投标人进场施工、供货或服务", "legal_basis": "", "rule_type": "process"},
    {"id": "全_5", "category": "招标环节", "prohibited_text": "中标候选公示前，以战略合作伙伴或招商引资为由提前确定中标人", "legal_basis": "", "rule_type": "process"},
    {"id": "全_6", "category": "招标环节", "prohibited_text": "同一项目类似采购内容划分不同的标段分别采购", "legal_basis": "", "rule_type": "process"},
    {"id": "全_7", "category": "招标环节", "prohibited_text": "以涉密为由规避招标采购", "legal_basis": "", "rule_type": "process"},
    {"id": "全_8", "category": "招标文件内容", "prohibited_text": "招标文件存在表意不清，前后矛盾等错误，未向所有潜在投标人统一作出公开澄清/修改", "legal_basis": "", "rule_type": "content"},
    {"id": "全_9", "category": "招标文件内容", "prohibited_text": "招标文件资格、技术、商务等条件设置与采购标的相矛盾，或者明显超出实际需要", "legal_basis": "", "rule_type": "content"},
    {"id": "全_10", "category": "招标文件内容", "prohibited_text": "招标文件以特定业绩、荣誉、奖励作为资格审查标准或评标标准", "legal_basis": "", "rule_type": "content"},
    {"id": "全_11", "category": "招标文件内容", "prohibited_text": "招标文件以特定专利、商标、品牌作为资格审查标准或评标标准", "legal_basis": "", "rule_type": "content"},
    {"id": "全_12", "category": "招标文件内容", "prohibited_text": "要求投标人在当地设立分公司、办事处，拥有一定办公面积或人员社保为当地缴纳", "legal_basis": "", "rule_type": "content"},
    {"id": "全_13", "category": "投标环节", "prohibited_text": "开标前，泄露投标人名字、数量等情况", "legal_basis": "", "rule_type": "process"},
    {"id": "全_14", "category": "投标环节", "prohibited_text": "招标人单独组织部分潜在投标人踏勘项目现场", "legal_basis": "", "rule_type": "process"},
    {"id": "全_15", "category": "投标环节", "prohibited_text": "在投标文件提交时间截止后，仍然接受投标文件递交", "legal_basis": "", "rule_type": "process"},
]

ALL_RULES = RULES_COMPILATION + RULES_PROCESS

# ============================================================
# 板块 3：文本提取（MinerU + 降级）
# ============================================================
def extract_text_mineru(uploaded_file, mineru_api_url):
    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name
    try:
        with open(tmp_path, 'rb') as f:
            files = {'file': f}
            response = requests.post(f"{mineru_api_url.rstrip('/')}/file_parse", files=files, data={"return_md": True}, timeout=120)
            response.raise_for_status()
            result = response.json()
            text = result.get("markdown") or result.get("text") or result.get("content") or ""
            return text
    except Exception as e:
        st.error(f"MinerU 提取失败: {e}")
        return ""
    finally:
        os.unlink(tmp_path)

def extract_text_fallback(uploaded_file):
    file_type = uploaded_file.name.split('.')[-1].lower()
    if file_type == "pdf":
        import pypdf
        reader = pypdf.PdfReader(uploaded_file)
        return "\n".join([page.extract_text() for page in reader.pages])
    elif file_type == "docx":
        from docx import Document
        doc = Document(uploaded_file)
        return "\n".join([p.text for p in doc.paragraphs])
    else:
        return uploaded_file.read().decode("utf-8")

def extract_contract_text(uploaded_file, mineru_api_url=None):
    if mineru_api_url and mineru_api_url.strip():
        with st.spinner("使用 MinerU 提取文本（含布局/表格）..."):
            text = extract_text_mineru(uploaded_file, mineru_api_url)
            if text:
                return text
            st.warning("MinerU 失败，降级为基础提取")
    with st.spinner("使用基础方式提取文本..."):
        return extract_text_fallback(uploaded_file)

# ============================================================
# 板块 4：文本分块
# ============================================================
def chunk_text(text, max_chars=1500):
    paragraphs = text.split('\n')
    chunks = []
    current = ""
    for para in paragraphs:
        if len(current) + len(para) + 1 <= max_chars:
            current += para + "\n"
        else:
            if current:
                chunks.append(current.strip())
            current = para + "\n"
    if current:
        chunks.append(current.strip())
    return chunks

# ============================================================
# 板块 5：提示词构造
# ============================================================
def build_content_prompt(chunk, rule):
    return f"""你是一个招标采购合同审查专家。请判断以下合同片段是否违反了【禁止事项】。

【禁止事项】{rule['prohibited_text']}

【法理依据摘要】{rule.get('legal_basis', '无')[:300]}

【合同片段】
{chunk}

请输出 JSON（不要包含其他文字）：
{{
    "violation": true/false,
    "evidence": "如果违规，引用原文最相关的一句话；否则填 '无'",
    "reason": "简要说明判断理由",
    "suggestion": "如果违规，给出修改建议；否则填 '无'"
}}
"""

def build_quantitative_prompt(chunk, rule):
    return f"""你是一个招标采购合同审查专家。请检查以下合同片段中的数值是否违反了量化禁止事项。

【禁止事项】{rule['prohibited_text']}

【合同片段】
{chunk}

请提取所有相关数值（如百分比、金额、天数），并与禁止标准比较。输出 JSON：
{{
    "violation": true/false,
    "extracted_values": {{"value": "提取到的数值", "threshold": "禁止阈值"}},
    "reason": "说明哪个数值超标",
    "suggestion": "建议修改为合规数值"
}}
"""

def build_process_prompt(chunk, rule):
    return f"""你是一个招标采购合规专家。以下规则属于招标程序行为，通常无法仅从合同文本完全判断。请检查片段中是否提及相关行为或承诺。

【程序事项】{rule['prohibited_text']}

【合同片段】
{chunk}

请输出 JSON：
{{
    "violation": "unknown",
    "note": "该规则属于招标过程行为，建议人工核查相关记录",
    "text_mentions": "片段中是否提到了相关事项？如有请列出"
}}
"""

# ============================================================
# 板块 6：LLM 初始化与审查
# ============================================================
def get_llm(model_type, api_key, model_name):
    if model_type == "OpenAI":
        return ChatOpenAI(openai_api_key=api_key, model_name=model_name, temperature=0)
    elif model_type == "Ollama":
        # 注意：使用 ChatOllama 而不是 Ollama
        return ChatOllama(model=model_name, temperature=0)
    return None

def review_rule_on_chunk(llm, chunk, rule):
    if rule["rule_type"] == "quantitative":
        prompt = build_quantitative_prompt(chunk, rule)
    elif rule["rule_type"] == "process":
        prompt = build_process_prompt(chunk, rule)
    else:
        prompt = build_content_prompt(chunk, rule)
    try:
        response = llm.predict(prompt)
        start = response.find('{')
        end = response.rfind('}') + 1
        if start != -1 and end != 0:
            result = json.loads(response[start:end])
        else:
            result = {"violation": False, "reason": "模型输出格式错误"}
        return result
    except Exception as e:
        return {"violation": False, "reason": f"审查异常: {str(e)}"}

# ============================================================
# 板块 7：主审查流程
# ============================================================
def review_contract(contract_text, rules, llm):
    chunks = chunk_text(contract_text)
    violations = []
    total = len(rules) * len(chunks)
    progress_bar = st.progress(0)
    status = st.empty()
    idx = 0
    for rule in rules:
        for chunk_idx, chunk in enumerate(chunks):
            status.text(f"审查中: {rule['prohibited_text'][:60]}... 片段 {chunk_idx+1}/{len(chunks)}")
            result = review_rule_on_chunk(llm, chunk, rule)
            if result.get("violation") is True:
                violations.append({
                    "rule_id": rule["id"],
                    "category": rule["category"],
                    "prohibited_text": rule["prohibited_text"],
                    "legal_basis": rule.get("legal_basis", ""),
                    "evidence": result.get("evidence", ""),
                    "reason": result.get("reason", ""),
                    "suggestion": result.get("suggestion", ""),
                    "chunk_preview": chunk[:200]
                })
            idx += 1
            progress_bar.progress(idx / total)
    status.empty()
    return violations

# ============================================================
# 板块 8：Streamlit UI
# ============================================================
st.set_page_config(page_title="中机建设招标采购合同审查助手", layout="wide")
st.title("📑 中机建设负面清单合同审查")
st.markdown("依据《招标采购文件编制负面清单》及《全过程负面清单》，AI 自动审查合同文本")

# 主区域：文件上传（必须放在按钮之前定义）
uploaded_file = st.file_uploader("📄 上传合同文件 (PDF / DOCX / TXT)", type=["pdf", "docx", "txt"])

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 模型配置")
    model_type = st.radio("选择模型", ["OpenAI", "Ollama"])
    if model_type == "OpenAI":
        api_key = st.text_input("OpenAI API Key", type="password")
        model_name = st.selectbox("模型", ["gpt-3.5-turbo", "gpt-4-turbo-preview"])
    else:
        api_key = None
        model_name = st.text_input("Ollama 模型名", value="llama3", help="例如 llama3, qwen:7b")
    
    st.divider()
    st.header("📋 规则筛选")
    categories = sorted(set(r["category"] for r in ALL_RULES))
    selected_cats = st.multiselect("选择规则分类", categories, default=categories[:3])
    rule_types = st.multiselect("规则类型", ["content", "quantitative", "process"], default=["content", "quantitative"])
    filtered_rules = [r for r in ALL_RULES if r["category"] in selected_cats and r["rule_type"] in rule_types]
    st.info(f"✅ 内置规则总数: {len(ALL_RULES)}  |  当前选中: {len(filtered_rules)} 条")
    
    st.divider()
    st.header("🔧 高级选项")
    mineru_url = st.text_input("MinerU API 地址 (可选)", placeholder="http://127.0.0.1:8000", help="留空则使用基础文本提取")
    
    st.divider()
    if st.button("🚀 开始审查", type="primary", use_container_width=True):
        if not uploaded_file:
            st.warning("请上传合同文件")
            st.stop()
        if not filtered_rules:
            st.warning("请至少选择一条规则分类或类型")
            st.stop()
        if model_type == "OpenAI" and not api_key:
            st.warning("OpenAI 模式需要填写 API Key")
            st.stop()
        
        contract_text = extract_contract_text(uploaded_file, mineru_url)
        if not contract_text:
            st.error("文本提取失败，请检查文件格式")
            st.stop()
        
        llm = get_llm(model_type, api_key, model_name)
        if llm is None:
            st.error("模型初始化失败")
            st.stop()
        
        violations = review_contract(contract_text, filtered_rules, llm)
        st.session_state["violations"] = violations
        st.session_state["done"] = True

# 结果展示
if "done" in st.session_state and st.session_state["done"]:
    violations = st.session_state["violations"]
    if violations:
        st.error(f"⚠️ 发现 {len(violations)} 处违规")
        for i, v in enumerate(violations):
            with st.expander(f"违规 {i+1}: {v['prohibited_text'][:80]}..."):
                st.markdown(f"**📌 禁止事项**：{v['prohibited_text']}")
                st.markdown(f"**🔍 证据原文**：\n> {v['evidence']}")
                st.markdown(f"**💡 判断理由**：{v['reason']}")
                if v.get('legal_basis'):
                    st.markdown(f"**⚖️ 法理依据**：{v['legal_basis'][:300]}{'...' if len(v['legal_basis'])>300 else ''}")
                st.markdown(f"**✏️ 修改建议**：{v['suggestion']}")
                with st.expander("查看合同相关片段"):
                    st.code(v['chunk_preview'], language="text")
    else:
        st.success("✅ 未发现任何违规，合同符合负面清单要求")
    
    report_json = json.dumps(violations, ensure_ascii=False, indent=2)
    st.download_button("📥 导出审查报告 (JSON)", report_json, file_name="review_report.json", mime="application/json")
    
    if uploaded_file is not None and st.session_state.get("last_file") != uploaded_file.name:
        st.session_state["done"] = False
        st.session_state["last_file"] = uploaded_file.name
    st.download_button("📥 导出审查报告 (JSON)", report_json, file_name="review_report.json", mime="application/json")
    
    if uploaded_file is not None and st.session_state.get("last_file") != uploaded_file.name:
        st.session_state["done"] = False
        st.session_state["last_file"] = uploaded_file.name