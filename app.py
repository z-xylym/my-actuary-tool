import pdfplumber
import fitz  # PyMuPDF
import json
import re
import io
import time
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from concurrent.futures import ThreadPoolExecutor, as_completed
from openai import OpenAI
import streamlit as st
import requests
from bs4 import BeautifulSoup
import plotly.express as px
import plotly.graph_objects as go

# --- 全局配置：勾稽校验规则 ---
# 说明：'check_type' 为 'cross' 表示跨年，'single' 表示年度内自校
DEFAULT_RULES = [
    # 跨年规则
    {"规则名称": "1. 股东权益期初平衡", "公式": "y25['期初股东权益'] == y24['期末股东权益']", "类型": "cross", "描述": "25期初=24期末"},
    {"规则名称": "2. CSM期初平衡", "公式": "y25['CSM期初余额'] == y24['CSM期末余额']", "类型": "cross", "描述": "25期初=24期末"},
    
    # 年度内规则 (程序会自动为 2024 和 2025 各跑一次)
    {"规则名称": "3. CSM余额勾稽", "公式": "abs(curr['CSM期末余额'] - (curr['CSM期初余额'] + curr['新业务CSM（集团口径）'] + curr['CSM计息'] + curr['CSM调整'] + curr['CSM摊销'] + curr['CSM其他'])) < 5", "类型": "single"},
    {"规则名称": "4. 新业务现值拆分", "公式": "abs(curr['新业务未来现金流入现值'] - (curr['新业务未来现金流入现值（盈利）'] + curr['新业务未来现金流入现值（亏损）'])) < 1", "类型": "single"},
    {"规则名称": "5. 保险服务收入合计", "公式": "abs(curr['保险服务收入合计'] - (curr['未采用保费分配法计量的保险合同保险服务收入'] + curr['采用保费分配法计量的保险合同保险服务收入'])) < 1", "类型": "single"},
    {"规则名称": "6. 非PAA收入拆分", "公式": "abs(curr['未采用保费分配法计量的保险合同保险服务收入'] - (curr['合同服务边际的摊销'] + curr['非金融风险调整的变动'] + curr['预计当期发生的保险服务费用'] + curr['与当期服务或过去服务相关得保费经验调整'] + curr['其他收入调整'] + curr['保险获取现金流的摊销（保险服务收入）'])) < 5", "类型": "single"},
    {"规则名称": "7. 保险服务费用合计", "公式": "abs(curr['保险服务费用合计'] - (curr['未采用保费分配法计量的保险合同保险服务费用'] + curr['采用保费分配法计量的保险合同保险服务费用'])) < 1", "类型": "single"},
    {"规则名称": "8. 非PAA费用拆分", "公式": "abs(curr['未采用保费分配法计量的保险合同保险服务费用'] - (curr['保险获取现金流的摊销（保险服务费用）'] + curr['亏损部分的确认及转回'] + curr['当期发生的赔款及其他相关费用'] + curr['已发生赔款负债相关的履约现金流量变动'])) < 5", "类型": "single"},
    {"规则名称": "9. 获取费用摊销一致性", "公式": "curr['保险获取现金流的摊销（保险服务收入）'] == curr['保险获取现金流的摊销（保险服务费用）']", "类型": "single"},
    {"规则名称": "10. PAA保险业绩", "公式": "abs(curr['采用保费分配法计量的保险合同保险业绩'] - (curr['采用保费分配法计量的保险合同保险服务收入'] - curr['采用保费分配法计量的保险合同保险服务费用'])) < 1", "类型": "single"},
    {"规则名称": "11. 服务收入总表一致性", "公式": "curr['保险服务收入'] == curr['保险服务收入合计']", "类型": "single"},
    {"规则名称": "12. 税前利润勾稽", "公式": "abs(curr['税前利润总额'] - (curr['保险利润'] + curr['投资利润'] + curr['其他利润'])) < 5", "类型": "single"},
    {"规则名称": "13. 分出保费的分摊相等", "公式": "curr['分出保费的分摊'] == curr['分出保费的分摊']", "类型": "single"},
    {"规则名称": "14. 摊回保险服务费用相等", "公式": "curr['减：摊回保险服务费用'] == curr['减：摊回保险服务费用']", "类型": "single"},
    {"规则名称": "15. 承保财务损益相等", "公式": "curr['承保财务损益'] == curr['承保财务损益']", "类型": "single"},
    {"规则名称": "16. 投资利润勾稽", "公式": "abs(curr['投资利润'] - (curr['净投资回报'] + curr['承保财务损益'] + curr['分出再保险财务损益'])) < 5", "类型": "single"},
    {"规则名称": "17. 净利润相等", "公式": "curr['净利润'] == curr['净利润']", "类型": "single"},
    {"规则名称": "18. 综合收益总额相等", "公式": "curr['综合收益总额'] == curr['综合收益总额']", "类型": "single"},
    {"规则名称": "19. 费用分类一致性", "公式": "abs((curr['获取费用'] + curr['维持费用'] + curr['非履约费用']) - (curr['职工薪酬'] + curr['物业及设备支出'] + curr['业务投入及监管费用支出'] + curr['行政办公支出'] + curr['其他支出'])) < 10", "类型": "single"}
]
# ==================== 1. 页面基础配置 (必须在最顶端且仅此一份) ====================
st.set_page_config(
    page_title="DigiLife 寿研数智 | 年报处理平台", 
    page_icon="Digi.png",  # ✨ 修改点：加上 .png 后缀
    layout="wide"
)

# ==================== 2. 设置左上角 Logo ====================
# 注意：确保 Digi.png 文件和 app.py 在同一个文件夹里
try:
    st.logo(
        "Digi.png",       # ✨ 修改点：加上 .png 后缀
        icon_image="Digi.png" # ✨ 修改点：加上 .png 后缀
    )
except Exception as e:
    st.error(f"Logo 加载失败，请检查文件名是否为 Digi.png。错误信息: {e}")


# ==================== 辅助函数：单位嗅探 ====================
def get_report_unit(text):
    """该公司披露年报的金额单位"""
    if not text: return None
    m1 = re.search(r'(?:金额)?单位[：:\s]*((?:人民币)?[元十百千万亿]+)', text)
    if m1: return m1.group(1).strip()
    m2 = re.search(r'(人民币[十百千万亿]+元)', text)
    if m2: return m2.group(1).strip()
    if "人民币元" in text: return "人民币元"
    return None

# ==================== 后台引擎：调用 DeepSeek 提取单页表格 ====================
def extract_single_page(pdf_bytes, page_num, expected_name, api_key):
    """使用 pdfplumber 提取文本，并用 LLM 转换为结构化数据"""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        if page_num < 1 or page_num > len(pdf.pages):
            return None, ""
        page = pdf.pages[page_num - 1]
        raw_text = page.extract_text(layout=True)
        
    if not raw_text:
        return None, ""

    prompt = f"""你是一个四大会计师事务所的资深数字化审计专家。
当前任务：提取【{expected_name}】。
请将以下纯文本中的【段落文字】原样保留为一行，将【财务表格】转化为用 "|" 严格分隔的标准网格格式。
【强制要求】：
1. 所有的科目名、附注编号、数字金额之间，必须被 "|" 隔开。严禁使用连续空格。
2. 精准对齐多级表头！如果某个单元格是空的，必须用 "|" 占位补齐！
3. 绝不能漏掉任何一行数据，保留金额里的逗号和括号。
4. 纯文本输出，不要使用 Markdown 代码块。不需要输出 SHEET_NAME，直接输出转化后的数据。

以下是原始文本，请开始输出：
{raw_text}"""

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.0
    )
    
    result_text = response.choices[0].message.content.strip()
    result_text = re.sub(r'^```(csv|text)?\n?', '', result_text, flags=re.MULTILINE)
    result_text = result_text.replace("```", "").strip()
    
    # 转换为 DataFrame
    lines = result_text.split('\n')
    parsed_data = []
    max_cols = 0
    for row in lines:
        clean_row = row.strip()
        if not clean_row: continue
        # 忽略大模型可能生成的 markdown 分割线
        if re.match(r'^[\s\|-]+$', clean_row) and '-' in clean_row: continue
        
        if clean_row.startswith('|'): clean_row = clean_row[1:]
        if clean_row.endswith('|'): clean_row = clean_row[:-1]
        
        if '|' not in clean_row:
            cols = [clean_row]
        else:
            cols = [col.strip() for col in clean_row.split('|')]
        parsed_data.append(cols)
        max_cols = max(max_cols, len(cols))
        
    for row in parsed_data:
        if len(row) < max_cols:
            row.extend([''] * (max_cols - len(row)))
            
    return pd.DataFrame(parsed_data), raw_text

# ==================== 后台引擎：调用 DeepSeek (混合双引擎) ====================
def ai_find_pages(pdf_bytes, api_key, target_tables):
    """混合双引擎：AI负责读目录推算主表，Python雷达负责深潜寻找附注表"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    # 1. 提取前 60 页（扩大视野，确保覆盖大型公司报表区）
    toc_text = ""
    for i in range(min(60, len(doc))):
        page = doc.load_page(i)
        # 获取文本并去掉多余换行，保留前 1200 个字（通常页眉和目录都在这里）
        page_text = page.get_text("text").replace("\n", " ")
        # 压缩多余空格，节省 AI 的 Token，让它看得更多
        page_text = " ".join(page_text.split())
        toc_text += f"---第{i+1}页---\n{page_text[:1200]}\n"
    # 2. ⚡ 附注专属雷达特征矩阵
    feature_matrix = {
        "保险合同-未采用保费分配法按计量组成部分分析 (原保险)": [
            ["未采用保费分配法", "不采用保费分配法"], 
            ["计量组成部分", "履约现金流量"], 
            ["合同服务边际", "非金融风险调整", "未来现金流量"] 
        ],
        "保险合同-非PAA按未到期责任负债和已发生赔款负债分析 (原保险)": [
            ["未采用保费分配法", "不采用保费分配法"], 
            ["未到期责任负债"], 
            ["已发生赔款负债", "已发生赔款"]
        ],
        "保险合同-PAA按未到期责任负债和已发生赔款负债分析 (原保险)": [
            ["采用保费分配法"], 
            ["未到期责任负债"], 
            ["已发生赔款负债", "已发生赔款"]
        ],
        "保险合同-当期初始确认对资产负债表的影响": [
            ["初始确认的合同", "初始确认的保险合同", "初始确认的影响","如下表所示","如表所示"],  
            ["非亏损合同", "亏损合同"],  
            ["未来现金流出", "未来现金流入", "保险获取现金流量"] 
        ],
        "业务及管理费 (附注)": [
            ["业务及管理费"],
            ["职工薪酬", "折旧", "办公费", "宣传费", "租赁"] # 增加二级科目关键词
        ],
        "评估假设-折现率假设": [
            ["折现率假设", "折现率曲线","精算假设"]
        ]
    }
    
    simple_title_OR = {
        "保险服务收入表 (附注)": ["保险服务收入"],
        "保险服务支出/费用表 (附注)": ["保险服务支出", "保险服务费用"],
        "其他综合收益/损益表 (附注)": ["其他综合收益", "其他综合损益"]
    }
    
    context_anchors = ["项目注释", "财务报表附注", "报表附注", "项目附注"]
    currency_indicators = ["人民币百万元", "人民币千元", "人民币元", "单位：千元", "单位：百万元"]
    time_indicators = ["本年", "上年", "本期", "上期", "期末", "期初", "年初", "年末"]
    
    found_hints = {table: [] for table in target_tables if table in feature_matrix or table in simple_title_OR}
    page_type_map = {} 
    
    for i in range(20, len(doc)):
        page = doc.load_page(i)
        text_raw = page.get_text("text")
        text_clean = text_raw.replace("\n", "").replace(" ", "")
        header_text = text_clean[:800]
        
        is_group = "本集团" in header_text or "合并" in header_text
        is_company = "本公司" in header_text or "母公司" in header_text
        if is_group:
            page_type_map[i+1] = 'group'
        elif is_company:
            page_type_map[i+1] = 'company'
        
        is_in_notes = any(anchor in text_clean for anchor in context_anchors)
        has_currency = any(c in text_clean for c in currency_indicators)
        has_time = any(t in text_clean for t in time_indicators)
        is_continued_table = "(续)" in header_text or "（续）" in header_text
        
        try:
            has_actual_table = len(page.find_tables().tables) > 0
        except:
            has_actual_table = False
            
        is_table_page = is_continued_table or has_actual_table or (has_currency and has_time)
        
        if is_in_notes and is_table_page:
            for table, condition_groups in feature_matrix.items():
                if table in target_tables:
                    if "(原保险)" in table and ("再保险" in header_text or "分出" in header_text):
                        continue
                    if all(any(kw in text_clean for kw in group) for group in condition_groups):
                        if any(kw in header_text for kw in condition_groups[0]) or \
                           (len(condition_groups)>1 and any(kw in header_text for kw in condition_groups[1])):
                            if (i + 1) not in found_hints[table]:
                                found_hints[table].append(i + 1)
                                
            for table, any_of_words in simple_title_OR.items():
                if table in target_tables:
                    if "保险" in table and ("再保险" in header_text or "分出" in header_text):
                        continue
                    if any(kw in header_text for kw in any_of_words):
                        if (i + 1) not in found_hints[table]:
                            found_hints[table].append(i + 1)

    doc.close()
    
    for table, pages in found_hints.items():
        if not pages:
            continue
        group_pages = [p for p in pages if page_type_map.get(p) == 'group']
        company_pages = [p for p in pages if page_type_map.get(p) == 'company']
        
        if group_pages and company_pages:
            found_hints[table] = [p for p in pages if p not in company_pages]
    
    hint_text = "\n\n【⚡ Python程序附注雷达线索】\n"
    for table, pages in found_hints.items():
        if pages:
            pages = sorted(list(set(pages)))
            hint_text += f"- {table} 真实表格物理页码位于: {pages[:4]}\n"
    
    prompt = f"""你是一个顶级的寿险审计专家，任务是定位 PDF 中的核心财务报表页码。

【关键定位逻辑（双模式自适应）】
模式 A：目录驱动（适用于有目录的报表）
- 如果前几页存在“目录”或“Contents”，请识别目标报表对应的页码。
- 计算物理偏移量（封面/说明页占用的页数），得出真实的物理页码。

模式 B：标题驱动（适用于无目录或目录不全的报表，如太平人寿）
- 如果没有明确目录，请直接扫描文本中 `---第N页---` 标记下方的页眉或正文大标题。
- 例如：若在 `---第8页---` 下方看到“合并利润表”，则物理页码即为 8。

【通用执行准则】
1. 🎯 三大主表（资产负债表、利润表、现金流量表）：优先寻找“合并”版本。
2. 🎯 跨页逻辑：由于财务报表通常包含“(续)”，请务必返回所有相关的物理页码数组。
   - 示例：若第 3 页是资产负债表，第 4 页是其负债部分的续表，返回 [3, 4]。
3. 🎯 附注表线索（最高优先级）：对于带有“(附注)”或“保险合同”字样的复杂底表，请【无脑完全信任】下方提供的【Python程序附注雷达线索】，直接返回线索中的页码。
4. 🎯 特别提醒：现金流量表有时在目录中比较隐蔽，请仔细搜寻；若无目录，请在资产负债表和利润表之后的 5-10 页内搜寻其标题。

需求表单列表：
{json.dumps(target_tables, ensure_ascii=False)}

【待扫描文本内容（前 60 页）】：
{toc_text}

【⚡ Python程序附注雷达辅助线索】：
{hint_text}

【强制输出格式】
只输出合法 JSON，键为表单名，值为物理页码数组。找不到填 [0]。
⚠️ 警告：输出的 JSON 的键（Key）必须【完全包含】《需求表单列表》中的所有字段，【绝对不能有任何遗漏】！如果文中真的找不到，请严格对应填写 [0]。
格式示例：{{"合并利润表": [8, 9], "资产负债表 (合并优先)": [2, 3, 4]}}"""

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": "你是一个只输出JSON格式的数据处理机器人。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.1
    )
    
    result_text = response.choices[0].message.content.strip()
    result_text = result_text.replace("```json", "").replace("```", "").strip()
    return json.loads(result_text)

# ==================== 全量 CSS 美化 ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: "PingFang SC", "HarmonyOS Sans", "Microsoft YaHei", "Noto Sans SC", sans-serif !important; color: #4A5568; }
    .stApp { background: linear-gradient(160deg, #F8FAFC 0%, #F1F5F9 50%, #E2E8F0 100%); }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #B0BEC5; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #90A4AE; }
    ::selection { background: #B8C5D6; color: #1E293B; }
    [data-testid="stSidebar"] { background: linear-gradient(180deg, #E2E8F0 0%, #CBD5E1 100%) !important; border-right: 1px solid #94A3B8; }
    [data-testid="stSidebar"] .stTextInput label, [data-testid="stSidebar"] .stMarkdown { color: #334155 !important; font-weight: 500; }
    h1 { color: #111827 !important; font-weight: 600 !important; letter-spacing: 2px; border-bottom: none !important; padding-bottom: 10px; background: transparent !important; }
    h3 { color: #1E293B !important; font-size: 17px !important; font-weight: 600 !important; letter-spacing: 1px; background: rgba(255, 255, 255, 0.45) !important; backdrop-filter: blur(10px); -webkit-backdrop-filter: blur(10px); border: 1px solid rgba(255, 255, 255, 0.8); border-radius: 8px !important; padding: 8px 18px !important; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03), inset 0 1px 0 rgba(255, 255, 255, 0.5) !important; width: fit-content; margin-bottom: 18px !important; margin-top: 10px !important; }
    .stButton > button { font-size: 14px !important; font-weight: 600 !important; letter-spacing: 1.5px; background-color: #94A3B8; color: #FFFFFF; border: 1px solid #64748B; border-radius: 4px; padding: 8px 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06); transition: all 0.2s ease; }
    .stButton > button:hover { background-color: #64748B; border-color: #475569; color: #F8FAFC; transform: translateY(-1px); box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); }
    .stButton > button:active { transform: translateY(1px); box-shadow: none; }
    .stButton > button:disabled { background-color: #CBD5E1 !important; color: #94A3B8 !important; border-color: #CBD5E1 !important; box-shadow: none !important; cursor: not-allowed; }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; background-color: transparent; border-bottom: 2px solid #CBD5E1; }
    .stTabs [data-baseweb="tab"] { font-size: 14px; letter-spacing: 0.5px; background-color: #EDF2F7; color: #64748B; border-radius: 6px 6px 0 0; border: 1px solid #CBD5E1; border-bottom: none; padding: 10px 22px; transition: all 0.15s ease; }
    .stTabs [data-baseweb="tab"]:hover { background-color: #E2E8F0; color: #475569; }
    .stTabs [aria-selected="true"] { background-color: #94A3B8 !important; color: #FFFFFF !important; font-weight: 600 !important; border-color: #64748B !important; }
    .stTextInput > div > div > input, .stNumberInput > div > div > input { background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 4px !important; color: #1E293B !important; font-weight: 500; transition: border-color 0.2s ease; }
    .stTextInput > div > div > input:focus, .stNumberInput > div > div > input:focus { border-color: #64748B !important; box-shadow: 0 0 0 2px rgba(100,116,139,0.15) !important; }
    .stMultiSelect [data-baseweb="tag"] { background-color: #CBD5E1 !important; color: #1E293B !important; border-radius: 4px !important; font-size: 13px !important; }
    [data-testid="stFileUploader"] section { border: 2px dashed #B0BEC5 !important; border-radius: 8px !important; background-color: #FAFBFC !important; transition: border-color 0.2s ease; }
    [data-testid="stFileUploader"] section:hover { border-color: #78909C !important; background-color: #F5F7FA !important; }
    hr { border: none !important; height: 1px !important; background: linear-gradient(to right, transparent, #B0BEC5, transparent) !important; margin: 20px 0 !important; }
    .info-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #94A3B8; border-radius: 6px; padding: 20px 24px; margin: 12px 0; box-shadow: 0 1px 3px rgba(0,0,0,0.04); }
    .info-card h4 { color: #475569; margin: 0 0 8px 0; font-weight: 600; font-size: 15px; }
    .info-card p, .info-card li { color: #64748B; font-size: 14px; line-height: 1.8; }
    .info-card.pink { border-left-color: #D4A5A5; }
    .info-card.green { border-left-color: #A5C4B5; }
    .info-card.blue { border-left-color: #94A3B8; }
    .sidebar-brand { text-align: center; padding: 8px 0 16px 0; border-bottom: 1px solid #B0BEC5; margin-bottom: 20px; }
    .sidebar-brand .logo-text { font-size: 20px; font-weight: 700; color: #475569; letter-spacing: 3px; }
    .sidebar-brand .logo-sub { font-size: 11px; color: #94A3B8; letter-spacing: 4px; margin-top: 4px; }
    .placeholder-section { text-align: center; padding: 40px 20px; color: #94A3B8; }
    .placeholder-section .big-icon { font-size: 48px; margin-bottom: 16px; opacity: 0.6; }
    .placeholder-section p { font-size: 14px; line-height: 2; color: #64748B; }
</style>
""", unsafe_allow_html=True)

# ==================== 侧边栏 ====================
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="logo-text">寿研数智</div>
        <div class="logo-sub">ACTUARIAL INTELLIGENCE</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("#### 密钥中心")
    api_key = st.text_input("DeepSeek API Key", type="password", placeholder="sk-...")
    if not api_key:
        st.warning("请输入您的 API Key")
    else:
        st.success("欢迎回来~")
    st.divider()
    st.markdown("""
    <div style="font-size:12px; color:#64748B; line-height:2.2; letter-spacing:1px;">
    系统版本：v2.0 (Alpha)<br>
    开发者：林友沐
    </div>
    """, unsafe_allow_html=True)

# ==================== 主界面 ====================
st.title("📊 寿研数智・年报处理平台")

# 创建包含 Step 0 的全新导航栏
tab0, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    " 🌐 Step 0 ／ 官网年报监控 ",
    " 📑 Step 1 ／ 智能页码检索 ",
    " ⚡ Step 2 ／ 表格智能转换 ",
    " 📝 Step 3 ／ 目标表标准填报 ",
    " 🔍 Step 4 ／ 数据勾稽检查 ",
    " ⛓️‍💥 Step 5 ／ 多公司数据合并",
    " 📊 Step 6 ／ 可视化分析 "
])

# ─────────── Step 0：官网监控雷达 ───────────
with tab0:
    st.markdown("### 🌐 全网年报监控雷达")
    st.markdown("""
    <div class="info-card green">
        <h4>功能说明</h4>
        <p>上传包含保险公司官网链接的 Excel，系统将自动扫描各个网站，帮你检测最新年份的年报是否已经发布。<b>（避免每天人工刷网页）</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    col_t0_1, col_t0_2 = st.columns([1, 2])
    with col_t0_1:
        target_year = st.number_input("请输入监控年份", min_value=2010, max_value=2050, value=2023, step=1)
    
    with col_t0_2:
        company_file = st.file_uploader("上传公司链接清单 (Excel)", type=["xlsx"])

    if company_file and st.button("🔍 开始全网扫描 (约需几十秒)", use_container_width=True):
        df_links = pd.read_excel(company_file)
        if '公司' not in df_links.columns or '链接地址' not in df_links.columns:
            st.error("Excel中必须包含名为 '公司' 和 '链接地址' 的两列！")
        else:
            results = []
            progress_text = "正在扫描全网年报..."
            my_bar = st.progress(0, text=progress_text)
            total_companies = len(df_links)
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 100.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }

            for index, row in df_links.iterrows():
                company_name = row['公司']
                url = str(row['链接地址'])
                my_bar.progress((index + 1) / total_companies, text=f"正在扫描: {company_name}")
                
                status = "🔴 似乎未更新 / 无法访问"
                if url.endswith('.pdf'):
                    status = "⚠️ 直接PDF链接 (需手动核实)"
                elif "http" in url:
                    try:
                        response = requests.get(url, headers=headers, timeout=5)
                        response.encoding = response.apparent_encoding 
                        soup = BeautifulSoup(response.text, 'html.parser')
                        page_text = soup.get_text()
                        
                        if str(target_year) in page_text:
                            status = "🟢 极可能已更新!"
                    except Exception as e:
                        status = "🟡 网站拦截了扫描，需手动查看"
                else:
                    status = "无效链接"
                
                results.append({
                    "公司名称": company_name,
                    "监控状态": status,
                    "直达链接": url
                })
                time.sleep(0.5) 
                
            my_bar.empty()
            df_result = pd.DataFrame(results)
            st.data_editor(
                df_result,
                column_config={"直达链接": st.column_config.LinkColumn("点击前往下载")},
                hide_index=True,
                use_container_width=True
            )
            st.success("扫描完成！请留意标记为 🟢 的公司，并点击直达链接下载PDF，然后前往 Step 1。")


# ─────────── Step 1：智能页码检索 ───────────
with tab1:
    st.markdown("### 📑 智能页码检索")
    uploaded_file = st.file_uploader(
        "拖拽或选择一份已经下载好的年报 PDF 文件",
        type="pdf",
        help="推荐上传 2024/2025 年寿险公司完整年报"
    )

    if uploaded_file:
        if 'pdf_bytes' not in st.session_state or st.session_state.get('pdf_name') != uploaded_file.name:
            st.session_state['pdf_bytes'] = uploaded_file.read()
            st.session_state['pdf_name'] = uploaded_file.name
            if 'found_pages' in st.session_state:
                del st.session_state['found_pages']
            if 'edited_pages' in st.session_state:
                del st.session_state['edited_pages']
        
        pdf_bytes = st.session_state['pdf_bytes']
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)
        
        st.caption(f"当前加载文件：{uploaded_file.name}　|　文档共 {total_pages} 页")
        st.markdown("---")
        
        col_left, col_spacer, col_right = st.columns([1, 0.05, 1.2])

        with col_left:
            st.markdown("#### 检索目标设定")
            target_tables = st.multiselect(
                "请选择需要定位的报表：",
                [
                    "资产负债表 (合并优先)", "利润表 (合并优先)", "现金流量表 (合并优先)", "股东/所有者权益变动表 (合并优先)",
                    "保险服务收入表 (附注)", "保险服务支出/费用表 (附注)", "其他综合收益/损益表 (附注)", "业务及管理费 (附注)",
                    "保险合同-非PAA按未到期责任负债和已发生赔款负债分析 (原保险)", "保险合同-PAA按未到期责任负债和已发生赔款负债分析 (原保险)",
                    "保险合同-未采用保费分配法按计量组成部分分析 (原保险)", "保险合同-当期初始确认对资产负债表的影响", "评估假设-折现率假设"
                ],
                default=["资产负债表 (合并优先)", "利润表 (合并优先)", "保险合同-非PAA按未到期责任负债和已发生赔款负债分析 (原保险)"]
            )
            
            st.markdown("")  
            btn_search = st.button("启动智能检索", use_container_width=True)
            
            if btn_search:
                if not api_key:
                    st.error("请先在左侧密钥中心输入 API Key。")
                elif not target_tables:
                    st.error("请至少选择一张报表。")
                else:
                    with st.spinner("AI 引擎运转中，正在识别跨页逻辑……"):
                        try:
                            result = ai_find_pages(pdf_bytes, api_key, target_tables)
                            st.session_state['found_pages'] = result
                            st.success("检索完成！请在下方核对物理页码。")
                        except Exception as e:
                            st.error(f"引擎出现异常：{e}")

            if 'found_pages' in st.session_state:
                st.markdown("---")
                st.markdown("#### 结果核对")
                st.caption("提示：若表格跨页，请以英文逗号分隔页码（如 64, 65）。可结合右侧预览进行校准。")
                
                edited_pages = {}
                # ✅ 关键修改点：强制遍历用户选择的 target_tables！
                for table_name in target_tables:
                    # 如果 AI 返回了结果就用 AI 的，如果 AI 偷懒漏掉了，就默认给 [0]
                    page_data = st.session_state['found_pages'].get(table_name, [0])
                    
                    if isinstance(page_data, int):
                        page_data = [page_data]
                    str_val = ", ".join(map(str, page_data))
                    user_input = st.text_input(f"{table_name}", value=str_val, key=f"page_{table_name}")
                    
                    try:
                        edited_pages[table_name] = [int(x.strip()) for x in user_input.split(",") if x.strip().isdigit()]
                    except:
                        edited_pages[table_name] = page_data
                
                st.session_state['edited_pages'] = edited_pages
                
                st.markdown("")
                if st.button("确认页码，进入下一步", use_container_width=True):
                    st.success("页码已确认！请前往 Step 2 进行表格转换。")

        with col_right:
            st.markdown("#### 页面预览")
            if 'found_pages' in st.session_state and 'edited_pages' in st.session_state:
                # ✅ 修改点：换成依赖左侧整理好的完整字典（edited_pages）！
                table_to_view = st.selectbox("选择要预览的报表：", options=list(st.session_state['edited_pages'].keys()))
                
                # 获取该表对应的页码，如果是空，默认给个 [0] 兜底
                pages_to_preview = st.session_state['edited_pages'].get(table_to_view, [0])
                if not pages_to_preview:
                    pages_to_preview = [0]
                
                if len(pages_to_preview) > 1:
                    current_page = st.radio("该报表包含多页，请切换预览：", options=pages_to_preview, horizontal=True)
                else:
                    current_page = pages_to_preview[0]
                
                preview_idx = current_page - 1
                
                if 0 <= preview_idx < total_pages:
                    page = doc.load_page(preview_idx)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    st.image(pix.tobytes("png"), caption=f"当前预览：第 {current_page} 页 / 共 {total_pages} 页", use_column_width=True)
                elif current_page == 0:
                    st.info("尚未识别到页码，请在左侧输入框中手动填入。")
                else:
                    st.warning("该页码超出了文档总页数，请检查左侧修改。")
            else:
                st.markdown("")
                st.info("上传文件并启动检索后，此处将显示对应页面。")
        doc.close()


# ─────────── Step 2：表格智能转换 ───────────
with tab2:
    st.markdown("### ⚡ 表格智能转换")
    st.markdown("""
    <div class="info-card blue">
        <h4>功能说明</h4>
        <p>系统将调用 DeepSeek 对多页 PDF 进行网格化对齐重构，并自动拼接同表跨页数据，生成标准化 Excel。</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("")
    
    if 'edited_pages' not in st.session_state or 'pdf_bytes' not in st.session_state:
        st.warning("⚠️ 请先在 Step 1 中上传文件并完成【页码确认】。")
        st.button("开始提取结构化数据", disabled=True, use_container_width=True)
    elif not api_key:
        st.error("⚠️ 请先在左侧输入 API Key。")
        st.button("开始提取结构化数据", disabled=True, use_container_width=True)
    else:
        valid_tasks = {k: v for k, v in st.session_state['edited_pages'].items() if v != [0]}
        
        if not valid_tasks:
            st.warning("⚠️ 没有找到任何有效的页码，无需提取。")
        else:
            if st.button("🚀 开始极速提取", use_container_width=True):
                extracted_sheets = {}
                global_unit = "未能自动提取，需人工核对"
                all_tasks = []
                for table_name, pages in valid_tasks.items():
                    for page_num in pages:
                        all_tasks.append({"table": table_name, "page": page_num})
                
                total_pages_to_process = len(all_tasks)
                pages_done = 0
                
                progress_bar = st.progress(0)
                status_box = st.status(f"启用并发引擎，共分配了 {total_pages_to_process} 个子任务...", expanded=True)
                temp_results = {table_name: {} for table_name in valid_tasks.keys()}
                
                with status_box:
                    with ThreadPoolExecutor(max_workers=5) as executor:
                        future_to_task = {
                            executor.submit(extract_single_page, st.session_state['pdf_bytes'], task["page"], task["table"], api_key): task for task in all_tasks
                        }
                        
                        for future in as_completed(future_to_task):
                            task = future_to_task[future]
                            t_name = task["table"]
                            p_num = task["page"]
                            try:
                                df, raw_text = future.result()
                                if df is not None and not df.empty:
                                    temp_results[t_name][p_num] = df
                                    st.write(f"✅ [{t_name}] - 第 {p_num} 页提取完成！")
                                    if global_unit == "未能自动提取，需人工核对":
                                        unit = get_report_unit(raw_text)
                                        if unit:
                                            global_unit = unit
                                            st.write(f"&nbsp;&nbsp;&nbsp;&nbsp;🔎 识别到该公司报表单位：【{global_unit}】")
                                else:
                                    st.write(f"⚠️ [{t_name}] - 第 {p_num} 页未提取到有效数据。")
                            except Exception as e:
                                st.error(f"❌ [{t_name}] - 第 {p_num} 页提取失败: {str(e)}")
                                
                            pages_done += 1
                            progress_bar.progress(pages_done / total_pages_to_process)
                
                for table_name, pages in valid_tasks.items():
                    table_dfs = []
                    for p_num in pages: 
                        if p_num in temp_results[table_name]:
                            table_dfs.append(temp_results[table_name][p_num])
                    
                    if table_dfs:
                        merged_df = pd.concat(table_dfs, ignore_index=True)
                        safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', table_name)[:30]
                        extracted_sheets[safe_sheet_name] = merged_df
                            
                status_box.update(label="🎉 所有任务提取完成！", state="complete", expanded=False)
                st.session_state['extracted_data'] = extracted_sheets
                st.session_state['global_unit'] = global_unit

    if 'extracted_data' in st.session_state:
        st.markdown("---")
        st.markdown("#### 📊 提取结果预览")
        extracted_data = st.session_state['extracted_data']
        company_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
        
        unit_df = pd.DataFrame([
            {"项目": "源文件名称", "信息": company_name},
            {"项目": "探测到的报表单位", "信息": st.session_state.get('global_unit', '')}
        ])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            unit_df.to_excel(writer, sheet_name="基本信息_单位", index=False)
            for sheet_name, df in extracted_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        excel_data = output.getvalue()
        
        st.download_button(
            label="⬇️ 一键下载结构化提取表 (Excel)",
            data=excel_data,
            file_name=f"{company_name}_数据提取.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.markdown("")
        sheet_names = list(extracted_data.keys())
        if sheet_names:
            tabs = st.tabs(sheet_names)
            for i, tab in enumerate(tabs):
                with tab:
                    st.dataframe(extracted_data[sheet_names[i]], use_container_width=True)


# ─────────── Step 3：目标表标准填报 ───────────
with tab3:
    st.markdown("### 📝 目标表标准填报")
    st.markdown("""
    <div class="info-card pink">
        <h4>功能说明</h4>
        <p>上传标准化底稿模板，AI 将自动寻找科目数据填充数值，并自动生成底稿 Excel 计算公式，减少 Copy-Paste 工作量。</p>
    </div>
    """, unsafe_allow_html=True)
    
    template_file = st.file_uploader("请上传目标表模板 (.xlsx)", type="xlsx", key="template_uploader")
    
    if template_file:
        st.success(f"模板文件 {template_file.name} 已加载！")
        
        COL_COMPANY, COL_CATEGORY, COL_FIELD_NAME, COL_FIELD_TYPE, COL_NOTE, COL_RULE = "公司", "类别", "字段名", "字段类型", "注释", "计算规则"
        
        if 'extracted_data' not in st.session_state:
            st.warning("⚠️ 尚未找到提取的数据，请先完成 Step 1 和 Step 2。")
        else:
            if st.button("✨ 启动智能填报与公式生成", use_container_width=True):
                status_box = st.status("正在初始化填报引擎...", expanded=True)
                with status_box:
                    st.write("📄 正在解析模板表结构...")
                    template_df = pd.read_excel(template_file)
                    col_2024 = next((c for c in template_df.columns if '2024' in str(c)), None)
                    col_2025 = next((c for c in template_df.columns if '2025' in str(c)), None)
                    
                    if not col_2024 or not col_2025:
                        st.error("❌ 错误：模板中未找到包含 2024 或 2025 的列头！")
                    else:
                        st.write("🎯 正在准备目标指标清单...")
                        input_items = template_df[template_df[COL_FIELD_TYPE].astype(str).str.strip() == "输入"]
                        ai_target_list = [{"类别": str(r.get(COL_CATEGORY, "")), "标准字段名": str(r.get(COL_FIELD_NAME, "")), "别名参考": str(r.get(COL_NOTE, "")) if pd.notna(r.get(COL_NOTE)) else ""} for _, r in input_items.drop_duplicates(subset=[COL_FIELD_NAME]).iterrows()]
                            
                        st.write("🧠 正在组装财务上下文供 AI 分析...")
                        extracted_data = st.session_state['extracted_data']
                        context_text = "".join([f"\n[表名: {name}]\n{df.to_csv(index=False, sep='|')}\n" for name, df in extracted_data.items()])
                            
                        SYSTEM_PROMPT = """你是一个资深的寿险精算审计专家。任务：将 PDF 提取的财务明细精准填入目标底稿，并严格区分 2024 和 2025 年度。

【通用执行准则：绝对原样提取】
1. ⚠️ 绝对指令：除了下述特殊的“业务及管理费”加总需求外，所有指标必须【原封不动地照抄】原文里的文本！
   - 必须保留括号（表示负数）、逗号、正负号。例如原文是 "(295,992.00)"，必须输出为 "(295,992.00)"。
   - 严禁擅自删除符号、严禁自行进行四舍五入。
2. 别名匹配：若标准名找不到，请查看“别名参考”列，或利用你的精算知识寻找同义词。
3. 报表锁定：提取“新业务相关指标”时，务必在包含“当期确认”、“初始确认”或“新业务价值”字样的表格中查找。
4. 空值处理：若原文为“—”或“无”请输出 "0"。若完全找不到该指标，请输出 null。

【特殊指令：业务及管理费（业管费）科目拆分】
业务及管理费附注表通常包含“小计/明细”和“减项”。请按以下精算逻辑进行识别：
1. 【获取费用】：必须提取“减：与保险合同履约直接相关的支出”下方，明确标注为“计入...保险获取现金流量”或“保险获取”字样的数值。
2. 【维持费用】：必须提取“减：与保险合同履约直接相关的支出”下方，明确标注为“计入保险服务费用”或“维持费用”字样的数值。
3. 【非履约费用】：提取该附注表最底部的“合计”项数值。逻辑上，该合计 = 总支出 - 获取费用 - 维持费用。
4. 【二级明细归类】（若目标字段属于明细科目且无汇总）：
   - 【职工薪酬】：加总 工资, 奖金, 补贴, 社保, 福利, 公积金, 辞退福利等。
   - 【物业及设备】：加总 折旧, 摊销, 租赁费, 物业费, 维修费等。
   - 【业务投入及监管】：加总 宣传费, 招待费, 保障基金, 监管费, 服务费, 咨询审计费。
   - 【行政办公】：加总 差旅费, 办公费, 邮电印刷, 会议费等。

【输出格式】
仅输出合法的 JSON 格式，严禁带有任何 Markdown 标记或文字说明。
格式示例：{"字段名": {"y24": "(1,234.0) + 500.0", "y25": "9,876.54"}}"""
                        st.write("🚀 正在呼叫 DeepSeek 进行语义映射与精准取数...")
                        client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                        try:
                            response = client.chat.completions.create(
                                model="deepseek-chat",
                                messages=[
                                    {"role": "system", "content": SYSTEM_PROMPT},
                                    {"role": "user", "content": f"目标指标:\n{json.dumps(ai_target_list, ensure_ascii=False)}\n\n数据:\n{context_text}"}
                                ],
                                temperature=0.0
                            )
                            ai_res = response.choices[0].message.content.strip()
                            ai_res = re.sub(r'^```(json)?\n?', '', ai_res, flags=re.MULTILINE).replace("```", "").strip()
                            ai_data = json.loads(ai_res)
                            st.write("✅ 数据映射成功！正在回填表格...")
                            
                            working_df = template_df.copy()
                            raw_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
                            company_short = re.sub(r'202\d年.*', '', raw_name).strip()
                            if COL_COMPANY in working_df.columns:
                                working_df[COL_COMPANY] = company_short
                                
                            # --- 兼容“绝对原样”与“公式回放”的回填引擎 ---
                            for idx, row in working_df.iterrows():
                                fname = str(row.get(COL_FIELD_NAME, "")).strip()
                                if fname in ai_data:
                                    item_data = ai_data[fname]
                                    def process_final_val(v):
                                        if v is None or str(v).lower() in ["null", "none", "", "—", "无"]: 
                                            return 0.0
                                        
                                        s = str(v).strip()
                                        
                                        # 内部函数：把财务括号 "(1,234)" 转为可计算的 "-1234"
                                        def to_calc_str(val_str):
                                            tmp = val_str.replace(',', '').strip()
                                            if (tmp.startswith('(') and tmp.endswith(')')) or (tmp.startswith('（') and tmp.endswith('）')):
                                                return "-" + tmp[1:-1].strip()
                                            return tmp

                                        # 情况 1：如果是业务及管理费的加总公式
                                        if "+" in s:
                                            parts = s.split("+")
                                            # 处理每一项的括号和逗号，转化为标准数字字符串
                                            cleaned_parts = [to_calc_str(p) for p in parts]
                                            # 返回 Excel 公式格式：=-100+200
                                            return "=" + "+".join(cleaned_parts)
                                        
                                        # 情况 2：如果是通用项目的原样提取
                                        # 先处理掉括号和逗号看能不能转成数字
                                        num_candidate = to_calc_str(s)
                                        try:
                                            return float(num_candidate)
                                        except:
                                            return s # 如果是文字，保留文字原样

                                    if isinstance(item_data, dict):
                                        val_24 = item_data.get("y24")
                                        val_25 = item_data.get("y25")
                                        working_df.at[idx, col_2024] = process_final_val(val_24)
                                        working_df.at[idx, col_2025] = process_final_val(val_25)
                                    else:
                                        # 如果 AI 返回的是 null 或其他非字典内容，直接填充 0
                                        working_df.at[idx, col_2024] = 0.0
                                        working_df.at[idx, col_2025] = 0.0
                            # --- 🚀 注入 Excel 单元格公式引擎 ---
                            st.write("⚙️ 正在生成可追溯的单元格公式...")
                            temp_buffer = io.BytesIO()
                            working_df.to_excel(temp_buffer, index=False)
                            temp_buffer.seek(0)
                            
                            wb = openpyxl.load_workbook(temp_buffer)
                            ws = wb.active
                            
                            # 获取列索引
                            cols = {cell.value: cell.column for cell in ws[1]}
                            c24_idx = cols.get(col_2024)
                            c25_idx = cols.get(col_2025)
                            
                            # 1. 激活“输入项”中的 AI 相加公式 (例如：=-123+456)
                            for r in range(2, ws.max_row + 1):
                                for c_idx in [c24_idx, c25_idx]:
                                    if c_idx:
                                        cell = ws.cell(row=r, column=c_idx)
                                        # 如果格子里是字符串且以 = 开头，openpyxl 会自动将其识别为公式
                                        if isinstance(cell.value, str) and cell.value.startswith('='):
                                            # 无需特殊操作，重新赋值即可触发 openpyxl 的公式识别
                                            cell.value = cell.value 

                            # 2. 执行原本的针对“计算”类型字段的公式生成逻辑 (例如：A+B 变 A2+B2)
                            # 这里是你原本的代码逻辑...
                            field_to_row = {str(ws.cell(row=r, column=cols[COL_FIELD_NAME]).value).strip(): r 
                                            for r in range(2, ws.max_row + 1) 
                                            if ws.cell(row=r, column=cols[COL_FIELD_NAME]).value}
                            sorted_fields = sorted(field_to_row.keys(), key=len, reverse=True)
                            
                            c24_let = get_column_letter(c24_idx) if c24_idx else ""
                            c25_let = get_column_letter(c25_idx) if c25_idx else ""

                            for r in range(2, ws.max_row + 1):
                                ftype = str(ws.cell(row=r, column=cols[COL_FIELD_TYPE]).value).strip()
                                rule = str(ws.cell(row=r, column=cols[COL_RULE]).value).strip()
                                if ftype == "计算" and rule not in ["nan", "None", ""]:
                                    f24, f25 = rule, rule
                                    for f in sorted_fields:
                                        if f in rule:
                                            row_num = field_to_row[f]
                                            if c24_let: f24 = f24.replace(f, f"{c24_let}{row_num}")
                                            if c25_let: f25 = f25.replace(f, f"{c25_let}{row_num}")
                                    
                                    # 写入计算项公式
                                    try:
                                        if c24_idx: ws.cell(row=r, column=c24_idx).value = f"={f24}"
                                        if c25_idx: ws.cell(row=r, column=c25_idx).value = f"={f25}"
                                    except: pass
                              # 遍历所有数据格，设置格式并确保数字类型正确
                            for r in range(2, ws.max_row + 1):
                                for c_idx in [c24_idx, c25_idx]:
                                    if c_idx:
                                        cell = ws.cell(row=r, column=c_idx)
                                        # 1. 财务格式定义：正数千分位，负数括号，零显示为横杠
                                        cell.number_format = '#,##0_ ;(#,##0);"-"'
                                        
                                        # 2. 数据类型微调：确保数字不是以字符串形式存在
                                        # 如果单元格不是公式且能转成数字，强制转成 float
                                        if cell.value and not str(cell.value).startswith('='):
                                            try:
                                                # 先清洗掉可能存在的干扰字符
                                                clean_val = str(cell.value).replace(',','').replace('(','-').replace(')','')
                                                cell.value = float(clean_val)
                                            except:
                                                pass

                            # 最后保存到最终的二进制流
                            final_buffer = io.BytesIO()
                            wb.save(final_buffer)

                            final_buffer.seek(0)
                            st.session_state['filled_excel'] = final_buffer.getvalue()
                            st.session_state['final_df'] = working_df 
                            status_box.update(label="🎉 填报与公式生成完毕！", state="complete", expanded=False)

                        except Exception as e:
                            st.error(f"❌ 流程中断: {str(e)}")
                            status_box.update(label="处理失败", state="error", expanded=True)

    if 'filled_excel' in st.session_state:
        st.markdown("---")
        st.markdown("#### 📋 智能填报结果预览")
        company_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
        
        st.download_button(
            label="⬇️ 下载已填报的审计底稿 (含公式)",
            data=st.session_state['filled_excel'],
            file_name=f"{company_name}_自动填报表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("💡 *预览表仅展示数值，下载后的 Excel 已包含自动关联公式。*")
        st.dataframe(st.session_state['final_df'], use_container_width=True)


# ─────────── Step 4 分析与校验 ───────────

with tab4:
    col_title, col_btn = st.columns([4, 1])
    with col_title:
        st.subheader("🏁 精算数据勾稽关系检查")
    with col_btn:
        # 💡 增加一个手动刷新按钮，点击后会清空缓存重新运行本页逻辑
        if st.button("🔄 重新执行检查", use_container_width=True):
            st.rerun() 
            
    if 'final_df' in st.session_state:
        df = st.session_state['final_df'].copy()
        
        with st.expander("🛠️ 勾稽规则配置 (点击展开/修改公式)"):
            st.info("可以在此处修改勾稽公式。'curr' 代表当前年度数据，'y24/y25' 代表特定年度。")
            rules_df = st.data_editor(pd.DataFrame(DEFAULT_RULES), num_rows="dynamic", key="rules_editor_v3")

        target_24 = str(st.session_state.get('col_2024', '2024'))
        target_25 = str(st.session_state.get('col_2025', '2025'))
        
        col_24_name = next((col for col in df.columns if target_24 in str(col)), None)
        col_25_name = next((col for col in df.columns if target_25 in str(col)), None)
        
        if not col_24_name or not col_25_name:
            st.error(f"❌ 表格中找不到包含 '{target_24}' 和 '{target_25}' 的列。当前列名: {list(df.columns)}")
            st.stop()

        def parse_finance_num(val):
            if pd.isnull(val): return 0.0
            v_str = str(val).strip()
            if v_str in ['-', '']: return 0.0
            if v_str.startswith('(') and v_str.endswith(')'): v_str = '-' + v_str[1:-1]
            v_str = v_str.replace(',', '')
            try: return float(v_str)
            except ValueError: return 0.0

        # —— 💡 黑科技黑科技：智能容错字典 —— 
        class SmartDict(dict):
            """一个无视空格、全半角冒号/括号的字典"""
            def _clean(self, k):
                if not isinstance(k, str): return k
                # 统一转成英文标点，并删除所有空格、换行
                return k.replace('（','(').replace('）',')').replace('：',':').replace(' ','').replace('\n','').replace('\xa0','')
            
            def __init__(self, mapping):
                super().__init__(mapping)
                # 建立 [清洗后的名字 : 原名] 的映射
                self._map = {self._clean(k): k for k in mapping.keys()}
                
            def __getitem__(self, k):
                clean_k = self._clean(k)
                if clean_k in self._map:
                    return super().__getitem__(self._map[clean_k])
                raise KeyError(k) # 如果连清洗后都找不到，才报错原名
        # ————————————————————————————————

        # 3. 提取数据，并使用 SmartDict 包装！
        base_y24 = {str(k).strip(): parse_finance_num(v) for k, v in zip(df['字段名'], df[col_24_name])}
        base_y25 = {str(k).strip(): parse_finance_num(v) for k, v in zip(df['字段名'], df[col_25_name])}
        
        y24_map = SmartDict(base_y24)
        y25_map = SmartDict(base_y25)
        
        check_results = []

        # 4. 执行校验与智能诊断
        for _, rule in rules_df.iterrows():
            rule_name = rule['规则名称']
            formula = rule['公式']
            rule_type = rule.get('类型', 'single')

            targets = [("2024-2025", None)] if rule_type == "cross" else [("2024", y24_map), ("2025", y25_map)]

            for year_label, curr_map in targets:
                ctx = {'curr': curr_map, 'y24': y24_map, 'y25': y25_map, 'abs': abs}
                try:
                    is_pass = eval(formula, {"__builtins__": None}, ctx)
                    status = "✅ PASS" if is_pass else "❌ FALSE"
                    detail = "勾稽无误" if is_pass else "差额超过允许阈值"
                except KeyError as e:
                    status = "⚠️ ERROR"
                    detail = f"缺失字段: 找不到 {str(e)}"
                    is_pass = False
                except Exception as e:
                    status = "⚠️ ERROR"
                    detail = "公式语法错误"
                    is_pass = False

                check_results.append({
                    "年度": year_label, 
                    "规则名称": rule_name, 
                    "检查结果": status, 
                    "诊断详情": detail
                })

        # 5. 展示结果报表
        res_display_df = pd.DataFrame(check_results)
        
        def color_status(val):
            if '✅' in str(val): return 'background-color: #C6EFCE; color: #006100;'
            elif '❌' in str(val): return 'background-color: #FFC7CE; color: #9C0006;'
            elif '⚠️' in str(val): return 'background-color: #FFEB9C; color: #9C5700;'
            return ''

        st.dataframe(
            res_display_df.style.applymap(color_status, subset=['检查结果']),
            use_container_width=True, hide_index=True
        )

        # 6. 导出逻辑（接力模式：确保保留 Step 3 的所有原始公式）
        if st.button("📥 导出带勾稽结果的报告"):
            if 'filled_excel' not in st.session_state:
                st.error("❌ 找不到填报后的数据，请先重新运行 Step 3")
                st.stop()
                
            # --- 💡 接力核心：从 session_state 加载 Step 3 已经写好的 Excel ---
            input_buffer = io.BytesIO(st.session_state['filled_excel'])
            workbook = openpyxl.load_workbook(input_buffer)
            
            # --- 💡 动作 1：处理数据明细页的格式 ---
            # 这一步是为了确保即便 Step 3 漏了格式，这里也能补上
            ws_data = workbook.active  # 默认第一个 Sheet 是数据明细
            ws_data.title = "数据明细" # 统一命名
            
            # --- 💡 动作 2：新建勾稽报告页 ---
            # 如果已存在则删除，防止重复点击报错
            if "勾稽报告" in workbook.sheetnames:
                del workbook["勾稽报告"]
            ws_res = workbook.create_sheet("勾稽报告")
            
            # 写入勾稽报告表头
            headers = list(res_display_df.columns)
            for c_idx, header in enumerate(headers, 1):
                ws_res.cell(row=1, column=c_idx).value = header
            
            # 写入勾稽报告内容
            from openpyxl.styles import PatternFill
            fill_green = PatternFill(start_color='C6EFCE', fill_type='solid')
            fill_red = PatternFill(start_color='FFC7CE', fill_type='solid')
            fill_yellow = PatternFill(start_color='FFEB9C', fill_type='solid')
            
            for r_idx, row_data in enumerate(res_display_df.values, 2):
                for c_idx, value in enumerate(row_data, 1):
                    cell = ws_res.cell(row=r_idx, column=c_idx)
                    cell.value = value
                    # 在结果列（第3列）加颜色
                    if c_idx == 3:
                        val_str = str(value)
                        if "PASS" in val_str: cell.fill = fill_green
                        elif "FALSE" in val_str: cell.fill = fill_red
                        elif "ERROR" in val_str: cell.fill = fill_yellow
            
            # 自动调整列宽
            ws_res.column_dimensions['B'].width = 30
            ws_res.column_dimensions['C'].width = 15
            ws_res.column_dimensions['D'].width = 40

            # --- 💡 动作 3：保存并导出 ---
            output = io.BytesIO()
            workbook.save(output)
            
            st.download_button(
                label="点击下载精算复核底稿", 
                data=output.getvalue(), 
                file_name=f"勾稽复核底稿_{st.session_state.get('pdf_name','未命名').replace('.pdf','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("💡 提取结果将在此处显示。请先在上方点击“开始提取”按钮。")
        
# ─────────── Step 5 多公司合并目标表 ───────────

with tab5:
    st.header("📊 多公司数据集成与汇率转换")
    st.info("功能说明：上传多家公司已核查完的“勾稽复核底稿”，系统将自动进行年度拆分、汇率换算并合并为可画图的行业对标长表。")

    # 1. 汇率设置区
    col_rate1, col_rate2 = st.columns(2)
    with col_rate1:
        rate_24 = st.number_input("💵 2024年汇率 (1外币=?人民币)", value=1.0, step=0.0001, format="%.4f")
    with col_rate2:
        rate_25 = st.number_input("💵 2025年汇率 (1外币=?人民币)", value=1.0, step=0.0001, format="%.4f")

    # 2. 多文件上传
    uploaded_files = st.file_uploader("请上传多家公司的勾稽底稿 (可多选)", type="xlsx", accept_multiple_files=True)

    if uploaded_files:
        combined_list = []
        
        for file in uploaded_files:
            # 读取第一个 Sheet (数据明细)
            df_single = pd.read_excel(file, sheet_name=0)
            
            # 识别列名 (兼容你之前的定义)
            c_24 = next((c for c in df_single.columns if '2024' in str(c)), None)
            c_25 = next((c for c in df_single.columns if '2025' in str(c)), None)
            
            if not c_24 or not c_25:
                st.error(f"文件 {file.name} 格式不正确，找不到年份列。")
                continue

            # 基础列
            base_cols = ["公司", "类别", "字段名", "字段类型"]
            existing_base = [c for c in base_cols if c in df_single.columns]

            # 提取 2024 数据
            df_24 = df_single[existing_base + [c_24]].copy()
            df_24["报告年份"] = "2024"
            df_24["汇率"] = rate_24
            df_24 = df_24.rename(columns={c_24: "原币金额"})
            
            # 提取 2025 数据
            df_25 = df_single[existing_base + [c_25]].copy()
            df_25["报告年份"] = "2025"
            df_25["汇率"] = rate_25
            df_25 = df_25.rename(columns={c_25: "原币金额"})

            # 合并当年两载数据
            df_merged = pd.concat([df_24, df_25], ignore_index=True)
            
            # 计算本币金额 (处理可能的清洗逻辑)
            def clean_to_float(val):
                try:
                    if isinstance(val, str):
                        val = val.replace(',','').replace('(','-').replace(')','')
                    return float(val)
                except: return 0.0

            df_merged["(百万)原币"] = df_merged["原币金额"].apply(clean_to_float)
            df_merged["(百万)人民币"] = df_merged["(百万)原币"] * df_merged["汇率"]
            
            combined_list.append(df_merged)

        if combined_list:
            final_all_df = pd.concat(combined_list, ignore_index=True)
            
            # 调整列顺序对齐图二
            final_cols = ["公司", "类别", "字段名", "字段类型", "报告年份", "(百万)原币", "汇率", "(百万)人民币"]
            # 只取存在的列
            actual_cols = [c for c in final_cols if c in final_all_df.columns]
            final_all_df = final_all_df[actual_cols]

            st.success(f"✅ 已成功合并 {len(uploaded_files)} 家公司数据！")
            
            # 展示预览
            st.dataframe(final_all_df, use_container_width=True)

            # 下载集成后的表
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_all_df.to_excel(writer, index=False, sheet_name='行业集成分析表')
                # 可以在这里给数值列加格式...
            
            st.download_button(
                label="📥 下载行业集成对标表 (长表格式)",
                data=output.getvalue(),
                file_name="行业集成对标分析表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
# ─────────── Step 6 可视化 ───────────
# KPMG 官方色卡 (RGB 转换为 Hex 方便 Plotly 使用)
KPMG_COLORS = {
    "KPMG Blue": "#00338D",      # (0, 51, 141)
    "Cobalt Blue": "#1E49E2",    # (30, 73, 226)
    "Dark Blue": "#0C233C",      # (12, 35, 60)
    "Light Blue": "#ACEAFF",     # (172, 234, 255)
    "Pacific Blue": "#00B8F5",   # (0, 184, 245)
    "Purple": "#7213EA",         # (114, 19, 234)
    "Pink": "#FD349C",           # (253, 52, 156)
    "Grey 1": "#333333",
    "Grey 2": "#666666",
    "Light Green": "#63EBB2"
}
# 定义用于图表的颜色序列
COLOR_SEQUENCE = [KPMG_COLORS["KPMG Blue"], KPMG_COLORS["Pacific Blue"], KPMG_COLORS["Cobalt Blue"], 
                  KPMG_COLORS["Pink"], KPMG_COLORS["Purple"], KPMG_COLORS["Light Green"]]


with tab6:
    st.markdown("### 📊 年报可视化分析")
    
    viz_file = st.file_uploader("上传已填写的标准目标表 (Excel)", type=["xlsx"], key="viz_uploader")
    
    if viz_file:
        df_viz = pd.read_excel(viz_file)
        # 🟢 解决 2024.5 问题的关键：强制转为字符串且去掉小数点
        df_viz['报告年份'] = df_viz['报告年份'].astype(str).str.replace('.0', '', regex=False)
        
        # 1. 基础配置
        all_companies = df_viz['公司'].unique().tolist()
        all_fields = df_viz['字段名'].unique().tolist()
        val_col = '(百万)元人民币' if '(百万)元人民币' in df_viz.columns else df_viz.columns[-1]

        # 2. 交互控制区
        c1, c2 = st.columns([1, 2])
        with c1:
            selected_cos = st.multiselect("🏗️ 选择对比公司", all_companies, default=all_companies[:3])
            selected_field = st.selectbox("🎯 选择分析指标", all_fields, index=0)
            
            # 🟢 让领导可以修改代码的开关
            show_editor = st.checkbox("🛠️ 开启代码编辑器 (高级模式)")

        # 准备绘图数据
        filtered_df = df_viz[(df_viz['公司'].isin(selected_cos)) & (df_viz['字段名'] == selected_field)]

        # 3. 默认绘图脚本（作为初始代码）
        default_code = f"""# 你可以修改这里的代码来调整图表呈现
fig = px.bar(
    filtered_df, 
    x="公司", 
    y="{val_col}", 
    color="报告年份", 
    barmode="group",
    text_auto='.2s',
    title=f"{{selected_field}} 年度对比",
    color_discrete_sequence=["#00338D", "#00B8F5"] # KPMG Blue, Pacific Blue
)

fig.update_layout(
    font_family="Microsoft YaHei",
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis_type='category' # 🟢 强制 X 轴为分类，解决 2024.5 问题
)
"""

        # 4. 代码编辑沙盒
        if show_editor:
            st.info("💡 你可以直接在下方修改代码，变量名为 `filtered_df`。修改后点击下方‘运行代码’。")
            user_code = st.text_area("可视化代码编辑器", value=default_code, height=300)
        else:
            user_code = default_code

        # 5. 执行代码并渲染
        if st.button("🚀 运行/刷新图表", use_container_width=True) or not show_editor:
            try:
                # 环境准备：将变量传递给 exec 内部
                local_vars = {
                    'px': px, 'go': go, 'filtered_df': filtered_df, 
                    'selected_field': selected_field, 'KPMG_COLORS': KPMG_COLORS
                }
                # 运行用户编写的代码
                exec(user_code, {}, local_vars)
                
                # 获取执行后的 fig 对象
                if 'fig' in local_vars:
                    st.plotly_chart(local_vars['fig'], use_container_width=True)
                else:
                    st.error("代码中未定义 `fig` 变量！请确保代码最后生成了 fig 对象。")
                    
            except Exception as e:
                st.error(f"代码运行出错: {e}")

        # 6. 辅助饼图（修复之前的 NameError）
        if not show_editor:
            st.divider()
            st.markdown("#### 其他预设维度")
            # 修正后的饼图逻辑
            pie_fig = px.pie(
                filtered_df[filtered_df['报告年份'] == filtered_df['报告年份'].max()], 
                values=val_col, names="公司", 
                title=f"最新年份 {selected_field} 市场份额分布",
                color_discrete_sequence=px.colors.qualitative.Prism
            )
            st.plotly_chart(pie_fig, use_container_width=True)

    else:
        st.info("请先上传目标表 Excel 以开启可视化。")

# ==================== 页脚 ====================
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 13px; letter-spacing: 1px; margin-top: 50px; padding: 20px; border-top: 1px solid #CBD5E1;">
    Actuarial Data Intelligence · Built for KPMG Actuary Team
</div>
""", unsafe_allow_html=True)
