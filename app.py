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

# ==================== 1. 页面基础配置  ====================
st.set_page_config(page_title="寿研数智・年报提取平台", layout="wide", page_icon="💠")

# ==================== 辅助函数：单位====================
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
    
    # 1. 提取前 40 页
    toc_text = ""
    for i in range(min(30, len(doc))):
        page_text = doc.load_page(i).get_text("text").replace("\n", " ")
        toc_text += f"---第{i+1}页---\n{page_text}\n"
        
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
            ["业务及管理费"]
        ],
        "评估假设-折现率假设": [
            ["折现率假设", "折现率曲线"]
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
    
    prompt = f"""你是一个四大寿险精算咨询专家。以下是一份寿险公司年报的前20页文本（主要是目录），以及程序全量扫描的极其精准的附注辅助线索。
请帮我找出以下核心报表在年报中的【物理页码】。

【核心精算业务规则】（非常重要）：
1. 🎯 主表依靠目录推算（重点！）：对于“资产负债表”、“利润表”、“现金流量表”、“股东/所有者权益变动表”，请你【务必仔细阅读前20页的目录文本】，找到它们标注的页码，并加上封面偏移量得出真实物理页码！特别注意“现金流量表”，它一定在目录中，请千万不要漏掉！
2. 🎯 附注表无脑信任线索：对于带有“(附注)”或“保险合同”字眼的表单，目录里找不到，请【无脑完全信任】下面提供的【Python程序附注雷达线索】！直接将线索里的数字转化为数组返回。
3. 合并报表优先：对于三大主表，如果同时存在“合并”和无合并字样，必定提取“合并”的页码；若无“合并”字样则提取原表。
4. 跨页数组返回：因为要连续两年数据，大部分表格跨越连续两页。请返回包含所有相关物理页码的【数组格式】。

需求表单列表：
{json.dumps(target_tables, ensure_ascii=False)}

年报前20页文本（供你寻找主表和推算偏移量）：
{toc_text}
{hint_text}

【强制要求】
1. 只输出合法的 JSON 格式。不要输出任何 Markdown 标记。
2. 键为表单名称，值为纯数字的【数组（List）】。找不到填 [0]。
示例：{{"现金流量表 (合并优先)": [125, 126], "保险合同-未采用保费分配法按计量组成部分分析 (原保险)": [150, 151]}}"""
    
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
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    " 🌐 Step 0 ／ 官网年报监控 ",
    " 📑 Step 1 ／ 智能页码检索 ",
    " ⚡ Step 2 ／ 表格智能转换 ",
    " 📝 Step 3 ／ 目标表标准填报 ",
    " 🔍 Step 4 ／ 数据勾稽检查 ",
    " 📊 Step 5 ／ 可视化分析 "
])

# ─────────── Step 0：官网监控雷达 ───────────
with tab0:
    st.markdown("### 🌐 全网年报监控雷达")
    st.markdown("""
    <div class="info-card green">
        <h4>功能说明</h4>
        <p>上传包含保险公司官网链接的 Excel，系统将像雷达一样自动扫描各个网站，帮你侦测最新年份的年报是否已经发布。<b>（避免每天人工刷网页的烦恼）</b></p>
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
                for table_name, page_data in st.session_state['found_pages'].items():
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
                table_to_view = st.selectbox("选择要预览的报表：", options=list(st.session_state['found_pages'].keys()))
                pages_to_preview = st.session_state['edited_pages'].get(table_to_view, [1])
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
                    st.info("尚未识别到页码，请在左侧手动输入。")
                else:
                    st.warning("该页码超出了文档总页数，请检查修改。")
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
                            
                        SYSTEM_PROMPT = """你是一个专业的寿险精算审计助手。请在数据中寻找指定指标，并区分 2024 和 2025 年。
注意：
1. 别名匹配：若标准名找不到，请查看“别名参考”列。
2. 新业务指标：务必在包含“当期确认”、“初始确认”或“新业务价值”字样的表格中查找。
3. ⚠️绝对指令：输出时，务必【原封不动地照抄】原文里的文本！如果原文带有括号（表示负数）或逗号，比如 "(295,992)"，请直接将其作为【字符串】输出！千万不要删掉括号或负号！
4. 若原文为“—”或“无”请输出 "0"。找不到填 null。
5. 仅输出合法的 JSON，格式为: {"字段名": {"y24": "原样字符串", "y25": "原样字符串"}}"""

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
                                
                            for idx, row in working_df.iterrows():
                                fname = str(row.get(COL_FIELD_NAME, "")).strip()
                                if fname in ai_data:
                                    def to_num(v):
                                        if v is None: return None
                                        s = str(v).strip()
                                        if s in ["", "null", "None", "—", "-", "无"]: return 0.0
                                        s = s.replace(',', '')
                                        if (s.startswith('(') and s.endswith(')')) or (s.startswith('（') and s.endswith('）')): s = '-' + s[1:-1].strip()
                                        try: return float(s)
                                        except: return None
                                            
                                    working_df.at[idx, col_2024] = to_num(ai_data[fname].get("y24"))
                                    working_df.at[idx, col_2025] = to_num(ai_data[fname].get("y25"))
                                    
                            st.write("⚙️ 正在向 Excel 注入计算项目的公式...")
                            temp_buffer = io.BytesIO()
                            working_df.to_excel(temp_buffer, index=False)
                            temp_buffer.seek(0)
                            
                            wb = openpyxl.load_workbook(temp_buffer)
                            ws = wb.active
                            cols = {cell.value: cell.column for cell in ws[1]}
                            c24_idx, c25_idx = cols.get(col_2024), cols.get(col_2025)
                            c24_let, c25_let = get_column_letter(c24_idx), get_column_letter(c25_idx)
                            
                            field_to_row = {str(ws.cell(row=r, column=cols[COL_FIELD_NAME]).value).strip(): r for r in range(2, ws.max_row + 1) if ws.cell(row=r, column=cols[COL_FIELD_NAME]).value}
                            sorted_fields = sorted(field_to_row.keys(), key=len, reverse=True)
                            
                            for r in range(2, ws.max_row + 1):
                                ftype = str(ws.cell(row=r, column=cols[COL_FIELD_TYPE]).value).strip()
                                rule = str(ws.cell(row=r, column=cols[COL_RULE]).value).strip()
                                if ftype == "计算" and rule not in ["nan", "None", ""]:
                                    f24, f25 = rule, rule
                                    for f in sorted_fields:
                                        if f in rule:
                                            row_num = field_to_row[f]
                                            f24, f25 = f24.replace(f, f"{{{{R{row_num}C24}}}}"), f25.replace(f, f"{{{{R{row_num}C25}}}}")
                                    for row_num in field_to_row.values():
                                        f24, f25 = f24.replace(f"{{{{R{row_num}C24}}}}", f"{c24_let}{row_num}"), f25.replace(f"{{{{R{row_num}C25}}}}", f"{c25_let}{row_num}")
                                    try: ws.cell(row=r, column=c24_idx).value, ws.cell(row=r, column=c25_idx).value = f"={f24}", f"={f25}"
                                    except: pass
                                    
                            final_buffer = io.BytesIO()
                            wb.save(final_buffer)
                            final_buffer.seek(0)
                            
                            st.session_state['filled_excel'] = final_buffer.getvalue()
                            st.session_state['working_df'] = working_df 
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
        st.dataframe(st.session_state['working_df'], use_container_width=True)

# ─────────── Step 4 & 5：分析与校验 ───────────
with tab4:
    st.markdown("### 🔍 数据勾稽检查")
    st.info("此模块正在开发中，未来将自动检查报表间的逻辑平衡关系...")

with tab5:
    st.markdown("### 📊 自定义可视化分析")
    st.caption("上传结构化数据文件（Excel / CSV），自由选择维度与指标进行探索式制图。")
    
    viz_file = st.file_uploader("拖拽结构化数据文件至此处", type=["xlsx", "csv"], key="viz_uploader")
    if viz_file:
        st.success(f"已成功加载：{viz_file.name}")
        st.markdown("---")
        st.markdown("#### 分析维度配置")
        col_v1, col_v2 = st.columns(2)
        with col_v1: st.selectbox("X 轴（横轴维度）", ["2023", "2024", "2025", "按公司横向对比"], index=2)
        with col_v2: st.multiselect("Y 轴（纵轴指标）", ["新业务价值", "合同服务边际", "净利润", "内含价值", "保险服务收入", "保险合同负债"], default=["净利润"])
        st.selectbox("图表类型", ["折线图（趋势分析）", "柱状图（横向对比）", "堆叠柱状图（结构分析）", "散点图（相关性分析）"])
        st.button("生成分析图表（开发中）", use_container_width=True)
    else:
        st.markdown("""
        <div class="placeholder-section">
            <div class="big-icon">📊</div>
            <p>请先完成前方流程提取数据后，将生成的 Excel 文件拖入此处进行可视化。</p>
        </div>
        """, unsafe_allow_html=True)

# ==================== 页脚 ====================
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 13px; letter-spacing: 1px; margin-top: 50px; padding: 20px; border-top: 1px solid #CBD5E1;">
    Actuarial Data Intelligence · Built for KPMG Actuary Team
</div>
""", unsafe_allow_html=True)
