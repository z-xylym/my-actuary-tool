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
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from plotly.subplots import make_subplots

# 1. 先定义这个名单（全局变量）
DEFAULT_COMPANIES = [
    {"类别": "上市", "公司": "中国人寿", "类型": "寿险", "链接地址": "https://www.chinalife.com.cn/chinalife/xxpl/gkxxpl/ndxx/"},
    {"类别": "上市", "公司": "平安寿险", "类型": "寿险", "链接地址": "https://life.pingan.com/p/#/list-page?disclosureChannel=iLifeInfoDisclosure&disclosureName=%E5%B9%B4%E5%BA%A6%E4%BF%A1%E6%81%AF&disclosureId=200c389bb2d2463c8e90dbcfb7965bb2"},
    {"类别": "上市", "公司": "太保寿险", "类型": "寿险", "链接地址": "https://www.cpic.com.cn/xrsbx/gkxxpl/ndxx/?subMenu=2&inSub=1"},
    {"类别": "上市", "公司": "泰康人寿", "类型": "寿险", "链接地址": "https://www.taikanglife.com/publicinfonew/annualinfonew/list_402_1.html"},
    {"类别": "上市", "公司": "新华人寿", "类型": "寿险", "链接地址": "https://www.newchinalife.com/node/398"},
    {"类别": "上市", "公司": "太平人寿", "类型": "寿险", "链接地址": "https://life.cntaiping.com/info-ndxxpl/"},
    {"类别": "上市", "公司": "人保寿险", "类型": "寿险", "链接地址": "https://www.picclife.com/picclifewebsite/webfile/ComprehensiveInformation/index.html"},
    {"类别": "上市", "公司": "阳光人寿", "类型": "寿险", "链接地址": "https://www.sinosig.com/v/pilu?type=sx&tabIndex=10001_31"},
    {"类别": "上市", "公司": "友邦人寿", "类型": "寿险", "链接地址": "https://www.aia.com.cn/zh-cn/gongkaixinxipilou/nianduxinxi"},
    {"类别": "银保", "公司": "中邮人寿", "类型": "寿险", "链接地址": "https://www.chinapost-life.com/publish/publish2/"},
    {"类别": "银保", "公司": "工银安盛", "类型": "寿险", "链接地址": "https://www.icbc-axa.com/public/public_year/public_year1/publicfirstindex.jsp"},
    {"类别": "银保", "公司": "交银人寿", "类型": "寿险", "链接地址": "https://www.bocommlife.com/101843/index.html"},
    {"类别": "银保", "公司": "建信人寿", "类型": "寿险", "链接地址": "https://www.ccb-life.com.cn/html/6182/3230/index.html"},
    {"类别": "银保", "公司": "中信保诚", "类型": "寿险", "链接地址": "https://www.citic-prudential.com.cn/annualinformation/list.html"},
    {"类别": "银保", "公司": "农银人寿", "类型": "寿险", "链接地址": "https://www.abchinalife.com/xxpl/ndxx/index.shtml"},
    {"类别": "银保", "公司": "招商信诺", "类型": "寿险", "链接地址": "https://www.cignacmb.com/xinxi/niandubaogao/ndxibg.html"},
    {"类别": "银保", "公司": "中银三星", "类型": "寿险", "链接地址": "https://www.boc-samsunglife.cn/information?code=GW647"},
    {"类别": "养老健康", "公司": "太平养老", "类型": "养老", "链接地址": "https://www.cntaiping.com/about-ndxx/"},
    {"类别": "养老健康", "公司": "恒安标准养老", "类型": "养老", "链接地址": "https://www.haslpension.com/henganyanglao/gkxxpl/ndxx/index.html"},
    {"类别": "养老健康", "公司": "人保健康", "类型": "健康", "链接地址": "https://www.picc.com.cn/xwzx/gkxx/ndxx/"},
    {"类别": "养老健康", "公司": "平安养老", "类型": "养老", "链接地址": "https://yl.pingan.com/branding/disclosure/annual-info"},
    {"类别": "养老健康", "公司": "泰康养老", "类型": "养老", "链接地址": "https://www.tkpension.com/ndxx/"},
    {"类别": "养老健康", "公司": "平安健康", "类型": "健康", "链接地址": "https://health.pingan.com/P/pubInfoDisclosure?menuName=public_information_disclosure&tagId=9223372036854806087&menuId=4035225266123994694"},
    {"类别": "养老健康", "公司": "太保健康", "类型": "健康", "链接地址": "https://www.cpic.com.cn/jkx/gkxxpl/ndxx/?subMenu=2&inSub=1"},
    {"类别": "外资", "公司": "中英人寿", "类型": "寿险", "链接地址": "https://www.aviva-cofco.com.cn/website/xxzx/gkxxpl/ndxxpl/list-1.shtml"},
    {"类别": "外资", "公司": "中意人寿", "类型": "寿险", "链接地址": "https://www.generalichina.com/ndxx/"},
    {"类别": "外资", "公司": "中荷人寿", "类型": "寿险", "链接地址": "https://www.bob-cardif.com/xinxipilu/nianduxinxi/index.html"},
    {"类别": "外资", "公司": "同方全球", "类型": "寿险", "链接地址": "https://www.aegonthtf.com/gkxxpl/ndxx/ndxx/"},
    {"类别": "外资", "公司": "陆家嘴国泰", "类型": "寿险", "链接地址": "https://www.cathaylife.cn/ndxxplbg/index.html"},
    {"类别": "外资", "公司": "复星保德信", "类型": "寿险", "链接地址": "https://www.pflife.com.cn/fbofficialweb/Annual?submenu=Annual"},
    {"类别": "外资", "公司": "恒安标准", "类型": "寿险", "链接地址": "https://www.hengansl.com/hengan/gkxxpl/ndxx/index.html"},
    {"类别": "小型", "公司": "财信吉祥人寿", "类型": "寿险", "链接地址": "https://life.hnchasing.com/annual_info/annual_info_disclosure/"},
    {"类别": "小型", "公司": "东吴人寿", "类型": "寿险", "链接地址": "https://www.soochowlife.net/cs/gkxxpl/ndxx/index.html"},
    {"类别": "小型", "公司": "国富人寿", "类型": "寿险", "链接地址": "https://www.e-guofu.com/yearsInfo/index.html"},
    {"类别": "小型", "公司": "瑞泰人寿", "类型": "寿险", "链接地址": "https://www.oldmutual-chnenergy.com/InfoPublish/anualReport/"},
    {"类别": "小型", "公司": "东方嘉富人寿", "类型": "寿险", "链接地址": "https://www.sinokorealife.com.cn/ndxx.jhtml"}


]


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
    {"规则名称": "19. 费用分类一致性", "公式": "abs((curr['获取费用'] + curr['维持费用'] + curr['非履约费用']) - (curr['职工薪酬'] + curr['物业及设备支出'] + curr['业务投入及监管费用支出'] + curr['行政办公支出'] + curr['其他支出'])) < 10", "类型": "single"},
    {"规则名称": "20. 期初保险合同负债总额", "公式": "abs(curr['期初保险合同负债总额'] - (curr['CSM期初余额'] + curr['BEL期初余额'] + curr['RA期初余额'])) < 1", "类型": "single"},
    {"规则名称": "21. 期末保险合同负债总额", "公式": "abs(curr['期末保险合同负债总额'] - (curr['CSM期末余额'] + curr['BEL期末余额'] + curr['RA期末余额'])) < 1", "类型": "single"},
    {"规则名称": "22. 非PAA期初余额合计", "公式": "abs(curr['非PAA期初余额合计'] - (curr['LRC非亏损部分期初余额（非PAA）'] + curr['LRC亏损部分期初余额（非PAA）'] + curr['LIC期初余额（非PAA）'])) < 1", "类型": "single"},
    {"规则名称": "23. 非PAA期末余额合计", "公式": "abs(curr['非PAA期末余额合计'] - (curr['LRC非亏损部分期末余额（非PAA）'] + curr['LRC亏损部分期末余额（非PAA）'] + curr['LIC期末余额（非PAA）'])) < 1", "类型": "single"},
    {"规则名称": "24. PAA期初余额合计", "公式": "abs(curr['PAA期初余额合计'] - (curr['LRC非亏损部分期初余额（PAA）'] + curr['LRC亏损部分期初余额（PAA）'] + curr['LIC期初BEL余额（PAA）'] + curr['LIC期初RA余额（PAA）'])) < 1", "类型": "single"},
    {"规则名称": "25. PAA期末余额合计", "公式": "abs(curr['PAA期末余额合计'] - (curr['LRC非亏损部分期末余额（PAA）'] + curr['LRC亏损部分期末余额（PAA）'] + curr['LIC期末BEL余额（PAA）'] + curr['LIC期末RA余额（PAA）'])) < 1", "类型": "single"}
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
import base64

# ==================== 新增后台引擎：调用 Vision VLM 提取图片表格 ====================
def extract_single_page_vision(pdf_bytes, page_num, expected_name, api_key, base_url, model_name):
    """将 PDF 转为高清图片，使用视觉大模型提取表格"""
    try:
        # 1. 内存中将 PDF 页面转为 Base64 高清图
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        if page_num < 1 or page_num > len(doc):
            return None, ""
        page = doc.load_page(page_num - 1)
        
        # 放大 3倍截图，保证文字清晰
        zoom_matrix = fitz.Matrix(3.0, 3.0) 
        pix = page.get_pixmap(matrix=zoom_matrix, alpha=False)
        img_data = pix.tobytes("png")
        base64_image = base64.b64encode(img_data).decode('utf-8')
        doc.close()

        # 2. 设置读图的 Prompt
        prompt = f"""你是一个四大会计师事务所的资深数字化审计专家。
我为你提供了一张【保险公司年报的单页高清截图】，目标是精准提取：【{expected_name}】。
表格中的【数字是由图片构成的】，请发挥强大的视觉识别能力，提取所有文字和表格。
【强制要求】：
1. 所有的科目名、数字金额之间，必须被 "|" 隔开。严禁使用连续空格。
2. 精准对齐多级表头！如果有留白单元格，必须用 "|" 占位补齐！
3. 绝不能漏掉任何一行数据，保留金额里的逗号和括号。
4. 纯文本输出，不要使用 Markdown 代码块。不需要输出 SHEET_NAME 标签，直接排版。"""

        # 3. 呼叫视觉大模型
        client = OpenAI(api_key=api_key, base_url=base_url)
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}}
                    ]
                }
            ],
            temperature=0.0
        )
        
        # 4. 解析结果
        result_text = response.choices[0].message.content.strip()
        result_text = re.sub(r'^```(csv|text)?\n?', '', result_text, flags=re.MULTILINE).replace("```", "").strip()
        
        lines = result_text.split('\n')
        parsed_data = []
        max_cols = 0
        for row in lines:
            clean_row = row.strip()
            if not clean_row: continue
            if re.match(r'^[\s\|-]+$', clean_row) and '-' in clean_row: continue
            
            if clean_row.startswith('|'): clean_row = clean_row[1:]
            if clean_row.endswith('|'): clean_row = clean_row[:-1]
            
            cols = [col.strip() for col in clean_row.split('|')] if '|' in clean_row else [clean_row]
            parsed_data.append(cols)
            max_cols = max(max_cols, len(cols))
            
        for row in parsed_data:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
                
        return pd.DataFrame(parsed_data), "【提示：该页已使用 OCR 视觉引擎处理，原文本为图片结构】"
        
    except Exception as e:
        return None, f"图像处理失败: {str(e)}"
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
    # 使用分列布局，把“❓”变成一个小图标放在右上角
    c_title, c_icon = st.columns([5, 1])
    with c_icon:
        with st.popover("❓"):
            # 使用 HTML 美化框内排版
            st.markdown("""
            <div style="font-size: 13px; color: #444; line-height: 1.6;">
                <div style="color: #00338D; font-weight: bold; margin-bottom: 4px;">🔑 API Key 去哪找？</div>
                • <b>DeepSeek:</b> 官方开放平台获取，用于大部分PDF的文本提取。<br>
                • <b>Vision API:</b> 视觉模型（如千问/通义等），用于处理图片和扫描件的PDF。<br>
                <br>
                <div style="color: #00338D; font-weight: bold; margin-bottom: 4px;">🔗 接口与模型参数</div>
                • <b>Base URL:</b> 模型接口地址。官方保持默认，中转站填代理地址。<br>
                • <b>模型名称:</b> 供应商代号（如 <code style="font-size:12px; color:#c7254e;">deepseek-chat</code>）。
            </div>
            """, unsafe_allow_html=True)
            
            # 美化的免责声明提示框
            st.markdown("""
            <div style="background-color: #FFF9E6; padding: 10px; border-radius: 5px; font-size: 12px; color: #856404; border-left: 4px solid #FFD966; margin-top: 15px; margin-bottom: 15px;">
                <b>⚠️ 温馨提示：</b><br>
                由于大模型偶尔抽风，使用过程中出现bug可以多试几次，个别公司的目标表填充结果可能并不充分，最终数据核对方面<b>还是需要人工复核</b>。<br>
                后续会不断改进，请期待~ ✨
            </div>
            """, unsafe_allow_html=True)

            # ======== 👇 新增的视频播放区域 👇 ========
            st.markdown("<div style='color: #00338D; font-weight: bold; font-size: 13px; margin-bottom: 5px;'>📺 网页使用教程</div>", unsafe_allow_html=True)
            
            try:
                # ⚠️ 把下面引号里的内容，替换成你刚刚【右键复制的那个 Release 链接】！
                video_url = "https://github.com/z-xylym/my-actuary-tool/releases/download/v2.0/-DIGILIFE.mp4"
                
                # st.video 可以直接读取网络上的 mp4 链接并播放
                st.video(video_url)
            except Exception as e:
                st.caption("教程视频加载中...")

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
    st.markdown("#### 📸 (仅用于处理图片PDF)")
    st.caption("提示：推荐使用阿里的 qwen-vl-max 或 gpt-4o")
    vision_api_key = st.text_input("Vision API Key", type="password", placeholder="sk-...", key="vision_key")
    vision_base_url = st.text_input("Vision Base URL", value="https://dashscope.aliyuncs.com/compatible-mode/v1", key="vision_url")
    vision_model = st.text_input("模型名称 (Model)", value="qwen-vl-max", key="vision_model")
    st.markdown("""
    <div style="font-size:12px; color:#64748B; line-height:2.2; letter-spacing:1px;">
    系统版本：v2.0 (Alpha)<br>
    开发者：林友沐
    </div>
    """, unsafe_allow_html=True)

# ==================== 主界面 ====================
st.title("寿研数智・年报处理平台")

# 创建包含 Step 0 的全新导航栏
tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    " 🌐 Step 0 ／ 官网年报监控 ",
    " 📑 Step 1 ／ 智能页码检索 ",
    " ⚡ Step 2 ／ 表格智能转换 ",
    " 📝 Step 3 ／ 目标表标准填报 ",
    " 🔍 Step 4 ／ 数据勾稽检查 ",
    " ⛓️‍💥 Step 5 ／ 多公司数据合并",
    " 📊 Step 6 ／ 可视化分析 ",
    " 🖼️ Step 7 ／ 图片链接PPT "
])

# ─────────── Step 0：年报更新监控 ───────────
with tab0:
    st.markdown("### 🌐 寿险公司官网年报监控")
    
    # 1. 年份选择器
    col_t0_1, _ = st.columns([1, 2])
    with col_t0_1:
        target_year = st.number_input("📅 请选择监控年份", min_value=2010, max_value=2050, value=2025, step=1)
    
    st.markdown(f"""
    <div class="info-card green">
        <h4>功能说明</h4>
        <p>系统将自动扫描下列 36 家险企官网，检测网页中是否出现 <b>{target_year}</b> 字样。您可以<b>双击下方表格修改网址</b>。检测到后请点击链接手动下载 PDF 并前往 Step 1。</p>
    </div>
    """, unsafe_allow_html=True)

    # 2. 将预设数据加载到 Session State
    if 'company_df' not in st.session_state:
        # 使用之前提供的 36 家公司 DEFAULT_COMPANIES 列表
        st.session_state.company_df = pd.DataFrame(DEFAULT_COMPANIES)
    
    # 3. 交互式数据表 (可直接修改网址)
    st.markdown("#### 🏢 监控目标名单 (双击下方链接可修改)")
    edited_df = st.data_editor(
        st.session_state.company_df,
        column_config={
            "链接地址": st.column_config.TextColumn("🌐 网页链接地址 (支持双击编辑)", max_chars=1000, width="large"),
            "类别": st.column_config.TextColumn(disabled=True),
            "公司": st.column_config.TextColumn(disabled=True),
            "类型": st.column_config.TextColumn(disabled=True),
        },
        use_container_width=True,
        hide_index=True,
        key="company_data_editor_step0"
    )
    st.session_state.company_df = edited_df

    st.markdown("---")
    
    # 4. 扫描按钮逻辑
    col_btn1, col_btn2 = st.columns([1, 1])
    
    with col_btn1:
        target_company = st.selectbox("🎯 单家快速检索：", options=edited_df['公司'].tolist())
        single_scan = st.button(f"🔍 检索【{target_company}】", use_container_width=True)
    
    with col_btn2:
        st.write("") # 对齐
        batch_scan = st.button("🚀 启动全网 36 家批量扫描", type="primary", use_container_width=True)

    # 5. 执行扫描逻辑 (通用函数)
    if single_scan or batch_scan:
        # 确定要扫描的任务列表
        if single_scan:
            tasks = edited_df[edited_df['公司'] == target_company].to_dict('records')
        else:
            tasks = edited_df.to_dict('records')
            
        results = []
        my_bar = st.progress(0, text="准备启动扫描...")
        total_tasks = len(tasks)
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        for index, row in enumerate(tasks):
            company_name = row['公司']
            url = str(row['链接地址'])
            my_bar.progress((index + 1) / total_tasks, text=f"正在扫描 ({index+1}/{total_tasks}): {company_name}")
            
            status = "🔴 未更新 / 无法访问"
            if url.lower().endswith('.pdf'):
                status = "⚠️ 直接PDF链接 (需手动核实)"
            elif "http" in url:
                try:
                    # 原有的爬虫逻辑开始
                    response = requests.get(url, headers=headers, timeout=8)
                    response.encoding = response.apparent_encoding 
                    soup = BeautifulSoup(response.text, 'html.parser')
                    page_text = soup.get_text()
                    
                    # 检查网页文本中是否包含用户输入的年份
                    if str(target_year) in page_text:
                        status = "🟢 极可能已更新!"
                    # 原有的爬虫逻辑结束
                except Exception as e:
                    status = "🟡 网站拦截/超时，需手动查看"
            else:
                status = "无效链接"
            
            results.append({
                "公司名称": company_name,
                "监控状态": status,
                "检测结果描述": f"网页中已检索到 {target_year} 字样" if "🟢" in status else "未发现关键字",
                "直达链接": url
            })
            time.sleep(0.3) # 稍微提速

        my_bar.empty()
        st.success(f"🎉 {target_year}年度检测扫描完成！")
        
        # 6. 展示扫描结果
        df_result = pd.DataFrame(results)
        st.data_editor(
            df_result,
            column_config={
                "直达链接": st.column_config.LinkColumn("点击前往网页"),
                "监控状态": st.column_config.TextColumn("状态", width="medium")
            },
            hide_index=True,
            use_container_width=True,
            key="scan_result_display"
        )
        st.info("💡 提示：请点击标记为 🟢 的公司链接下载 PDF，下载后请前往 [Step 1] 进行页码定位。")
# ─────────── Step 1：智能页码检索 ───────────
with tab1:
    st.markdown("### 📑 智能页码检索")
    uploaded_file = st.file_uploader(
        "拖拽或选择一份已经下载好的年报 PDF 文件",
        type="pdf",
        help="推荐上传一份寿险公司的完整年报"
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
                    with st.spinner("🌸 别着急，我正在努力识别您需要的表格，看看有没有跨页，请稍等一下呦~"):
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
        <p>系统将调用 DeepSeek 对多页 PDF 进行网格化对齐重构，并自动拼接同表跨页数据，生成标准化 Excel。<b>如果出现Is this really a PDF?的提取失败标识<b>，1.请用浏览器（Chrome 或 Edge）打开这个报错的 PDF。2.点击打印 (Ctrl + P)在打印机选项里选择 “另存为 PDF”。3.保存为一个新文件，然后再上传到你的系统。</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("")

    # === 🌟 新增：引擎选择开关 ===
    st.markdown("#### PDF类型选择")
    use_vision = st.toggle("📸 开启图片扫描模式 (适用于扫描版、纯图片PDF)", value=False)
    st.markdown("---")
    
    if 'edited_pages' not in st.session_state or 'pdf_bytes' not in st.session_state:
        st.warning("⚠️ 请先在 Step 1 中上传文件并完成【页码确认】。")
        st.button("开始提取结构化数据", disabled=True, use_container_width=True)
    else:
        # === 校验对应的 API Key 是否已填写 ===
        can_run = True
        if not use_vision and not api_key:
            st.error("⚠️ 当前为文本模式：请先在左侧输入 DeepSeek API Key。")
            can_run = False
        if use_vision and not st.session_state.get('vision_key'):
            st.error("⚠️ 当前为图片扫描模式：请先在左侧输入 Vision API Key。（如：千问、gpt)")
            can_run = False

        if not can_run:
            st.button("开始提取结构化数据", disabled=True, use_container_width=True)
        else:
            valid_tasks = {k: v for k, v in st.session_state['edited_pages'].items() if v != [0]}
            
            if not valid_tasks:
                st.warning("⚠️ 没有找到任何有效的页码，无需提取。")
            else:
                if st.button("🚀 开始极速提取", use_container_width=True):
                    # 1. 补全了冒号 2. 下方所有代码整体向右缩进一步
                    with st.spinner("☕ 趁现在喝口水吧，正在后台拼命打工中..."):
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
                            # 💡 核心优化：视觉模型并发太高容易被限流，设定为2；普通文本设为5
                            workers = 2 if use_vision else 5 
                            
                            with ThreadPoolExecutor(max_workers=workers) as executor:
                                future_to_task = {}
                                
                                for task in all_tasks:
                                    # 👇 核心分流逻辑：根据开关决定派发哪个函数 👇
                                    if use_vision:
                                        future = executor.submit(
                                            extract_single_page_vision, 
                                            st.session_state['pdf_bytes'], 
                                            task["page"], 
                                            task["table"], 
                                            st.session_state.vision_key,
                                            st.session_state.vision_url,
                                            st.session_state.vision_model
                                        )
                                    else:
                                        future = executor.submit(
                                            extract_single_page, 
                                            st.session_state['pdf_bytes'], 
                                            task["page"], 
                                            task["table"], 
                                            api_key
                                        )
                                    future_to_task[future] = task
                                
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
                        
                        # 拼接和导出的代码也需要在 spinner 内部执行
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
    
    # 1. 功能说明卡片
    st.markdown("""
    <div class="info-card pink">
        <h4>功能说明</h4>
        <p>AI 将自动寻找科目数据填充数值，生成 Excel 计算公式。支持内置模板或上传自定义模板。</p>
    </div>
    """, unsafe_allow_html=True)

    # 2. 建立公司类型映射字典
    COMPANY_TYPE_MAP = {item["公司"]: item["类别"] for item in DEFAULT_COMPANIES}
    
    # 3. 模板选择逻辑 (内置 vs 上传)
    col_t1, col_t2 = st.columns([1, 1])
    with col_t1:
        use_default = st.toggle("使用系统默认模板", value=True, help="开启后将直接使用内置的 default_template.xlsx")
    
    template_file = None
    default_path = "default_template.xlsx" 

    if use_default:
        import os
        if os.path.exists(default_path):
            template_file = default_path
            st.info(f"💡 当前正在使用：系统内置模板 ({default_path})")
        else:
            st.error("❌ 未找到默认模板文件，请切换为上传模式！")
            use_default = False 

    if not use_default:
        # 这里只保留一个 file_uploader
        template_file = st.file_uploader("上传自定义目标表模板 (.xlsx)", type="xlsx", key="unique_template_uploader")

    # 4. 核心填报逻辑
    if template_file:
        # 定义标准列名
        COL_COMPANY, COL_CATEGORY, COL_FIELD_NAME, COL_FIELD_TYPE, COL_NOTE, COL_RULE = "公司", "类别", "字段名", "字段类型", "注释", "计算规则"
        COL_CO_TYPE = "公司类型"

        if 'extracted_data' not in st.session_state:
            st.warning("⚠️ 尚未找到提取的数据，请先完成 Step 1 和 Step 2。")
        else:
            if st.button("✨ 启动智能填报与公式生成", use_container_width=True):
                # 使用你想要的友好等待语
                status_box = st.status("🌸 抽空放松一下，我在努力，稍等一下啦~", expanded=True)
                
                with status_box:
                    st.write("📄 正在解析模板表结构...")
                    template_df = pd.read_excel(template_file)
                    
                    # 识别年份列
                    col_2024 = next((c for c in template_df.columns if '2024' in str(c)), None)
                    col_2025 = next((c for c in template_df.columns if '2025' in str(c)), None)
                    
                    if not col_2024 or not col_2025:
                        st.error("❌ 错误：模板中未找到包含 2024 或 2025 的列头！")
                    else:
                        # --- 🚀 A. 公司名与公司类型自动匹配 ---
                        raw_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
                        company_short = re.sub(r'202\d年.*', '', raw_name).strip()
                        
                        # 模糊匹配公司类型
                        matched_type = "其他"
                        for c_name, c_type in COMPANY_TYPE_MAP.items():
                            if c_name[:2] in company_short or company_short[:2] in c_name:
                                matched_type = c_type
                                break
                        
                        working_df = template_df.copy()
                        if COL_COMPANY in working_df.columns:
                            working_df[COL_COMPANY] = company_short
                        if COL_CO_TYPE in working_df.columns:
                            working_df[COL_CO_TYPE] = matched_type
                        
                        # --- B. 准备 AI 任务 ---
                        st.write("🎯 正在准备目标指标清单...")
                        input_items = working_df[working_df[COL_FIELD_TYPE].astype(str).str.strip() == "输入"]
                        ai_target_list = [{"类别": str(r.get(COL_CATEGORY, "")), "标准字段名": str(r.get(COL_FIELD_NAME, "")), "别名参考": str(r.get(COL_NOTE, "")) if pd.notna(r.get(COL_NOTE)) else ""} for _, r in input_items.drop_duplicates(subset=[COL_FIELD_NAME]).iterrows()]
                            
                        st.write("🧠 正在组装财务上下文供 AI 分析...")
                        extracted_data = st.session_state['extracted_data']
                        context_text = "".join([f"\n[表名: {name}]\n{df.to_csv(index=False, sep='|')}\n" for name, df in extracted_data.items()])
                        

                            
                        SYSTEM_PROMPT = """你是一个资深的寿险精算审计专家。任务：将 PDF 提取的财务明细精准填入目标底稿，并严格区分 2024 和 2025 年度。

【通用执行准则：绝对原样提取】
1. ⚠️ 绝对指令：除了下述特殊的加总需求外，所有指标必须【原封不动地照抄】原文里的文本！
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

【特殊指令：保险合同负债新准则科目映射（最高寻址优先级）】
为了防止名称混淆，当提取以下指标时，【必须严格锁定对应的表名】，并应用精算缩写翻译规则：
1. 锁定表名：[表名: 保险合同-未采用保费分配法按计量组成部分分析 (原保险)]
   - 映射字典：BEL = 未来现金流量现值；RA = 非金融风险调整；总额 = 合计。
   - 目标字段限定：当遇到【BEL期初余额、BEL期末余额、RA期初余额、RA期末余额、期初保险合同负债总额、期末保险合同负债总额】时，仅在此表中寻址取数。

2. 锁定表名：[表名: 保险合同-非PAA按未到期责任负债和已发生赔款负债分析 (原保险)]
   - 映射字典：LRC = 未到期责任负债；LIC = 已发生赔款负债。
   - 目标字段限定：当遇到任何带有【（非PAA）】后缀的字段（如LRC非亏损部分期初余额（非PAA）、LIC期末余额（非PAA）、非PAA期末余额合计等）时，仅在此表中寻址。

3. 锁定表名：[表名: 保险合同-PAA按未到期责任负债和已发生赔款负债分析 (原保险)]
   - 映射字典：LRC = 未到期责任负债；LIC = 已发生赔款负债。
   - 目标字段限定：当遇到任何带有【（PAA）】后缀的字段（如LRC非亏损部分期初余额（PAA）、LIC期初BEL余额（PAA）、PAA期初余额合计等）时，仅在此表中寻址。
4.【时间维度与排版防反转绝对指令（适用任意年份）】
中国企业财报的标准排版规则是：【左边的数据列为当期（最新年份），右边的数据列为上期（历史年份）】。
- 当填报要求中包含“较晚/最新年份”（如当期/本年/年末）的键时：必须从原表表头的“本期”、“本年”、“年末”或【最左侧的数据列】取数。
- 当填报要求中包含“较早/历史年份”（如上期/上年/年初）的键时：必须从原表表头的“上期”、“上年”、“年初”或【紧挨着它的右侧数据列】取数。
⚠️ 警告：大语言模型极易犯“线性思维”错误（误以为小年份排在左边，大年份排在右边）。要求你必须打破定势，仔细核对表头！永远记住：左边的列才是最新当期！
例如：若输出键包含 y24 和 y25，则 y25 (最新年) 取报表左侧当期列，y24 (历史年) 取报表右侧上期列。
仅输出合法的 JSON 格式，严禁带有任何 Markdown 标记或文字说明。
格式示例：{"字段名": {"y24": "(1,234.0) + 500.0", "y25": "9,876.54"}}"""
                        st.write("🚀 正在呼叫大模型进行语义映射与精准取数...")
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
                            # --- 🚀 新增：公司类型自动匹配逻辑 ---
                            raw_name = st.session_state.get('pdf_name', '未命名').replace(".pdf", "")
                            # 提取公司简称，如 "新华保险"
                            company_short = re.sub(r'202\d年.*', '', raw_name).strip()
                            
                            # 匹配公司类型
                            matched_type = "其他" # 默认值
                            # 模糊匹配逻辑：遍历大名单，看识别出的名字是否包含在大名单里，或者大名单里的名字是否在识别名里
                            for c_name, c_type in COMPANY_TYPE_MAP.items():
                                # 比如 "新华保险" 和 "新华人寿"，只要前两个字对上通常就对了
                                if c_name[:2] in company_short or company_short[:2] in c_name:
                                    matched_type = c_type
                                    break
                            
                            # 回填“公司”列
                            if COL_COMPANY in working_df.columns:
                                working_df[COL_COMPANY] = company_short
                            
                            # 回填“公司类型”列（根据你的截图，第一列是“公司类型”）
                            COL_CO_TYPE = "公司类型"
                            if COL_CO_TYPE in working_df.columns:
                                working_df[COL_CO_TYPE] = matched_type
                            # ----------------------------------------
                                
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
            label="⬇️ 下载已填报的底稿 (含公式)",
            data=st.session_state['filled_excel'],
            file_name=f"{company_name}_自动填报表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("💡 *预览表仅展示数值，下载后的 Excel 已包含自动关联公式。*")
        st.dataframe(st.session_state['final_df'], use_container_width=True)


# ─────────── Step 4 分析与校验 ───────────
with tab4:
    # 🎨 注入自定义 CSS，强制缩小按钮和 Popover 的尺寸
    st.markdown("""
        <style>
        /* 缩小 Popover 按钮 */
        div[data-testid="stPopover"] > button {
            padding: 2px 8px !important;
            min-height: 26px !important;
            height: 26px !important;
            font-size: 12px !important;
            background-color: #f8f9fa !important;
            border: 1px solid #ddd !important;
        }
        /* 缩小并修正刷新按钮 */
        div[data-testid="stColumn"] button[kind="secondary"] {
            padding: 2px 10px !important; /* 左右间距稍微大一点点 */
            height: 26px !important;
            font-size: 12px !important;
            white-space: nowrap !important; /* 强制文字不换行 */
            min-width: 70px !important;    /* 保证按钮有最小宽度容纳 🔄刷新 */
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
        }
        /* 调整 subheader 间距 */
        .stSubheader { margin-bottom: -10px !important; }
        </style>
    """, unsafe_allow_html=True)

    # 1. 顶部工具栏布局
    # 调整比例，新增一个列用来放下载按钮
    col_title, col_btn_sample, col_btn_refresh, col_manual_area = st.columns([3.5, 1.8, 1.1, 1.8])
    
    with col_title:
        st.subheader("🏁 精算数据勾稽关系检查")
        
    with col_btn_sample:
        # 使用弹窗容纳下载按钮，保持界面 CSS 高度一致
        with st.popover("校验表获取", use_container_width=True):
            st.caption("将人工校验两列复制到需要校对的表格对应位置")
            try:
                # 🎈 去掉了 assets/，直接读取当前目录的文件
                with open("新华样例.xlsx", "rb") as f:
                    st.download_button("⬇️ 下载样例文件", f, file_name="新华人工校验样例.xlsx", use_container_width=True)
            except FileNotFoundError:
                st.error("找不到文件，请确认新华样例.xlsx和app.py在一起")
    
    with col_btn_refresh:
        if st.button("🔄 刷新", use_container_width=True):
            st.rerun() 

    # --- 📖 核对重点与语音 (合二为一的区域) ---
    with col_manual_area:
        # 在这一个列里再分两个超微型列，实现“字标同行”
        c1, c2 = st.columns([1, 0.4])
        
        # --- 📖 操作手册弹出框 ---
        manual_content = """
        精算数据人工核对指南
        
        1. 模板准备： 请将新华样例中的“人工校验”两列手动复制到当前需要核对的表格中。
        2. 年度对齐： 重点检查 2024 年和 2025 年的数据是否有填反现象。
        3. PAA 计算： 采用保费分配法保险业绩 = 保费分配法下的服务收入 - 服务费用。
        4. 符号修正： 部分公司利润表营业支出披露为正，需人工对每一项添加负号（原本负号变正）。
        5. 费用拆解： 业务及管理费中，职工薪酬等项有时需人工尝试加总核对，维持费用、获取费用应为正数，保证进出项总和相等。
        6. 外部搜索： 折现率假设、非金融风险置信水平需手动在 PDF 中全文搜索并填入。
        7. CSM 摊销： 1-20年合同服务边际摊销之和应等于 1。请查看“未来现金流现值、风险调整、CSM”页下方注释，或手动计算CSM摊销比例。
        """

        with c1:
            with st.popover("🕶️核对指南", use_container_width=True):
                # 💡 这里的 HTML 字符串必须紧贴左侧，或者确保没有多余的缩进空格
                html_code = f"""
<div style="font-family: 'Microsoft YaHei', sans-serif; padding: 5px;">
    <h4 style="color: #1f4e79; margin: 0 0 10px 0; font-size: 1rem; border-bottom: 2px solid #e1e4e8; padding-bottom: 5px;">
        精算数据人工核对指南
    </h4>
    <div style="line-height: 1.6; font-size: 0.85rem; color: #333;">
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">1. 模板准备：</strong> 
            将新华样例的“人工校验”两列手动复制到核对表中。
        </p>
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">2. 年度对齐：</strong> 
            检查 <span style="background-color: #fff3cd; padding: 1px 3px;">2024</span> 和 <span style="background-color: #fff3cd; padding: 1px 3px;">2025</span> 数据是否有填反。
        </p>
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">3. PAA 计算：</strong> 
            采用保费分配法保险业绩 = 保费分配法下的服务收入 - 服务费用。
        </p>
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">4. 符号修正：</strong> 
            利润表营业支出里的项目如为正，需人工加负号（原本负号变正）。
        </p>
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">5. 费用拆解：</strong> 
            业管费分类需人工加总，维持费用和获取费用为正，确保进出项总和相等。
        </p>
        <p style="margin-bottom: 8px;">
            <strong style="color: #d32f2f;">6. 外部搜索：</strong> 
            折现率和置信水平需在 PDF 中全文搜索填入。
        </p>
        <p style="margin-bottom: 0;">
            <strong style="color: #d32f2f;">7. CSM 摊销：</strong> 
            1-20年摊销之和应等于 1。查看合同服务边际页下注释或手动计算摊销比例。
        </p>
    </div>
</div>
"""
                # ⚠️ 关键点：unsafe_allow_html=True 必须有，且 html_code 前后不要有导致 Markdown 识别为代码块的空格
                st.markdown(html_code, unsafe_allow_html=True)
        with c2:
            # 清洗语音文本
            clean_voice_text = manual_content.replace('\n', '。').replace('**', '')
            tts_html = f"""
            <div style="display: flex; align-items: center; height: 26px;">
                <button id="tts-btn" onclick="toggleSpeak()" style="
                    border: 1px solid #ddd;
                    background-color: #f8f9fa;
                    border-radius: 4px;
                    cursor: pointer;
                    width: 26px;
                    height: 26px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 12px;
                ">📢</button>
            </div>
            <script>
            var msg = new SpeechSynthesisUtterance();
            msg.text = "{clean_voice_text}";
            msg.lang = 'zh-CN';
            msg.rate = 1.2;
            function toggleSpeak() {{
                const btn = document.getElementById('tts-btn');
                if (window.speechSynthesis.speaking) {{
                    window.speechSynthesis.cancel();
                    btn.style.backgroundColor = "#f8f9fa";
                    btn.innerText = "📢";
                }} else {{
                    window.speechSynthesis.speak(msg);
                    btn.style.backgroundColor = "#ffebeb"; 
                    btn.innerText = "⏹";
                    msg.onend = function() {{
                        btn.style.backgroundColor = "#f8f9fa";
                        btn.innerText = "📢";
                    }};
                }}
            }}
            </script>
            """
            st.components.v1.html(tts_html, height=30)

            
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
            res_display_df.style.map(color_status, subset=['检查结果']),
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
    col_title, col_btn = st.columns([4, 1])
    with col_title:
        st.subheader("⛓️‍💥 多公司数据集成与汇率/单位转换")
    st.info("功能说明：支持上传单文件多Sheet或多文件。系统将自动提取所有公司，请在下方为不同公司配置对应的汇率和原始表格的单位。")

    # 1. 多文件上传
    uploaded_files = st.file_uploader("请上传已完成勾稽检查的底稿 (支持多文件或单文件多Sheet)", type="xlsx", accept_multiple_files=True)

    if uploaded_files:
        # 第一步：初步扫描所有文件中的所有 Sheet，确定有哪些公司
        all_temp_data = [] # 存储 (文件名, Sheet名, df)
        found_companies = set()

        for file in uploaded_files:
            # 获取该 Excel 的所有 Sheet 名称
            xl = pd.ExcelFile(file)
            for sheet_name in xl.sheet_names:
                df_raw = pd.read_excel(file, sheet_name=sheet_name)
                # 尝试从“公司”列获取公司名，如果没这一列，就用 Sheet 名作为公司名
                current_company = ""
                if "公司" in df_raw.columns and not df_raw["公司"].empty:
                    current_company = str(df_raw["公司"].iloc[0])
                else:
                    current_company = sheet_name
                
                # 过滤掉一些明显的系统 Sheet (如 基本信息_单位)
                if "基本信息" in current_company or "Sheet" in current_company:
                    continue
                
                found_companies.add(current_company)
                all_temp_data.append({
                    "comp": current_company,
                    "df": df_raw,
                    "source": f"{file.name} - {sheet_name}"
                })

        # 2. 动态汇率与单位配置区
        st.markdown("#### 💵 汇率与数值单位配置盘")
        st.caption("目标表统一要求以【百万元人民币】展示。请根据原始底稿设置：1.兑人民币汇率 2.原始底稿的金额单位。")
        
        rate_config = {}
        unit_config = {}
        
        # 🌟 定义原始单位转换为“百万元”的乘法系数
        unit_multipliers = {
            "原表为: 百万元 (无需转换)": 1.0,
            "原表为: 元 (÷ 1,000,000)": 0.000001,
            "原表为: 千元 (÷ 1,000)": 0.001,
            "原表为: 亿元 (× 100)": 100.0,
            "原表为: 十亿元 (× 1,000)": 1000.0
        }
        
        # 创建一个多列布局来放输入框
        rate_cols = st.columns(3) 
        for i, comp in enumerate(sorted(list(found_companies))):
            with rate_cols[i % 3]:
                # 使用 container 包装一下让视觉更紧凑
                with st.container(border=True):
                    st.markdown(f"**🏢 {comp}**")
                    rate_config[comp] = st.number_input(f"汇率 (相对于RMB)", value=1.0, step=0.0001, format="%.4f", key=f"rate_{comp}")
                    unit_choice = st.selectbox(f"原表数值单位", list(unit_multipliers.keys()), key=f"unit_{comp}")
                    unit_config[comp] = unit_multipliers[unit_choice]

        # 3. 执行合并逻辑
        if st.button("🚀 开始集成并换算数据", type="primary", use_container_width=True):
            # 🌟 精确匹配白名单：只有【完全等于】以下名称的字段，才会被豁免换算
            exact_exempt_fields = [
                "折现率假设",
                "非金融风险的置信水平",
                "1年及1年以内合同服务边际",
                "1-5年合同服务边际",
                "5-10年合同服务边际",
                "10-20年合同服务边际",
                "20年合同服务边际"
            ]
            
            combined_list = []
            
            for item in all_temp_data:
                df_single = item["df"]
                comp_name = item["comp"]
                rate = rate_config[comp_name]
                unit_mult = unit_config[comp_name]
                
                c_24 = next((c for c in df_single.columns if '24' in str(c)), None)
                c_25 = next((c for c in df_single.columns if '25' in str(c)), None)
                
                if not c_24 or not c_25:
                    continue

                base_cols = ["公司", "类别", "字段名", "字段类型"]
                existing_base = [c for c in base_cols if c in df_single.columns]

                def clean_to_float(val):
                    try:
                        if isinstance(val, (int, float)): return float(val)
                        if isinstance(val, str):
                            val = val.replace(',','').replace('(','-').replace(')','').strip()
                            if val == '-' or val == '': return 0.0
                            return float(val)
                        return 0.0
                    except: return 0.0

                for year_label, col_name in [("2024", c_24), ("2025", c_25)]:
                    df_year = df_single[existing_base + [col_name]].copy()
                    df_year["公司"] = comp_name
                    df_year["报告年份"] = year_label
                    df_year["汇率"] = rate
                    
                    # 1. 先把所有文本转为干净的浮点数
                    cleaned_series = df_year[col_name].apply(clean_to_float)
                    
                    # 🌟 2. 核心修改：使用 isin() 进行【精确匹配】
                    # 使用 str.strip() 去除 Excel 单元格可能自带的头尾空格，防止因为多按了一个空格导致匹配失败
                    is_exempt = df_year["字段名"].astype(str).str.strip().isin(exact_exempt_fields)
                    
                    # 🌟 3. 条件赋值运算
                    df_year["(百万)原币"] = np.where(is_exempt, cleaned_series, cleaned_series * unit_mult)
                    df_year["(百万)人民币"] = np.where(is_exempt, cleaned_series, df_year["(百万)原币"] * rate)
                    
                    # 🔧 降级类型限制：把“汇率”列转为 object，使其可以同时容纳数字和文本
                    df_year["汇率"] = df_year["汇率"].astype(object)
                    # 如果是被豁免的行，强行把汇率说明改掉，方便底稿溯源
                    df_year.loc[is_exempt, "汇率"] = "豁免换算(绝对值项)"
                    
                    final_cols = ["公司", "类别", "字段名", "字段类型", "报告年份", "(百万)原币", "汇率", "(百万)人民币"]
                    actual_cols = [c for c in final_cols if c in df_year.columns]
                    combined_list.append(df_year[actual_cols])

            if combined_list:
                final_all_df = pd.concat(combined_list, ignore_index=True)
                st.session_state['integrated_data'] = final_all_df
                
                st.success(f"✅ 集成与换算完毕！共处理 {len(found_companies)} 家公司，生成 {len(final_all_df)} 条对标数据。")
                st.dataframe(final_all_df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_all_df.to_excel(writer, index=False, sheet_name='行业集成分析表')
                
                st.download_button(
                    label="📥 下载行业集成目标表 (长表格式)",
                    data=output.getvalue(),
                    file_name="行业集成目标数据表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.error("未能从上传的文件中提取到有效数据，请检查列名是否包含 2024/2025等年份。")

# ─────────── KPMG 官方色卡 ───────────
KPMG_CATEGORIES = {
    "Primary Colors": {"KPMG Blue": "#00338D", "Cobalt Blue": "#1E49E2", "Dark Blue": "#0C233C", "Light Blue": "#ACEAFF", "Pacific Blue": "#00B8F5", "Purple": "#7213EA", "Pink": "#FD349C"},
    "Accent Colors": {"Blue": "#76D2FF", "Dark Purple": "#510DBC", "Light Purple": "#B497FF", "Dark Pink": "#AB0D82", "Light Pink": "#FFA3DA", "Dark Green": "#098E7E", "Green": "#00C0AE", "Light Green": "#63EBB2"},
    "Traffic Light": {"Red": "#ED2124", "Amber": "#F1C44D", "Positive Green": "#269924"}
}
DEFAULT_COLORS = list(KPMG_CATEGORIES["Primary Colors"].values()) + list(KPMG_CATEGORIES["Accent Colors"].values())

with tab6:
    st.markdown("### 📊 可视化分析面板")
    
    if 'integrated_data' not in st.session_state: st.session_state['integrated_data'] = None
    source_choice = st.radio("数据源选择", ["直接引用集成后的数据", "上传集成表 Excel"], horizontal=True)
    df_raw = None
    if source_choice == "上传集成表 Excel":
        viz_file = st.file_uploader("上传行业集成目标表", type=["xlsx"])
        if viz_file: 
            df_raw = pd.read_excel(viz_file)
            st.session_state['integrated_data'] = df_raw
    else:
        df_raw = st.session_state.get('integrated_data')

    if df_raw is not None:
        df_clean = df_raw.copy()

        df_clean['报告年份'] = df_clean['报告年份'].astype(str).str.replace('.0', '', regex=False)
        val_col = "(百万)人民币" if "(百万)人民币" in df_clean.columns else df_clean.columns[-1]
        
        # 建立透视表
        df_pivot = df_clean.groupby(['公司', '报告年份', '字段名'])[val_col].sum().unstack('字段名').reset_index()
        all_fields = sorted(list(df_clean['字段名'].unique()))

        with st.expander("🎨 查看 KPMG 官方色卡"):
            for k, v in KPMG_CATEGORIES.items():
                st.markdown(f"**{k}**")
                html_str = "".join([f'<div style="display:inline-block; margin-right:15px; margin-bottom:8px;"><div style="width:14px; height:14px; background-color:{c}; display:inline-block; border-radius:3px; vertical-align:middle; border:1px solid #ddd;"></div><span style="font-size:13px; vertical-align:middle;"> {n} <b>({c})</b></span></div>' for n, c in v.items()])
                st.markdown(html_str, unsafe_allow_html=True)

        st.divider()
        
        # --- B. 交互控制区 ---
        with st.expander("🛠️ 核心配置面板", expanded=True):
            r1c1, r1c2, r1c3 = st.columns([1.5, 1, 1])
            
            with r1c1:
                chart_type = st.selectbox("📈 1. 图表类型", ["簇状柱状图", "堆积柱状图", "折线对比图", "饼图", "内外环结构对比图"])
                
                if "环" in chart_type or chart_type == "饼图":
                    calc_mode = "结构分析"
                    selected_multi_fields = st.multiselect("🎯 选择构成指标", all_fields, default=all_fields[:min(3, len(all_fields))])
                    selected_cos = st.selectbox("🏗️ 选择展示公司", sorted(df_pivot['公司'].unique().tolist()))
                    plot_df_base = df_clean[df_clean['字段名'].isin(selected_multi_fields) & (df_clean['公司'] == selected_cos)].copy()
                    plot_df_base = plot_df_base.rename(columns={val_col: 'final_val'})
                else:
                    calc_mode = st.radio("数据模式", ["单指标直显", "自定义公式运算"], horizontal=True)
                    if calc_mode == "单指标直显":
                        target_field = st.selectbox("🎯 选择显示指标", all_fields)
                        plot_df_base = df_pivot[['公司', '报告年份', target_field]].rename(columns={target_field: 'final_val'}).fillna(0)
                    else:
                        v1, v2 = st.columns(2)
                        var_a = v1.selectbox("变量 A", all_fields, index=0)
                        var_b = v2.selectbox("变量 B", ["无"] + all_fields, index=1 if len(all_fields)>1 else 0)
                        formula_input = st.text_input("✏️ 自定义公式 (使用 A 和 B)", value="A - B")
                        
                        # 🌟 修复后的公式引擎
                        try:
                            A_data = df_pivot[var_a].fillna(0)
                            B_data = df_pivot[var_b].fillna(0) if var_b != "无" else 0
                            
                            # 使用更安全的 pandas 内置 eval 或计算逻辑
                            # 将公式中的 A 和 B 替换为对应的数据列
                            res = pd.eval(formula_input, local_dict={'A': A_data, 'B': B_data})
                            df_pivot['final_val'] = res
                            df_pivot['final_val'] = df_pivot['final_val'].replace([np.inf, -np.inf], 0).fillna(0)
                            plot_df_base = df_pivot[['公司', '报告年份', 'final_val']].copy()
                        except Exception as e:
                            st.error(f"公式无效: {e}")
                            plot_df_base = pd.DataFrame(columns=['公司', '报告年份', 'final_val'])

                    selected_cos = st.multiselect("🏗️ 选择对比公司", sorted(df_pivot['公司'].unique().tolist()), default=sorted(df_pivot['公司'].unique().tolist())[:min(2, len(df_pivot['公司'].unique()))])

            with r1c2:
                x_axis_mode = st.radio("🔍 布局视角", ["以公司为横轴", "以年份为横轴"]) if "环" not in chart_type else "结构视角"
                decimals = st.number_input("🔢 小数位数", min_value=0, max_value=4, value=0)
                show_value = st.toggle("✅ 显示数据标签", value=True)
                
            with r1c3:
                # 🌟 十亿元回归
                unit_options = {"原始数值": 1.0, "亿元": 0.01, "十亿元": 0.001, "百分比(%)": 100.0}
                selected_unit = st.selectbox("📏 数值单位换算", list(unit_options.keys()))
                multiplier = unit_options[selected_unit]
                y_axis_title = st.text_input("📝 Y轴顶部单位说明", value=f"单位: {selected_unit.split(' ')[0]}")
                is_transparent = st.toggle("🌈 开启透明背景模式")
                show_avg = st.checkbox("平均值线 (仅柱状图)") if "柱" in chart_type else False
                avg_color = st.color_picker("基准线颜色", value="#ED2124") if show_avg else "#ED2124"

        # --- C. 重命名与颜色配置 (横向) ---
        if not plot_df_base.empty:
            if calc_mode == "结构分析":
                plot_df = plot_df_base.copy()
                color_val = "字段名"
            else:
                plot_df = plot_df_base[plot_df_base['公司'].isin(selected_cos)].copy()
                color_val = "报告年份" if x_axis_mode == "以公司为横轴" else "公司"

            plot_df['绘制金额'] = plot_df['final_val'] * multiplier
            
            st.markdown("#### 🎨 自定义图例标签与颜色")
            unique_items = sorted(plot_df[color_val].unique().tolist())
            rename_map = {}
            color_map = {}
            
            # 🌟 横向排列配置
            c_cols = st.columns(4)
            for i, item in enumerate(unique_items):
                with c_cols[i % 4]:
                    st.caption(f"原始值: {item}")
                    new_label = st.text_input(f"显示名称", value=str(item), key=f"rename_{item}")
                    new_color = st.color_picker(f"选择颜色", value=DEFAULT_COLORS[i % len(DEFAULT_COLORS)], key=f"color_{item}")
                    rename_map[item] = new_label
                    color_map[new_label] = new_color
            
            plot_df[color_val] = plot_df[color_val].map(rename_map)

            # --- D. 绘图引擎 ---
            fig = go.Figure()
            fmt = f'.{decimals}f'
            suffix = "%" if selected_unit == "百分比(%)" else ""
            label_fmt = f'%{{y:{fmt}}}{suffix}'

            if "柱状图" in chart_type:
                barmode = 'group' if "簇状" in chart_type else 'relative'
                for item in rename_map.values():
                    d = plot_df[plot_df[color_val] == item]
                    fig.add_trace(go.Bar(
                        x=d["公司" if color_val=="报告年份" else "报告年份"], 
                        y=d["绘制金额"], 
                        name=str(item), 
                        marker_color=color_map[item],
                        text=d["绘制金额"] if show_value else None,
                        texttemplate=label_fmt if show_value else None,
                        textposition='outside', # 强制在外面
                        textangle=0             # 强制横向
                    ))
                fig.update_layout(barmode=barmode)

            elif "折线对比图" in chart_type:
                for item in rename_map.values():
                    d = plot_df[plot_df[color_val] == item]
                    fig.add_trace(go.Scatter(
                        x=d["公司" if color_val=="报告年份" else "报告年份"], 
                        y=d["绘制金额"], 
                        name=str(item), 
                        mode='lines+markers+text' if show_value else 'lines+markers',
                        marker_color=color_map[item],
                        text=d["绘制金额"],
                        texttemplate=label_fmt,
                        textposition="top center"
                    ))

            elif chart_type == "饼图":
                latest = plot_df['报告年份'].max()
                d = plot_df[plot_df['报告年份'] == latest]
                fig = px.pie(d, values='绘制金额', names=color_val, hole=0.4, color=color_val, color_discrete_map=color_map)
                fig.update_traces(textinfo='percent+label' if show_value else 'percent')

            elif chart_type == "内外环结构对比图":
                years = sorted(plot_df['报告年份'].unique().tolist())
                if len(years) < 2: 
                    st.warning("环形图对比需要至少两年的数据。")
                else:
                    # 外环 (最新)
                    d_outer = plot_df[plot_df['报告年份'] == years[-1]]
                    fig.add_trace(go.Pie(labels=d_outer[color_val], values=d_outer['绘制金额'], hole=0.7, name=years[-1], 
                                         marker=dict(colors=[color_map[f] for f in d_outer[color_val]]),
                                         textinfo='percent+label' if show_value else 'percent'))
                    # 内环 (较早)
                    d_inner = plot_df[plot_df['报告年份'] == years[0]]
                    fig.add_trace(go.Pie(labels=d_inner[color_val], values=d_inner['绘制金额'], hole=0.4, name=years[0],
                                         domain={'x': [0.15, 0.85], 'y': [0.15, 0.85]}, 
                                         marker=dict(colors=[color_map[f] for f in d_inner[color_val]]),
                                         textinfo='percent' if show_value else 'none'))
                    fig.update_layout(annotations=[dict(text=f'内:{years[0]}<br>外:{years[-1]}', x=0.5, y=0.5, font_size=12, showarrow=False)])

            # --- E. 样式统一修缮 ---
            bg_color = "rgba(0,0,0,0)" if is_transparent else "white"
            if show_avg and "柱" in chart_type:
                avg_v = plot_df['绘制金额'].mean()
                fig.add_hline(y=avg_v, line_dash="dash", line_color=avg_color, 
                              annotation_text=f"平均: {avg_v:{fmt}}{suffix}",
                              annotation_font=dict(color=avg_color))

            fig.update_layout(
                font_family="Microsoft YaHei",
                plot_bgcolor=bg_color,
                paper_bgcolor=bg_color,
                margin=dict(t=120, l=10, r=10, b=20), # 顶部留白增加，彻底防止遮挡
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                annotations=[dict(
                    x=0, y=1.18, # 单位说明位置优化
                    xref='paper', yref='paper',
                    text=f"<b>{y_axis_title}</b>",
                    showarrow=False,
                    font=dict(size=14, color="#333"),
                    xanchor='left'
                )]
            )
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
            
            st.plotly_chart(fig, use_container_width=True)
            
            with st.expander("📄 查看底层数据明细"):
                st.dataframe(plot_df, use_container_width=True)
    else:
        st.info("💡 请先完成数据集成或上传目标底稿。")


# ─────────── Step 7 固定图的展示和 PPT ───────────
with tab7:
    st.markdown("### 🖼️ PPT展示 - 关键年报数据")

    # 1. 检查数据
    if 'integrated_data' not in st.session_state or st.session_state['integrated_data'] is None:
        st.warning("⚠️ 请先在 Step 6 完成数据集成。")
    else:
        df_raw = st.session_state['integrated_data'].copy()
        
        # --- ⚙️ 全局设置区 ---
        st.write("#### ⚙️ 图表设置")
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            all_cos = sorted(df_raw['公司'].unique())
            selected_cos = st.multiselect("选择展示的公司（按顺序）", options=all_cos, default=all_cos)
        with c2:
            # 单位选择器
            unit_label = st.selectbox("显示单位", ["十亿元", "亿元", "百万元","十万元"], index=0)
            unit_map = {"十亿元": 1000, "亿元": 100, "百万元": 1, "十万元": 0.1}
            divisor = unit_map[unit_label]
        with c3:
            # 字体大小
            font_size = st.slider("字体大小", 10, 20, 12)

        # --- 📈 核心绘图逻辑1.柱状图。保险服务收入、费用、业绩 ---
        def create_kpmg_chart(df, field_name, title_prefix, show_labels, pct_font_size, global_gap):
            # A. 数据清洗与计算
            d = df[df['公司'].isin(selected_cos)].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            
            if field_name == "保险服务业绩":
                df_inc = d[d['字段名'] == "保险服务收入合计"].set_index(['公司', '报告年份'])['(百万)人民币']
                df_exp = d[d['字段名'] == "保险服务费用合计"].set_index(['公司', '报告年份'])['(百万)人民币']
                res_series = df_inc - df_exp
                df_plot = res_series.reset_index(); df_plot.columns = ['公司', '报告年份', 'value']
            else:
                df_plot = d[d['字段名'] == field_name].copy()
                df_plot = df_plot.rename(columns={'(百万)人民币': 'value'})

            df_plot['value'] = df_plot['value'] / divisor
            years = sorted(df_plot['报告年份'].unique())
            if len(years) < 2:
                fig = go.Figure(); fig.add_annotation(text="数据年份不足", showarrow=False); return fig
            
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            color_map = {y_old: "#1E49E2", y_new: "#00338D"}

            #计算全图统一的偏移量
            all_max = df_plot['value'].max() if not df_plot['value'].empty else 100
            fixed_offset = all_max * 0.1 

            # B. 绘制柱状图
            for yr in [y_old, y_new]:
                df_yr = df_plot[df_plot['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", 
                    x=df_yr['公司'], 
                    y=df_yr['value'],
                    marker_color=color_map[yr], 
                    text=[f"{v:.1f}" if not np.isnan(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(family="Microsoft YaHei", size=font_size)
                ))

            # C. 添加涨幅箭头
            df_old = df_plot[df_plot['报告年份'] == y_old].set_index('公司')
            df_new = df_plot[df_plot['报告年份'] == y_new].set_index('公司')

            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v_old, v_new = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v_old) and pd.notna(v_new) and v_old != 0:
                        pct = (v_new - v_old) / abs(v_old)
                        color = "#FD349C" if pct >= 0 else "#269924"
                        arrow = "↗" if pct >= 0 else "↘"
                        
                        # 这里的 y 使用 max + fixed_offset，确保高矮柱子间距一样
                        fig.add_annotation(
                            x=co, 
                            y=max(v_old, v_new) + fixed_offset, 
                            text=f"<b>{arrow} {pct:.1%}</b>",
                            showarrow=False,
                            font=dict(family="Microsoft YaHei", color=color, size=pct_font_size),
                            xshift=15
                        )

            # D. 样式美化
            fig.update_layout(
                title=dict(text=f"<b>{title_prefix}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", 
                           font=dict(family="Microsoft YaHei", size=18, color="#00338D"), x=0.01),
                barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                bargroupgap=0.0,    # 组内间隙设为0，确保两年柱子贴在一起
                bargap=global_gap,  # 组间间隙，控制公司之间的距离。间隙越大柱子越细！
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(t=120, b=40, l=20, r=20), height=500
            )
            fig.update_xaxes(showgrid=False, zeroline=False, showline=False, linecolor='#333333')
            fig.update_yaxes(showgrid=False, zeroline=True, showline=False,zerolinecolor='#333333')
            
            return fig


        # --- 📈 核心绘图逻辑 2：堆积图。保险业务构成-PAA与非PAA---
        def create_kpmg_composition_chart(df, fields, title_prefix, show_labels, label_size, bar_width, co_font_size):
            # A. 数据过滤
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            
            if len(available_cos) == 0:
                fig = go.Figure(); fig.add_annotation(text="无对应数据", showarrow=False); return fig

            # B. 创建分面子图
            fig = make_subplots(
                rows=1, cols=len(available_cos),
                shared_yaxes=True,
                # 💡 这里的 co_font_size 控制顶部公司名的大小
                column_titles=[f"<span style='font-size:{co_font_size}px;'>{co}</span>" for co in available_cos],
                horizontal_spacing=0.015
            )

            short_names = {fields[0]: "采用保费分配法保险合同的保险服务收入", fields[1]: "未采用保费分配法保险合同的保险服务收入"}

            # C. 遍历绘制
            for i, co in enumerate(available_cos):
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                
                d_co = d_co.reindex(sorted(d_co.index.tolist()))
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                
                paa_val = d_co[fields[0]] / d_co['Total'] * 100
                non_paa_val = d_co[fields[1]] / d_co['Total'] * 100
                
                # 下层：未采用
                fig.add_trace(
                    go.Bar(
                        x=d_co.index, y=non_paa_val,
                        name=short_names[fields[1]] if i == 0 else None,
                        marker_color="rgb(227,207,251)",
                        # 💡 调节数据标签大小和柱子宽度
                        text=[f"{v:.0f}%" if v > 0 else "" for v in non_paa_val] if show_labels else None,
                        textposition='inside', insidetextanchor='middle',
                        textfont=dict(size=label_size), 
                        textangle=0,  
                        constraintext='none',
                        width=bar_width,
                        showlegend=(True if i == 0 else False), legendgroup="group2",
                        hoverinfo="none"
                    ), row=1, col=i+1
                )
                
                # 上层：采用
                fig.add_trace(
                    go.Bar(
                        x=d_co.index, y=paa_val,
                        name=short_names[fields[0]] if i == 0 else None,
                        marker_color="#510DBC",
                        text=[f"{v:.0f}%" if v > 0 else "" for v in paa_val] if show_labels else None,
                        textposition='inside', insidetextanchor='middle',
                        textfont=dict(size=label_size),
                        textangle=0,  
                        constraintext='none',
                        width=bar_width,
                        showlegend=(True if i == 0 else False), legendgroup="group1",
                        hoverinfo="none"
                    ), row=1, col=i+1
                )

            # D. 样式修缮
            fig.update_layout(
                title=dict(text=f"<b>{title_prefix}</b>", font=dict(family="Microsoft YaHei", size=18, color="#00338D")),
                barmode='stack', height=500,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=80, l=20, r=20),
                legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
                font_family="Microsoft YaHei"
            )
            #调整公司名称和柱子上边界的距离，0.05为更远
            for ann in fig.layout.annotations:
                ann.y = ann.y + 0.03  
            for i in range(1, len(available_cos) + 1):
                # 💡 调节 X 轴年份字体
                fig.update_xaxes(showgrid=False, showline=True, linecolor='#E0E0E0', tickfont=dict(size=10), row=1, col=i)
                fig.update_yaxes(showgrid=False, range=[0, 100], showline=True, linecolor='#E0E0E0', 
                                 showticklabels=(True if i==1 else False), row=1, col=i)
            
            return fig
        
        # --- 📈 核心绘图逻辑 3：堆积图。保险业务收入构成-来源构成 ---
        def create_kpmg_multi_composition_chart(df, field_map, color_map, title_prefix, show_labels, label_size, bar_width, co_font_size):
            # A. 数据过滤
            fields = list(field_map.keys())
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            
            if len(available_cos) == 0:
                fig = go.Figure(); fig.add_annotation(text="无对应数据", showarrow=False); return fig

            # B. 创建分面子图
            fig = make_subplots(
                rows=1, cols=len(available_cos),
                shared_yaxes=True,
                column_titles=[f"<span style='font-family: Microsoft YaHei; font-size:{co_font_size}px;'>{co}</span>" for co in available_cos],
                horizontal_spacing=0.015
            )

            # C. 遍历绘制
            for i, co in enumerate(available_cos):
                # 提取该公司数据并透视
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                
                # 补全缺失列
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                
                years_in_data = sorted(d_co.index.tolist())
                d_co = d_co.reindex(years_in_data)
                
                # 计算总和与百分比
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                
                # 倒序添加 Trace（为了让第一个指标显示在最下面，符合堆积习惯，或者按你要求的顺序）这里按 field_map 的顺序从下往上堆叠
                for field_name, display_name in field_map.items():
                    val_pct = d_co[field_name] / d_co['Total'] * 100
                    
                    fig.add_trace(
                        go.Bar(
                            x=d_co.index, y=val_pct,
                            name=display_name if i == 0 else None,
                            marker_color=color_map[field_name],
                            text=[f"{v:.0f}%" if v >= 1 else "" for v in val_pct] if show_labels else None,
                            textposition='inside', insidetextanchor='middle',
                            textangle=0,
                            textfont=dict(size=label_size, color="white" if "30, 73, 226" in color_map[field_name] or "114, 19, 234" in color_map[field_name] else "black"),
                            constraintext='none', # 保证小比例下文字大小不缩水
                            width=bar_width,
                            showlegend=(True if i == 0 else False),
                            legendgroup=field_name,
                            hoverinfo="none"
                        ), row=1, col=i+1
                    )

            # D. 样式修缮
            fig.update_layout(
                title=dict(text=f"<b>{title_prefix}</b>", font=dict(family="Microsoft YaHei", size=18, color="#00338D")),
                barmode='stack', height=600, # 维度多，高度建议增加
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=100, l=20, r=20),
                legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5, font=dict(size=10)),
                font_family="Microsoft YaHei"
            )

            # 提升公司名高度
            for ann in fig.layout.annotations:
                ann.y = ann.y + 0.05  

            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(showgrid=False, showline=True, linecolor='#E0E0E0', tickfont=dict(size=10), row=1, col=i)
                fig.update_yaxes(showgrid=False, range=[0, 100], showline=True, linecolor='#E0E0E0', 
                                 showticklabels=(True if i==1 else False), row=1, col=i)
            
            return fig        
 
        
        # --- 📈 核心绘图逻辑 4：利润贡献拆解图 (相对比例版) ---
        def create_profit_composition_chart(df, selected_cos, target_year, show_labels, label_size, bar_width, co_font_size):
            # 1. 明确源数据所需的列名
            source_fields = [
                "合同服务边际的摊销", 
                "非金融风险调整的变动", 
                "亏损部分的确认及转回", 
                "采用保费分配法计量的保险合同保险业绩", 
                "保险利润", 
                "再保净损益"
            ]
            
            # 2. 筛选并透视数据
            year_str = str(target_year).replace(".0", "")
            d_sub = df[(df['报告年份'].astype(str).str.contains(year_str)) & (df['公司'].isin(selected_cos))].copy()
            d_sub = d_sub[d_sub['字段名'].isin(source_fields)]
            
            if d_sub.empty:
                return None
                
            d_pivot = d_sub.pivot(index='公司', columns='字段名', values='(百万)人民币')
            
            # 补全缺失的源数据列，防止报错
            for f in source_fields:
                if f not in d_pivot.columns:
                    d_pivot[f] = 0.0
                    
            # 按选择的公司顺序排序
            d_pivot = d_pivot.reindex(selected_cos).fillna(0)
        
            # 3. 🎯 按照你的公式计算展示项 (Target Fields)
            # 计算营运偏差及其他 (使用源数据计算)
            d_pivot['营运偏差及其他_展示'] = (
                d_pivot['保险利润'] 
                - d_pivot['合同服务边际的摊销'] 
                - d_pivot['非金融风险调整的变动'] 
                + d_pivot['亏损部分的确认及转回'] 
                - d_pivot['采用保费分配法计量的保险合同保险业绩'] 
                - d_pivot['再保净损益']
            )
            
            # 映射其他展示列
            d_pivot['合同服务边际的释放_展示'] = d_pivot['合同服务边际的摊销']
            d_pivot['非金融风险调整的变动_展示'] = d_pivot['非金融风险调整的变动']
            d_pivot['亏损部分的确认及转回_展示'] = -d_pivot['亏损部分的确认及转回'] # 💡 取负数
            d_pivot['保费分配法业务净损益_展示'] = d_pivot['采用保费分配法计量的保险合同保险业绩']
            d_pivot['再保净损益_展示'] = d_pivot['再保净损益']
        
            # 4. 颜色与绘制顺序配置
            # 💡 为了实现“从0开始向上堆叠，且图例顺序如你所愿”，我们需要精心设计 Trace 的添加顺序。
            # 图像中 CSM 在最底端，所以要先画。然后我们用 legend 的 traceorder='reversed' 把图例倒过来。
            display_mapping = [
                ("亏损部分的确认及转回_展示", "亏损部分的确认及转回", "rgb(190, 190, 190)"),
                ("合同服务边际的释放_展示", "合同服务边际的释放", "rgb(30, 73, 226)"),
                ("非金融风险调整的变动_展示", "非金融风险调整的变动", "rgb(118, 210, 255)"),
                ("营运偏差及其他_展示", "营运偏差及其他", "rgb(114, 19, 214)"),
                ("保费分配法业务净损益_展示", "保费分配法业务净损益", "rgb(253, 52, 156)"),
                ("再保净损益_展示", "再保净损益", "rgb(9, 142, 126)")
                 # 这个通常在负数区
            ]
            
            # 5. 计算分母：绝对值之和
            target_cols = [item[0] for item in display_mapping]
            d_pivot['Total'] = d_pivot[target_cols].abs().sum(axis=1).replace(0, 1)
        
            fig = go.Figure()
        
            # 6. 循环绘图
            for col_name, legend_name, color in display_mapping:
                vals_pct = (d_pivot[col_name] / d_pivot['Total']) * 100
                
                # 字体颜色逻辑：深色背景用白字
                txt_color = "white" if "30, 73, 226" in color or "114, 19, 214" in color or "9, 142, 126" in color else "black"
                
                fig.add_trace(go.Bar(
                    name=legend_name,
                    x=d_pivot.index,
                    y=vals_pct,
                    width=bar_width,
                    marker_color=color,
                    text=[f"{v:.0f}%" if abs(v) >= 2 else "" for v in vals_pct] if show_labels else None,
                    textposition='inside',
                    insidetextanchor='middle',
                    textfont=dict(size=label_size, color=txt_color),
                    constraintext='none' # 保证字体不缩小
                ))
        
            # 7. 样式与布局设置
            fig.update_layout(
                title=dict(
                    text=f"<b>{year_str}年各公司保险利润构成</b>",
                    x=0.5, y=0.95, xanchor='center',
                    font=dict(size=18, color='#00338D')
                ),
                barmode='relative', # 正负堆叠
                height=600,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)', # 极浅的灰色背景
                margin=dict(t=100, b=80, l=220, r=60), # 💡 左侧留出 220 像素放图例
                
                # 💡 图例设置：左侧、垂直排列、倒序排列(以完美匹配你的要求顺序)
                legend=dict(
                    orientation="v",
                    yanchor="middle", y=0.5, 
                    xanchor="right", x=-0.05, # x为负数，移到图表坐标轴左侧外面
                    traceorder="reversed",    # 反转顺序，让“再保”排第一
                    font=dict(size=11, color="#00338D")
                ),
                
                # 💡 Y轴设置：显示在右侧 (side='right')
                yaxis=dict(
                    side='right', 
                    showgrid=False, 
                    range=[-45, 105], # 底部多留点空间给负数
                    tickvals=[-40, -20, 0, 20, 40, 60, 80, 100],
                    ticktext=["-40%", "-20%", "0%", "20%", "40%", "60%", "80%", "100%"],
                    tickfont=dict(color="#666666"),
                    zeroline=True, 
                    zerolinecolor='orange', # 💡 橙黄色的 0 刻度线
                    zerolinewidth=2
                )
            )
        
            # 替换X轴标签，实现自定义公司名字号
            x_labels = [f"<span style='font-size:{co_font_size}px;color:#00338D;'>{co}</span>" for co in d_pivot.index]
            fig.update_xaxes(showgrid=False, tickvals=d_pivot.index, ticktext=x_labels,side="top")
        
            return fig
        
        
        # --- 📈 核心绘图逻辑 5：投资收益等图 ---    
        target_fields = ['净投资回报', '承保财务损益', '分出再保险财务损益']
        df_clean = df_raw[df_raw['字段名'].isin(target_fields)].drop_duplicates(
            subset=['公司', '报告年份', '字段名'], keep='first'
        )

        # 透视为宽表：行是公司和年份，列是那三个项目
        df_pivot = df_clean.pivot_table(
            index=['公司', '报告年份'], 
            columns='字段名', 
            values='(百万)人民币'
        ).fillna(0).reset_index()

        # 图1：直接用原表的‘净投资回报’
        # 图2：公式：承保财务净损益 = -(承保财务损益) - 分出再保险财务损益
        df_pivot['承保财务净损益'] = -df_pivot['承保财务损益'] - df_pivot['分出再保险财务损益']

        # 图3：公式：投资利润 = 净投资回报 - 承保财务损益
        df_pivot['投资利润'] = df_pivot['净投资回报'] - df_pivot['承保财务净损益']

        # 构造最终绘图用的长表
        df_plot_final = df_pivot.melt(
            id_vars=['公司', '报告年份'], 
            value_vars=['净投资回报', '承保财务净损益', '投资利润'],
            var_name='指标名称', value_name='value'
        )
        #该PPT画图函数
        def create_simple_kpmg_chart(df_source, field_name, title, show_labels, p_size, g_gap):
            # 筛选指标
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor # 转换单位
            
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            
            y_old, y_new = years[-2], years[-1]
            # 🎨 旧年粉色，新年酒红
            color_map = {y_old: "#FD349C", y_new: "#97014F"}
            
            fig = go.Figure()
            for yr in [y_old, y_new]:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'],
                    marker_color=color_map[yr],
                    text=[f"{v:.1f}" if not np.isnan(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size)
                ))

            # 涨幅箭头逻辑
            all_max = d['value'].max() if not d['value'].empty else 100
            off = all_max * 0.12
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr, col = ("↗", "#FD349C") if pct >= 0 else ("↘", "#269924")
                        fig.add_annotation(x=co, y=max(v0, v1) + off, text=f"<b>{arr} {pct:.1%}</b>",
                                         showarrow=False, font=dict(color=col, size=p_size), xshift=10)

            fig.update_layout(
                title=dict(text=f"<b>{title}</b>", font=dict(size=18, color="#00338D")),
                barmode='group', bargroupgap=0.0, bargap=g_gap,
                margin=dict(t=80, b=40, l=20, r=20), height=500,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)'
            )
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333')
            return fig
        
        
        # --- 📈 核心绘图逻辑 6：保险利润和投资利润等图 ---         
        # --- 🧮 1. 提取原始数据中的“投资利润”和“保险利润” ---
        target_profits = ['投资利润', '保险利润']
        
        # 提取并去重（确保每公司每年每个指标只有一行数据）
        df_profit_raw = df_raw[
            (df_raw['字段名'].isin(target_profits)) & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        
        # 统一列名方便后续函数调用
        df_profit_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)
        #函数内容:
        def create_profit_chart_v4(df_source, field_name, title, color_dict, show_labels, p_size, g_gap):
            # 筛选特定指标
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor # 单位转换
            
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            
            fig = go.Figure()
            
            # 绘制柱状图
            for yr in [y_old, y_new]:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'],
                    marker_color=color_dict[yr], # 应用传入的专属颜色
                    text=[f"{v:.1f}" if not np.isnan(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size)
                ))
        
            # 涨幅箭头
            all_max = d['value'].max() if not d['value'].empty else 100
            off = all_max * 0.12
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr, col = ("↗", "#FD349C") if pct >= 0 else ("↘", "#269924")
                        fig.add_annotation(x=co, y=max(v0, v1) + off, text=f"<b>{arr} {pct:.1%}</b>",
                                         showarrow=False, font=dict(color=col, size=p_size), xshift=10)
        
            fig.update_layout(
                title=dict(text=f"<b>{title}</b>", font=dict(size=18, color="#00338D")),
                barmode='group', bargroupgap=0.0, bargap=g_gap,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=80, b=40, l=20, r=20), height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333')
            return fig
        
        # --- 📈 核心绘图逻辑 7：净利润，税前利润，所得税等图 ---    
        def create_tax_subplot_chart(df_pivot, selected_cos, show_labels, bar_width, label_size, co_font_size, co_y_offset):
            from plotly.subplots import make_subplots
            import plotly.graph_objects as go
        
            available_cos = [co for co in selected_cos if co in df_pivot['公司'].unique()]
            if not available_cos: return go.Figure()
        
            # 计算全局最高值
            global_max = (df_pivot[['税前利润总额', '净利润']].max().max() / divisor)
            y_axis_limit = global_max * 1.2 
            tax_label_y = global_max * 1.17 
        
            fig = make_subplots(
                rows=1, cols=len(available_cos),
                shared_yaxes=True,
                column_titles=[f"<b><span style='font-size:{co_font_size}px;'>{co}</span></b>" for co in available_cos],
                horizontal_spacing=0.03
            )
        
            for i, co in enumerate(available_cos):
                d_co = df_pivot[df_pivot['公司'] == co].sort_values('报告年份')
                
                # 税前利润
                fig.add_trace(
                    go.Bar(
                        x=d_co['报告年份'], y=d_co['税前利润总额'] / divisor,
                        marker_color="#1E49E2",
                        text=[f"{v/divisor:.1f}" for v in d_co['税前利润总额']] if show_labels else None,
                        textposition='outside', textfont=dict(size=label_size),
                        width=bar_width, offsetgroup=1
                    ), row=1, col=i+1
                )
                
                # 净利润
                fig.add_trace(
                    go.Bar(
                        x=d_co['报告年份'], y=d_co['净利润'] / divisor,
                        marker_color="#C7A0F7",
                        text=[f"{v/divisor:.1f}" for v in d_co['净利润']] if show_labels else None,
                        textposition='outside', textfont=dict(size=label_size),
                        width=bar_width, offsetgroup=2
                    ), row=1, col=i+1
                )
        
                # 税率标注
                for idx, row in d_co.iterrows():
                    rate = row['有效税率']
                    t_color = "#97014F" if rate >= 0 else "#269924"
                    fig.add_annotation(
                        x=row['报告年份'], y=tax_label_y,
                        text=f"<b>{rate:.0%}</b>",
                        showarrow=False,
                        font=dict(size=label_size + 2, color=t_color),
                        xref=f"x{i+1}", yref=f"y1"
                    )
        
            fig.update_layout(
                title=dict(text="<b>关键年报数据 - 税前利润和净利润变动趋势</b>", font=dict(size=20, color="#00338D"), x=0.01),
                height=550, margin=dict(t=120, b=50, l=40, r=40),
                # 💡 这里改为透明
                plot_bgcolor='rgba(0,0,0,0)', 
                paper_bgcolor='rgba(0,0,0,0)',
                showlegend=False
            )
        
            for i in range(1, len(available_cos) + 1):
                # 灰色边框保留，因为透明背景下它们是区分公司的关键
                fig.update_xaxes(showline=True, linecolor='#D3D3D3', showgrid=False, tickfont=dict(size=10), row=1, col=i)
                fig.update_yaxes(showline=True, linecolor='#D3D3D3', showgrid=False, showticklabels=False, range=[0, y_axis_limit], row=1, col=i)
            
            for ann in fig.layout.annotations:
                if "span" in ann.text:
                    ann.y = co_y_offset 
        
            return fig
        # --- 数据准备 ---
        target_tax_items = ['税前利润总额', '净利润']
        df_tax_sub = df_raw[
            (df_raw['字段名'].isin(target_tax_items)) & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()

        df_tax_pivot = df_tax_sub.pivot_table(
            index=['公司', '报告年份'], columns='字段名', values='(百万)人民币'
        ).fillna(0).reset_index()

        df_tax_pivot['有效税率'] = np.where(
            df_tax_pivot['税前利润总额'] != 0,
            (df_tax_pivot['税前利润总额'] - df_tax_pivot['净利润']) / df_tax_pivot['税前利润总额'], 0
        )
        df_tax_pivot['报告年份'] = df_tax_pivot['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"

  

       # --- 📈 核心绘图逻辑 8：净利润，其他综合收益等图 ---   
        def create_financial_trend_chart_v5(df_source, field_name, title, show_labels, p_size, g_gap, current_divisor, current_unit_label):
            import plotly.graph_objects as go
            import numpy as np
            import pandas as pd
        
            # 1. 筛选与单位转换
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            
            # 💡 关键修改：使用传入的全局 divisor 进行单位缩放
            d['value'] = d['value'] / current_divisor 
            
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            
            color_new = "rgb(0, 51, 141)"   
            color_old = "rgb(15, 101, 253)" 
            
            fig = go.Figure()
            
            for yr, col in zip([y_old, y_new], [color_old, color_new]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", 
                    x=df_yr['公司'], 
                    y=df_yr['value'],
                    marker_color=col,
                    text=[f"{v:.1f}" if not np.isnan(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size),
                    cliponaxis=False
                ))
            
            # 2. 涨幅箭头镜像逻辑（计算偏移量时也已经基于新单位）
            all_vals = d['value'].dropna()
            max_range = all_vals.max() - all_vals.min() if not all_vals.empty else 100
            off = max_range * 0.12 
            
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color = "#FD349C" if pct >= 0 else "#269924"
                        arr_symbol = "↗" if pct >= 0 else "↘"
                        
                        if v1 >= 0:
                            y_pos = max(v0, v1) + off
                            v_align = "bottom"
                        else:
                            y_pos = min(v0, v1) - off
                            v_align = "top"
                        
                        fig.add_annotation(
                            x=co, y=y_pos,
                            text=f"<b>{arr_symbol} {pct:.1%}</b>",
                            showarrow=False,
                            font=dict(color=arr_color, size=p_size),
                            xshift=10, 
                            valign=v_align
                        )
            
            # 3. 布局与单位显示
            fig.update_layout(
                # 💡 关键修改：在标题下加入动态单位标签
                title=dict(
                    text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({current_unit_label}人民币)</span>", 
                    font=dict(size=18, color="#00338D"), 
                    x=0.01
                ),
                barmode='group', 
                bargroupgap=0,  
                bargap=g_gap,     
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=60, l=40, r=40), # 增加顶部边距以容纳两行标题
                height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
        
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333', zerolinewidth=1)
            fig.update_xaxes(showline=False, showgrid=False)
            
            return fig
        # --- 数据预处理 ---
        target_metrics = ['净利润', '其他综合收益', '综合收益总额']
        
        df_fin_raw = df_raw[
            (df_raw['字段名'].isin(target_metrics)) & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        
        # 统一格式
        df_fin_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)


       # --- 📈 核心绘图逻辑 9：I19资产端分类图 --- 
        def create_asset_composition_chart(df, selected_cos, field_map, color_map, title_prefix, show_labels, label_size, bar_width, co_font_size):
            fields = list(field_map.keys())
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            if len(available_cos) == 0:
                fig = go.Figure(); fig.add_annotation(text="无对应数据", showarrow=False); return fig
        
            fig = make_subplots(
                rows=1, cols=len(available_cos), shared_yaxes=True,
                column_titles=[f"<span style='font-size:{co_font_size}px; color:#00338D;'>{co}</span>" for co in available_cos],
                horizontal_spacing=0.015
            )
        
            for i, co in enumerate(available_cos):
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                d_co = d_co.reindex(sorted(d_co.index.tolist()))
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                
                for field_name in fields:
                    val_pct = d_co[field_name] / d_co['Total'] * 100
                    # 颜色亮度逻辑
                    txt_color = "white" if field_name in ["FVOCI", "指定FVOCI", "AC"] else "black"
                    
                    # 💡 关键：为图例名字包裹 <center> 标签实现居中
                    display_name = field_map[field_name]
                    
                    fig.add_trace(
                        go.Bar(
                            x=d_co.index, y=val_pct,
                            name=display_name if i == 0 else None,
                            marker_color=color_map[field_name],
                            text=[f"{v:.0f}%" for v in val_pct] if show_labels else None,
                            textposition='inside', 
                            insidetextanchor='middle',
                            textangle=0,
                            textfont=dict(size=label_size, color=txt_color),
                            # 💡 核心：禁止 Plotly 根据条形大小约束文字大小
                            constraintext='none', 
                            width=bar_width,
                            showlegend=(True if i == 0 else False),
                            legendgroup=field_name,
                            hoverinfo="skip"
                        ), row=1, col=i+1
                    )
        
            fig.update_layout(
                title=dict(text=f"<b>{title_prefix}</b>", font=dict(size=18, color="#00338D")),
                barmode='stack', height=600,
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                # 💡 设置 uniformtext 确保所有文字尽量一致
                uniformtext=dict(minsize=label_size, mode='show'),
                margin=dict(t=120, b=120, l=40, r=40),
                legend=dict(
                    orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5, 
                    font=dict(size=11),
                    itemsizing="constant",
                    traceorder="normal"
                )
            )
        
            # 上移公司名
            for ann in fig.layout.annotations:
                ann.y = ann.y + 0.05 
            
            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(showgrid=False, showline=False, linecolor='#E0E0E0', tickfont=dict(size=10, color="#00338D"), row=1, col=i)
                fig.update_yaxes(
                    showgrid=False, range=[0, 101], # 稍微多给1%，防止顶端切断
                    showline=False, 
                    tickvals=[0, 25, 50, 75, 100], 
                    ticktext=["0%", "25%", "50%", "75%", "100%"],
                    showticklabels=(True if i==1 else False), row=1, col=i
                )
            
            return fig
        
      # --- 📈 核心绘图逻辑 10：其他综合收益（1/2）图 ---   

        def create_oci_chart(df_raw, year, title, show_labels, co_font_size, bar_gap, unit_divisor, unit_str, selected_cos):
            """
            综合收益变动情况拆解图
            自带公式计算与去重，适配全局单位调节与透明背景
            """
            # 1. 提取数据与去重
            oci_fields = ['可转损益OCI合计', '不可转损益OCI合计', '净利润', '综合收益总额']
            
            # 兼容传入的 year 可能是字符串或浮点数，统一转格式匹配
            df_raw_copy = df_raw.copy()
            df_raw_copy['报告年份'] = df_raw_copy['报告年份'].astype(str).str.replace(".0", "", regex=False)
            target_year = str(year).replace(".0", "")
            
            df_sub = df_raw_copy[(df_raw_copy['字段名'].isin(oci_fields)) & (df_raw_copy['报告年份'] == target_year)].copy()
            df_sub = df_sub.drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='last')
            
            # 如果该年份无数据，返回空图
            if df_sub.empty:
                return go.Figure().update_layout(title=f"<b style='color:#00338D;'>{title}</b><br>无数据", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
        
            # 2. 数据透视
            df_pivot = df_sub.pivot_table(index='公司', columns='字段名', values='(百万)人民币', aggfunc='sum').reset_index().fillna(0)
            
            # 补全缺失列（防止报错）
            for f in oci_fields:
                if f not in df_pivot.columns: df_pivot[f] = 0.0
        
            # 3. 公式计算与全局单位换算
            df_pivot['净利润'] = df_pivot['净利润'] / unit_divisor
            df_pivot['可转损益OCI合计'] = df_pivot['可转损益OCI合计'] / unit_divisor
            df_pivot['不可转损益OCI合计'] = df_pivot['不可转损益OCI合计'] / unit_divisor
            df_pivot['综合收益总额'] = df_pivot['综合收益总额'] / unit_divisor
            
            # 💡 其他 = 综合收益 - 净利润 - 可转 - 不可转
            df_pivot['其他'] = df_pivot['综合收益总额'] - df_pivot['净利润'] - df_pivot['可转损益OCI合计'] - df_pivot['不可转损益OCI合计']
        
            # 4. 图例颜色配置
            metrics_config = {
                "净利润": {"color": "rgb(0,176,240)", "name": "净利润"},
                "可转损益OCI合计": {"color": "rgb(253,52,156)", "name": "可转损益OCI变动+FVOCI债券公允价值变动"},
                "不可转损益OCI合计": {"color": "rgb(114,19,234)", "name": "不可转损益负债OCI变动+FVOCI股权公允价值变动"},
                "其他": {"color": "rgb(127,127,127)", "name": "其他"},
                "综合收益总额": {"color": "rgb(172,234,255)", "name": "综合收益总额"}
            }
        
            # 5. 过滤出选中的公司并保持顺序
            df_pivot = df_pivot.set_index('公司')
            valid_companies = [c for c in selected_cos if c in df_pivot.index]
            num_cols = len(valid_companies)
            if num_cols == 0:
                 return go.Figure().update_layout(title=f"<b style='color:#00338D;'>{title}</b><br>无所选公司数据", paper_bgcolor='rgba(0,0,0,0)')
            plot_cols = ['净利润', '可转损益OCI合计', '不可转损益OCI合计', '其他', '综合收益总额']
            all_vals = df_pivot[plot_cols].values.flatten()
            max_val = np.nanmax(all_vals) if not np.isnan(all_vals).all() else 10
            min_val = np.nanmin(all_vals) if not np.isnan(all_vals).all() else -10
            
            # 增加 25% 的 buffer，留出给数据标签的空间
            buffer = (max_val - min_val) * 0.25
            y_range = [
                min_val - buffer if min_val < 0 else 0,  # 如果有负数则向下延伸
                max_val + buffer                         # 向上延伸
            ]        
            # 6. 开始画分面图
            fig = make_subplots(
                rows=1, cols=num_cols, shared_yaxes=True,
                horizontal_spacing=0.015, subplot_titles=[""] * num_cols
            )
        
            for col_idx, co in enumerate(valid_companies):
                co_data = df_pivot.loc[co] if isinstance(df_pivot.loc[co], pd.Series) else df_pivot.loc[co].iloc[0]
                
                for m_key, m_info in metrics_config.items():
                    val = co_data.get(m_key, 0)
                    text_label = f"{val:.0f}" if (show_labels and not pd.isna(val) and val != 0) else ""
                    show_legend = True if col_idx == 0 else False
                    
                    fig.add_trace(
                        go.Bar(
                            name=m_info["name"], x=[m_key], y=[val], text=[text_label],
                            textposition='outside', textfont=dict(size=11, color='#00338D'),
                            marker_color=m_info["color"], width=0.8,
                            legendgroup=m_key, showlegend=show_legend
                        ),
                        row=1, col=col_idx + 1
                    )
                    
                # 设置 x, y 轴边框与公司名
                fig.update_xaxes(
                    title_text=co, title_font=dict(size=co_font_size, color="#333"), 
                    showticklabels=False, showline=True, linecolor='lightgrey', linewidth=1, mirror=True,
                    row=1, col=col_idx + 1
                )
                fig.update_yaxes(
                    showline=True, linecolor='lightgrey', linewidth=1, mirror=True,
                    zeroline=True, zerolinecolor='lightgrey', zerolinewidth=1, showgrid=False,
                    row=1, col=col_idx + 1
                )
        
            # 7. 整体布局 (背景透明、右侧单位对齐)
            fig.update_layout(
                title=dict(text=f"<b style='color:#00338D;'>{title}</b>", font=dict(size=18)),
                paper_bgcolor='rgb(243, 244, 246)', 
                plot_bgcolor='rgba(0,0,0,0)',
                barmode='group',
                bargap=bar_gap,
                height=450, 
                margin=dict(t=60, b=40, l=40, r=20),
                legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5, font=dict(size=11), traceorder="normal"),
                annotations=[dict(
                    text=f"单位：{unit_str}人民币",
                    xref="paper", yref="paper",
                    x=1.0, y=1.12, showarrow=False,
                    font=dict(size=12, color="#666"), xanchor="right"
                )]
            )
            fig.update_yaxes(
                range=y_range, # 使用刚才计算出的带 buffer 的范围
                showline=True, linecolor='lightgrey', linewidth=1, mirror=True,
                zeroline=True, zerolinecolor='lightgrey', zerolinewidth=1, showgrid=False,
                row=1, col="all" # 对所有列生效
            )
            return fig
        
      # --- 📈 核心绘图逻辑 11：其他综合收益（2/2）图 ---     
        def create_asset_liab_oci_chart(df_raw, selected_cos, unit_divisor, unit_str, bar_gap, co_font_size, show_labels):
            """
            创建负债OCI与FVOCI债权变动对比分面图
            优化点：单位联动、无网格、图例下移、标签控制
            """
            fields = ['可转损益的负债OCI', 'FVOCI债券公允价值']
            df_sub = df_raw[df_raw['字段名'].isin(fields)].copy()
            df_sub['报告年份'] = df_sub['报告年份'].astype(str).str.replace(".0", "", regex=False)
            
            # 筛选最近两年
            years = sorted(df_sub['报告年份'].unique(), reverse=True)[:2]
            years = sorted(years) # 顺序显示，如 2024, 2025
            
            num_cos = len(selected_cos)
            fig = make_subplots(
                rows=1, cols=num_cos, shared_yaxes=True,
                horizontal_spacing=0.01, subplot_titles=[""] * num_cos
            )
        
            colors = {"可转损益的负债OCI": "rgb(0, 184, 245)", "FVOCI债券公允价值": "rgb(253, 52, 156)"}
        
            # 计算全局 Y 轴范围并预留标签空间 (25% buffer)
            all_vals = df_sub[df_sub['公司'].isin(selected_cos)]['(百万)人民币'] / unit_divisor
            if not all_vals.empty:
                v_max, v_min = all_vals.max(), all_vals.min()
                y_range = [v_min - abs(v_min)*0.25 if v_min < 0 else 0, v_max + abs(v_max)*0.25]
            else:
                y_range = [0, 10]
        
            for idx, co in enumerate(selected_cos):
                co_df = df_sub[df_sub['公司'] == co]
                
                for f_name in fields:
                    f_vals, f_texts, f_x = [], [], []
                    for y in years:
                        val_row = co_df[(co_df['报告年份'] == y) & (co_df['字段名'] == f_name)]
                        val = val_row['(百万)人民币'].iloc[0] / unit_divisor if not val_row.empty else 0
                        f_vals.append(val)
                        # 💡 标签控制
                        f_texts.append(f"{val:.0f}" if (show_labels and val != 0) else "")
                        f_x.append(f"{y}")
        
                    fig.add_trace(
                        go.Bar(
                            name=f_name, x=f_x, y=f_vals, text=f_texts,
                            textposition='outside', textfont=dict(size=11, color='#00338D'),
                            marker_color=colors[f_name],
                            legendgroup=f_name, showlegend=(idx == 0),
                            width=0.4
                        ),
                        row=1, col=idx + 1
                    )
                
                # 💡 去除网格线 (showgrid=False)
                fig.update_xaxes(
                    title_text=co, title_font=dict(size=co_font_size), 
                    row=1, col=idx+1, showline=True, linecolor='lightgrey', mirror=True
                )
                fig.update_yaxes(
                    range=y_range, row=1, col=idx+1, 
                    showline=True, linecolor='lightgrey', mirror=True, 
                    zeroline=True, zerolinecolor='lightgrey', 
                    showgrid=False # ❌ 去掉背景横线
                )
        
            fig.update_layout(
                title=dict(text="<b>各上市公司可转损益负债OCI变动和FVOCI债权公允价值变动情况</b>", font=dict(size=18, color='#00338D')),
                barmode='group', bargap=bar_gap, height=400, # 稍微调高一点
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=80, b=80, l=50, r=20), # 💡 底部 margin 留大一点给图例
                # 💡 图例往下移动 (y 调小)
                legend=dict(orientation="h", yanchor="top", y=-0.28, x=0.5, xanchor="center"),
                annotations=[dict(
                    text=f"单位：{unit_str}人民币", xref="paper", yref="paper", 
                    x=1, y=1.18, showarrow=False, font=dict(size=12, color="#666")
                )]
            )
            return fig
        
        def calculate_oci_analysis_table(df_raw, selected_cos):
            """计算资负OCI变动分析表"""
            fields = ['AC', 'FVOCI', 'FVTPL', '指定FVOCI', 'FVOCI债券公允价值', '可转损益的负债OCI']
            df = df_raw[df_raw['字段名'].isin(fields)].copy()
            df['报告年份'] = df['报告年份'].astype(str).str.replace(".0", "", regex=False)
            
            years = sorted(df['报告年份'].unique(), reverse=True)[:2]
            if len(years) < 2: return pd.DataFrame()
            curr_y, prev_y = years[0], years[1]
        
            results = []
            for co in selected_cos:
                data = {}
                for y in [curr_y, prev_y]:
                    for f in fields:
                        val = df[(df['公司'] == co) & (df['报告年份'] == y) & (df['字段名'] == f)]['(百万)人民币']
                        data[f"{f}_{y}"] = val.iloc[0] if not val.empty else 0
                
                # 1. FVOCI债权占比 (两年)
                def get_ratio(y):
                    denom = data[f"AC_{y}"] + data[f"FVOCI_{y}"] + data[f"FVTPL_{y}"] + data[f"指定FVOCI_{y}"]
                    return data[f"FVOCI_{y}"] / denom if denom != 0 else 0
                
                # 2. 变动率指标
                # FVOCI变动 = 今年/去年 - 1 (或者是纯比例，根据你图表显示，这里用 今年/去年)
                # 提示：你要求的公式是 最新年/去年
                ratio_fvoci = data[f"FVOCI债券公允价值_{curr_y}"] / data[f"FVOCI债券公允价值_{prev_y}"] if data[f"FVOCI债券公允价值_{prev_y}"] != 0 else 0
                ratio_liab = data[f"可转损益的负债OCI_{curr_y}"] / data[f"可转损益的负债OCI_{prev_y}"] if data[f"可转损益的负债OCI_{prev_y}"] != 0 else 0
                ratio_final = ratio_fvoci / ratio_liab if ratio_liab != 0 else 0
                
                results.append({
                    "公司": co,
                    f"FVOCI占比_{prev_y}": get_ratio(prev_y),
                    f"FVOCI占比_{curr_y}": get_ratio(curr_y),
                    "FVOCI变动": ratio_fvoci,
                    "负债OCI变动": ratio_liab,
                    "两年资负OCI变动比率": ratio_final
                })
            
            return pd.DataFrame(results), curr_y, prev_y       
        
        
         # --- 📈 核心绘图逻辑 12：净资产，总资产图 ---   
        def create_asset_trend_chart(df_source, field_name, title, show_labels, p_size, g_gap, current_divisor, current_unit_label):
            import plotly.graph_objects as go
            import numpy as np
            import pandas as pd
        
            # 1. 筛选与单位转换
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            
            # 转换为全局选择的单位
            d['value'] = d['value'] / current_divisor 
            
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            
            # 严格使用你要求的颜色
            color_new = "rgb(0, 51, 141)"   
            color_old = "rgb(15, 101, 253)" 
            
            fig = go.Figure()
            
            # 2. 画柱子
            for yr, col in zip([y_old, y_new], [color_old, color_new]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", 
                    x=df_yr['公司'], 
                    y=df_yr['value'],
                    marker_color=col,
                    # 注意：这里没有 width 参数，靠 bargroupgap=0 贴紧
                    text=[f"{v:.1f}" if not np.isnan(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size),
                    cliponaxis=False
                ))
            
            # 3. 涨幅箭头逻辑 (虽然资产很少为负，但保留镜像逻辑以防万一)
            all_vals = d['value'].dropna()
            max_range = all_vals.max() - all_vals.min() if not all_vals.empty else 100
            off = max_range * 0.12 
            
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color = "#FD349C" if pct >= 0 else "#269924"
                        arr_symbol = "↗" if pct >= 0 else "↘"
                        
                        if v1 >= 0:
                            y_pos = max(v0, v1) + off
                            v_align = "bottom"
                        else:
                            y_pos = min(v0, v1) - off
                            v_align = "top"
                        
                        fig.add_annotation(
                            x=co, y=y_pos,
                            text=f"<b>{arr_symbol} {pct:.1%}</b>",
                            showarrow=False,
                            font=dict(color=arr_color, size=p_size),
                            xshift=10, 
                            valign=v_align
                        )
            
            # 4. 布局修饰
            fig.update_layout(
                title=dict(
                    text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({current_unit_label}人民币)</span>", 
                    font=dict(size=18, color="#00338D"), 
                    x=0.01
                ),
                barmode='group', 
                bargroupgap=0,  # 让旧年和新年柱子严丝合缝
                bargap=g_gap,   # 控制粗细
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=60, l=40, r=40), 
                height=500,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
        
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333', zerolinewidth=1)
            fig.update_xaxes(showline=False, showgrid=False)
            
            return fig
        
        
         # --- 📈 核心绘图逻辑 13：合同服务边际余额图（1/6） ---   
        def create_csm_trend_chart(df_source, field_name, title, show_labels, p_size, g_gap, current_divisor, current_unit_label):
            import plotly.graph_objects as go
            import numpy as np
            import pandas as pd
        
            # 1. 数据准备与单位转换
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / current_divisor 
            
            # 自动获取年份（支持 2023, 2024, 2025 三年展示）
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            
            # 定义颜色：2025:深蓝, 2024:中蓝, 2023:浅蓝
            color_map = {
                years[-1]: "rgb(0, 51, 141)", 
                years[-2]: "rgb(15, 101, 253)"
            }
            # 如果有第三年数据（2023），给一个浅蓝色
            if len(years) > 2:
                color_map[years[-3]] = "rgb(173, 216, 230)"
            
            fig = go.Figure()
            
            # 2. 绘制柱状图
            for yr in years:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", 
                    x=df_yr['公司'], 
                    y=df_yr['value'],
                    marker_color=color_map.get(yr, "gray"),
                    text=[f"{v:.0f}" if not np.isnan(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size),
                    cliponaxis=False
                ))
            
            # 3. 涨幅箭头（计算最新两年的变动）
            y_old, y_new = years[-2], years[-1]
            all_vals = d['value'].dropna()
            max_range = all_vals.max() - all_vals.min() if not all_vals.empty else 100
            off = max_range * 0.12 
            
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color = "#FD349C" if pct >= 0 else "#269924"
                        arr_symbol = "↗" if pct >= 0 else "↘"
                        
                        # 镜像逻辑
                        if v1 >= 0:
                            y_pos = max(v0, v1) + off
                            v_align = "bottom"
                        else:
                            y_pos = min(v0, v1) - off
                            v_align = "top"
                        
                        fig.add_annotation(
                            x=co, y=y_pos,
                            text=f"<b>{arr_symbol} {pct:.1%}</b>",
                            showarrow=False,
                            font=dict(color=arr_color, size=p_size),
                            xshift=15, 
                            valign=v_align
                        )
            
            # 4. 布局修饰
            fig.update_layout(
                title=dict(
                    text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({current_unit_label}人民币)</span>", 
                    font=dict(size=18, color="#00338D"), 
                    x=0.01
                ),
                barmode='group', 
                bargroupgap=0,  # 柱子紧贴
                bargap=g_gap,   # 组间距调节粗细
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=60, l=40, r=40), 
                height=550,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
            )
        
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333', zerolinewidth=1)
            fig.update_xaxes(showline=False, showgrid=False)
            
            return fig



         # --- 📈 核心绘图逻辑 15：合同服务边际余额图（3/6） ---  
        def create_csm_composition_chart(df, selected_cos, target_year, show_labels, label_size, bar_width):
            # 1. 基础配置
            field_map = {
                "新业务CSM（集团口径）": "新业务 CSM",
                "CSM计息": "CSM 计息",
                "CSM调整": "CSM 调整"
            }
            color_map = {
                "新业务CSM（集团口径）": "rgb(0, 51, 140)",
                "CSM计息": "rgb(147, 157, 253)",
                "CSM调整": "rgb(253, 52, 156)"
            }
            fields = list(field_map.keys())
        
            # 2. 数据过滤与处理
            # 筛选指定年份和选定公司
            year_str = str(target_year).replace(".0", "")
            d_sub = df[(df['报告年份'].astype(str).str.contains(year_str)) & (df['公司'].isin(selected_cos))].copy()
            d_sub = d_sub[d_sub['字段名'].isin(fields)]
        
            if d_sub.empty:
                return None
        
            # 透视表：行是公司，列是字段
            d_pivot = d_sub.pivot(index='公司', columns='字段名', values='(百万)人民币')
            # 确保所有列都存在
            for f in fields:
                if f not in d_pivot.columns: d_pivot[f] = 0
            
            # 按照所选公司的顺序排列
            d_pivot = d_pivot.reindex(selected_cos).fillna(0)
        
            # 计算占比逻辑：分母为三项绝对值之和（确保比例反映贡献度，且正负分明）
            # 或者根据业务逻辑，分母为三项代数和。这里采用代数和作为基准，若为负则向下延伸
            d_pivot['Total'] = d_pivot[fields].abs().sum(axis=1).replace(0, 1)
        
            fig = go.Figure()
        
            # 3. 绘制柱状图
            for f_name in fields:
                # 计算百分比
                vals_pct = (d_pivot[f_name] / d_pivot['Total']) * 100
                
                # 标签颜色：深蓝背景用白字，其他用黑字或白字
                t_color = "white" if f_name == "新业务CSM（集团口径）" else "black"
                
                fig.add_trace(go.Bar(
                    name=field_map[f_name],
                    x=d_pivot.index,
                    y=vals_pct,
                    width=bar_width,
                    marker_color=color_map[f_name],
                    text=[f"{v:.0f}%" if abs(v) >= 1 else "" for v in vals_pct] if show_labels else None,
                    textposition='inside',
                    insidetextanchor='middle',
                    textfont=dict(size=label_size, color=t_color),
                    # 💡 关键：禁止标签随高度缩小
                    constraintext='none'
                ))
        
            # 4. 样式美化
            fig.update_layout(
                title=dict(
                    text=f"<b>当期摊销前CSM变动项占比 - {year_str}年</b>",
                    x=0.5, y=0.95, xanchor='center',
                    font=dict(size=20, color='#00338D')
                ),
                # 💡 relative 模式支持正负值从 0 线上下堆叠
                barmode='relative',
                height=550,
                margin=dict(t=80, b=100, l=60, r=40),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)', # 浅淡灰色背景
                legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
                yaxis=dict(
                    showgrid=False,       
                    range=[-25, 105],
                    tickvals=[-20, 0, 20, 40, 60, 80, 100],
                    ticktext=["-20%", "0%", "20%", "40%", "60%", "80%", "100%"], # 2. 这里直接写明要显示的文本
                    zeroline=True,           # 3. 保留 0 线
                    zerolinecolor='GREY', # 0 线颜色加深，使其可见
                    zerolinewidth=0.8
                ),
                xaxis=dict(showgrid=False)
            )
        
        
            return fig         
        
         # --- 📈 核心绘图逻辑 16：合同服务边际余额图 （4/6）---   
        # 🎨 预设颜色库（RGB转HEX）
        PRESET_COLORS = [
            "#C00000", # (192, 0, 0)
            "#0865EE", # (8, 101, 238)
            "#FEAED7", # (254, 174, 215)
            "#92D050", # (146, 208, 80)
            "#7030A0", # (112, 48, 160)
            "#EF9867", # (239, 152, 103)
            "#61CBF4", # (97, 203, 244)
            "#C7A0F7",  # (199, 160, 247)
            '#098E7E',
            '#AB0D82'
        ]
        
        def get_company_color_mapping(companies):
            """极其紧凑的公司颜色选择器"""
            st.write("### 🎨 颜色配置")
            
            # 修复后的颜色列表
            PRESET_COLORS_FIXED = ["#C00000", "#0865EE", "#FEAED7", "#92D050", "#7030A0", "#EF9867", "#61CBF4", "#C7A0F7"]
            
            # 动态创建列，列数等于公司数，让它们排成一行
            n_cos = len(companies)
            cols = st.columns(n_cos)
            
            mapping = {}
            for i, co in enumerate(companies):
                default_color = PRESET_COLORS_FIXED[i % len(PRESET_COLORS_FIXED)]
                with cols[i]:
                    # 使用 collapsed 隐藏标签，让界面非常紧凑
                    st.caption(f"**{co}**") # 仅用小字显示公司名
                    picked_color = st.color_picker(
                        label=co, 
                        value=default_color, 
                        key=f"cp_compact_{co}",
                        label_visibility="collapsed" 
                    )
                    mapping[co] = picked_color
            return mapping
        def create_ratio_line_chart_v3(df_plot, title, color_map, show_labels, marker_size, legend_font_size, y_axis_format=".1%"):
            import plotly.graph_objects as go
            
            fig = go.Figure()
            years = sorted(df_plot['报告年份'].unique())
            if not years: return fig
            latest_year = years[-1]
            
            # 计算最新一年的极值以确定实虚线
            latest_data = df_plot[df_plot['报告年份'] == latest_year].set_index('公司')['value'].dropna()
            max_co = latest_data.idxmax() if not latest_data.empty else None
            min_co = latest_data.idxmin() if not latest_data.empty else None
            
            for co in selected_cos:
                d_co = df_plot[df_plot['公司'] == co].sort_values('报告年份')
                if d_co.empty: continue
                
                is_highlight = (co == max_co or co == min_co)
                line_style = "solid" if is_highlight else "dot"
                
                # 确定显示模式：是否包含文字
                mode = 'lines+markers+text' if show_labels else 'lines+markers'
                
                fig.add_trace(go.Scatter(
                    x=d_co['报告年份'],
                    y=d_co['value'],
                    name=co,
                    mode=mode,
                    line=dict(color=color_map.get(co, "#333"), width=3 if is_highlight else 2, dash=line_style),
                    marker=dict(size=marker_size, symbol='circle'), # 💡 这里的圆形大小受按钮控制
                    text=[f"{v*100:.1f}%" if (is_highlight or i==0 or i==len(d_co)-1) else "" 
                          for i, v in enumerate(d_co['value'])] if show_labels else None,
                    textposition="top center",
                    cliponaxis=False
                ))
                
            fig.update_layout(
                title=dict(text=f"<b>{title}</b>", font=dict(size=18, color="#00338D"), x=0.01),
                paper_bgcolor='rgb(237, 243, 251)',
                plot_bgcolor='rgb(237, 243, 251)',
                margin=dict(t=150, b=40, l=60, r=40), # 💡 增加顶部边距给多行图例留空间
                height=600,
                
                # 💡 精确控制图例排版
                showlegend=True,
                legend=dict(
                    orientation="h",      # 水平排列
                    yanchor="top",
                    y=1.15,               # 放在标题下方
                    xanchor="right",
                    x=1,                  # 整体靠右
                    font=dict(size=legend_font_size),
                    entrywidth=110,       # 💡 关键：每个图例条目的固定宽度，设好后能实现每行4个
                    entrywidthmode="pixels",
                    traceorder="normal",
                    valign="middle"
                ),
                
                # 去掉背景白色网格线
                xaxis=dict(showgrid=False, showline=True, linecolor='#333'),
                yaxis=dict(
                    showgrid=False, 
                    zeroline=True, 
                    zerolinecolor='rgba(0,0,0,0.1)',
                    tickformat=y_axis_format
                )
            )
            return fig




         # --- 📈 核心绘图逻辑 ：新业务盈利合同 ---   
        def create_new_biz_csm_chart(df_source, field_name, title, show_labels, p_size, g_gap, current_divisor, current_unit_label):
            import plotly.graph_objects as go
            import numpy as np
            import pandas as pd
        
            # 1. 数据准备与单位转换
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / current_divisor 
            
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            
            y_old, y_new = years[-2], years[-1]
            
            # 严格保持统一的色彩标准
            color_old = "rgb(15, 101, 253)"
            color_new = "rgb(0, 51, 141)" 
            
            fig = go.Figure()
            
            # 2. 绘制柱状图
            for yr, col in zip([y_old, y_new], [color_old, color_new]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    # 💡 还原图例细节：将年份拼接入图例名称中
                    name=f"{yr}年新业务盈利合同（CSM）", 
                    x=df_yr['公司'], 
                    y=df_yr['value'],
                    marker_color=col,
                    text=[f"{v:.1f}" if not np.isnan(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside',
                    textfont=dict(size=font_size),
                    cliponaxis=False
                ))
            
            # 3. 涨幅箭头与镜像逻辑
            all_vals = d['value'].dropna()
            max_range = all_vals.max() - all_vals.min() if not all_vals.empty else 100
            off = max_range * 0.12 
            
            df_old = d[d['报告年份'] == y_old].set_index('公司')
            df_new = d[d['报告年份'] == y_new].set_index('公司')
            
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color = "#FD349C" if pct >= 0 else "#269924"
                        arr_symbol = "↗" if pct >= 0 else "↘"
                        
                        if v1 >= 0:
                            y_pos = max(v0, v1) + off
                            v_align = "bottom"
                        else:
                            y_pos = min(v0, v1) - off
                            v_align = "top"
                        
                        fig.add_annotation(
                            x=co, y=y_pos,
                            text=f"<b>{arr_symbol} {pct:.1%}</b>",
                            showarrow=False,
                            font=dict(color=arr_color, size=p_size),
                            xshift=10, 
                            valign=v_align
                        )
            
            # 4. 布局修饰
            fig.update_layout(
                title=dict(
                    # 动态生成标题前缀（例如：2025年和2024年保险上市集团...）
                    text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({current_unit_label}人民币)</span>", 
                    font=dict(size=18, color="#00338D"), 
                    x=0.01
                ),
                barmode='group', 
                bargroupgap=0,  # 柱子紧贴
                bargap=g_gap,   # 宽度控制
                paper_bgcolor='rgba(0,0,0,0)', 
                plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(t=120, b=60, l=40, r=40), 
                height=500,
                # 图例放在右上角以匹配原图
                legend=dict(orientation="v", yanchor="top", y=1.1, xanchor="right", x=1.0)
            )
        
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='#333333', zerolinewidth=1)
            fig.update_xaxes(showline=False, showgrid=False)
            
            return fig








        # --- 🖼️ 第一页PPT调用按钮设置：顺序显示普通柱状图 (Tasks 循环) ---
        tasks = [
            ("保险服务收入合计", "保险服务收入"),
            ("保险服务费用合计", "保险服务费用"),
            ("保险服务业绩", "保险服务业绩")
        ]
    
        figs_to_ppt = []
        global_gap_val = 0.3 # 给一个默认值
        
        for field, title in tasks:
            st.write(f"### {title}")
            
            # 如果是第一张图，渲染三个控件（加上全局柱宽/间隙控制）
            if field == "保险服务收入合计":
                ctrl1, ctrl2, ctrl3 = st.columns([1.5, 2, 3])
                with ctrl1:
                    label_on = st.toggle(f"显示数据标签", value=True, key=f"lab_{field}")
                with ctrl2:
                    p_size = st.slider("涨幅文字大小", 8, 24, 14, key=f"psize_{field}")
                with ctrl3:
                    # 这个滑块控制间隙，间隙越小柱子越粗，间隙越大柱子越细
                    global_gap_val = st.slider(
                        "调整公司间距与柱子粗细 (数值越小柱子越粗)", 
                        min_value=0.1, max_value=0.8, value=0.3, step=0.05, key="global_gap"
                    )
            # 如果是后两张图，只渲染它自己的两个控件
            else:
                ctrl1, ctrl2 = st.columns([1.5, 5])
                with ctrl1:
                    label_on = st.toggle(f"显示数据标签", value=True, key=f"lab_{field}")
                with ctrl2:
                    p_size = st.slider("涨幅文字大小", 8, 24, 14, key=f"psize_{field}")
            
            # 把 global_gap_val 传给所有的图表
            fig = create_kpmg_chart(df_raw, field, title, label_on, p_size, global_gap_val)
            st.plotly_chart(fig, use_container_width=True)
            figs_to_ppt.append((title, fig))

        # --- 🖼️  第2页PPT调用按钮设置：独立显示构成比例图---
        st.markdown("---")
        comp_title = "保险业务收入构成（1/2）"
        comp_fields = [
            "采用保费分配法计量的保险合同保险服务收入", 
            "未采用保费分配法计量的保险合同保险服务收入"
        ]
        
        st.write(f"### {comp_title}")
        
        # 💡 创建四列控件
        c_ctrl1, c_ctrl2, c_ctrl3, c_ctrl4 = st.columns([1, 1.5, 1.5, 1.5])
        with c_ctrl1:
            label_on_comp = st.toggle("显示数据标签", value=True, key="lab_comp")
        with c_ctrl2:
            l_size = st.slider("标签大小", 5, 20, 12, key="l_size_comp")
        with c_ctrl3:
            b_width = st.slider("柱子宽度", 0.1, 1.0, 0.6, 0.05, key="b_width_comp")
        with c_ctrl4:
            c_f_size = st.slider("公司名称大小", 10, 20, 14, key="c_f_size_comp")
        
        # 调用函数，传入新增的调节变量
        fig_comp = create_kpmg_composition_chart(
            df_raw, comp_fields, comp_title, label_on_comp, 
            label_size=l_size, 
            bar_width=b_width, 
            co_font_size=c_f_size
        )
        
        st.plotly_chart(fig_comp, use_container_width=True)
        figs_to_ppt.append((comp_title, fig_comp))

        # ──────── 🖼️  第3页PPT调用按钮设置：结构化分析图表 2/2 调用区 ───────────
        st.markdown("---")
        title_2 = "保险业务收入构成（2/2）"
        
        # 1. 定义原始名称到展示名称的映射 (顺序从下往上)
        field_map_2 = {
            "合同服务边际的摊销": "合同服务边际的释放",
            "非金融风险调整的变动": "非金融风险调整的变动",
            "预计当期发生的保险服务费用": "预期当期发生的保险服务费用",
            "保险获取现金流的摊销（保险服务收入）": "保险获取现金流的摊销",
            "与当期服务或过去服务相关得保费经验调整": "与当期服务或过去服务相关的保费经验调整",
            "其他收入调整": "其他"
        }
        
        # 2. 定义颜色映射 (RGB)
        color_map_2 = {
            "合同服务边际的摊销": "rgb(30, 73, 226)",
            "非金融风险调整的变动": "rgb(254, 174, 215)",
            "预计当期发生的保险服务费用": "rgb(0, 163, 161)",
            "保险获取现金流的摊销（保险服务收入）": "rgb(1, 184, 245)",
            "与当期服务或过去服务相关得保费经验调整": "rgb(0, 219, 214)",
            "其他收入调整": "rgb(114, 19, 234)"
        }

        st.write(f"### {title_2}")
        
        # 控件区
        c2_ctrl1, c2_ctrl2, c2_ctrl3, c2_ctrl4 = st.columns([1, 1.5, 1.5, 1.5])
        with c2_ctrl1:
            lab_2 = st.toggle("显示数据标签", value=True, key="lab_2")
        with c2_ctrl2:
            l_sz_2 = st.slider("标签大小", 8, 20, 10, key="l_sz_2")
        with c2_ctrl3:
            b_wd_2 = st.slider("柱子宽度", 0.1, 1.0, 0.7, 0.05, key="b_wd_2")
        with c2_ctrl4:
            c_sz_2 = st.slider("公司名称大小", 10, 20, 14, key="c_sz_2")
            
        # 调用函数
        fig_multi = create_kpmg_multi_composition_chart(
            df_raw, field_map_2, color_map_2, title_2, lab_2, 
            label_size=l_sz_2, bar_width=b_wd_2, co_font_size=c_sz_2
        )
        
        st.plotly_chart(fig_multi, use_container_width=True)
        figs_to_ppt.append((title_2, fig_multi))         
        
        # ─────────── 🖼️  第4页PPT调用按钮设置：利润贡献分析 ───────────
        st.markdown("---")
        st.write("利润贡献-保险利润")
    
        # 控制面板
        c1, c2, c3, c4 = st.columns(4)
        with c1: ui_prof_lab = st.toggle("显示数据标签", value=False, key="prof_lab")
        with c2: ui_prof_sz = st.slider("标签字号", 8, 16, 11, key="prof_sz")
        with c3: ui_prof_bw = st.slider("柱子宽度", 0.2, 0.8, 0.4, key="prof_bw")
        with c4: ui_prof_co = st.slider("公司名字号", 10, 20, 14, key="prof_co")
    
        # 渲染 2025年
        st.write("### 📅 2025年度")
        fig_prof_2025 = create_profit_composition_chart(df_raw, selected_cos, 2025, ui_prof_lab, ui_prof_sz, ui_prof_bw, ui_prof_co)
        if fig_prof_2025:
            st.plotly_chart(fig_prof_2025, use_container_width=True)
        else:
            st.info("暂无 2025 年利润构成数据")
    
        st.write("---")
    
        # 渲染 2024年
        st.write("### 📅 2024年度")
        fig_prof_2024 = create_profit_composition_chart(df_raw, selected_cos, 2024, ui_prof_lab, ui_prof_sz, ui_prof_bw, ui_prof_co)
        if fig_prof_2024:
            st.plotly_chart(fig_prof_2024, use_container_width=True)
        else:
            st.info("暂无 2024 年利润构成数据")
            
        # --- 🖼️ 第5页PPT调用按钮设置 ---
        display_tasks = [
            ("净投资回报", "上市公司净投资回报变动趋势", "直接提取原始项"),
            ("承保财务净损益", "上市公司承保财务净损益变动趋势", "-(承保财务损益) - 分出再保险财务损益"),
            ("投资利润", "上市公司投资利润变动趋势", "净投资回报 - 承保财务损益")
        ]

        global_gap_val = 0.3
        for field, title, formula in display_tasks:
            st.write(f"### 📊 {field}")
            
            # 1. 控件布局
            c1, c2, c3 = st.columns([2, 2, 3])
            with c1: lab = st.toggle("显示标签", True, key=f"l_{field}")
            with c2: psz = st.slider("文字大小", 8, 20, 12, key=f"s_{field}")
            if field == "净投资回报":
                with c3: global_gap_val = st.slider("调整柱子粗细", 0.1, 0.8, 0.3)

            # 2. 绘图
            fig = create_simple_kpmg_chart(df_plot_final, field, title, lab, psz, global_gap_val)
            st.plotly_chart(fig, use_container_width=True)

            # 3.  核心预览区：仅展示公式、原始数据项目、结果
            with st.expander(f"🔍 查看【{field}】计算逻辑与明细"):
                st.info(f"**计算公式：** {formula}")
                
                # 确定预览要显示的列
                if field == "净投资回报":
                    show_cols = ['公司', '报告年份', '净投资回报']
                elif field == "承保财务净损益":
                    show_cols = ['公司', '报告年份', '承保财务损益', '分出再保险财务损益', '承保财务净损益']
                else:
                    show_cols = ['公司', '报告年份', '净投资回报', '承保财务损益', '投资利润']
                
                # 只展示选中的公司和相关的列
                preview_df = df_pivot[df_pivot['公司'].isin(selected_cos)][show_cols].copy()
                st.dataframe(preview_df.style.format(precision=2), use_container_width=True)
            
        # --- 🖼️ 第6页PPT:保险利润与投资利润对比 ---
        st.markdown("---")
        st.write("## 关键年报数据 - 保险利润、投资利润变动趋势")

        # 动态识别年份（假设为 2023 和 2024）
        available_years = sorted(df_profit_raw['报告年份'].astype(str).str.replace(".0", "", regex=False).unique())
        
        if len(available_years) >= 2:
            yr_old, yr_new = available_years[-2], available_years[-1]

            # 定义任务配置（含专属配色）
            profit_tasks = [
                {
                    "field": "投资利润",
                    "title": "上市公司投资利润变动趋势",
                    "colors": {
                        yr_old: "#FC349B", # 旧年: RGB(252, 52, 155)
                        yr_new: "#97014F"  # 新年: RGB(151, 1, 79)
                    }
                },
                {
                    "field": "保险利润",
                    "title": "上市公司保险利润变动趋势",
                    "colors": {
                        yr_old: "#1E49E2", # 旧年: RGB(30, 73, 226)
                        yr_new: "#00338C"  # 新年: RGB(0, 51, 140)
                    }
                }
            ]

            global_gap_val = 0.3
            for task in profit_tasks:
                field = task["field"]
                st.write(f"### 📊 {field}")
                
                # 控件
                c1, c2, c3 = st.columns([2, 2, 3])
                with c1: lab = st.toggle("显示标签", True, key=f"profit_lab_{field}")
                with c2: psz = st.slider("文字大小", 8, 20, 12, key=f"profit_sz_{field}")
                if field == "投资利润":
                    with c3: global_gap_val = st.slider("调整柱子粗细", 0.1, 0.8, 0.3, key="gap_profit_page")

                # 绘图调用
                fig = create_profit_chart_v4(
                    df_profit_raw, field, task["title"], task["colors"], lab, psz, global_gap_val
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # 数据预览（仅展示该项目的数据，方便核对）
                with st.expander(f"🔍 查看【{field}】数据详情"):
                    # 这里的预览展示原始单位数值（百万）
                    view_df = df_profit_raw[df_profit_raw['指标名称'] == field].pivot(
                        index='公司', columns='报告年份', values='value'
                    ).round(2)
                    st.dataframe(view_df, use_container_width=True)
                
                figs_to_ppt.append((task["title"], fig))
        else:
            st.warning("数据年份不足，无法生成对比图。")            
            
        # --- 🖼️ 第7页PPT:税和利润相关按钮 ---
        st.write("### 税前利润与净利润趋势")
        c1, c2, c3, c4, c5 = st.columns(5) # 5列分布
        with c1: ui_show_lab = st.toggle("显示标签", value=True)
        with c2: ui_bar_w = st.slider("柱子粗细", 0.1, 0.8, 0.4)
        with c3: ui_fs = st.slider("数据字体", 8, 16, 10)
        with c4: ui_cfs = st.slider("公司名字体", 10, 24, 14)
        with c5: ui_co_y = st.slider("公司名高度", 1.0, 1.2, 1.05, step=0.01) # 新滑块

        # 调用函数 (现在是7个参数)
        fig_tax = create_tax_subplot_chart(
            df_tax_pivot, 
            selected_cos, 
            ui_show_lab, 
            ui_bar_w, 
            ui_fs, 
            ui_cfs,
            ui_co_y # 传入新参数
        )

        st.plotly_chart(fig_tax, use_container_width=True)
            
        # --- 🖼️ 第8页PPT:利润与综合收益相关按钮 ---
        st.markdown("---")
        st.write("## 综合收益变动趋势分析")
        
        # 💡 这里删掉了 ui_bar_w 的列，因为现在由 Gap 统一控制粗细，这样更整齐
        col_c1, col_c2, col_c3 = st.columns([1, 2, 2]) 
        with col_c1:
            ui_show_lab = st.toggle("显示数值", value=True, key="lab_fin_v5")
        with col_c2:
            # 💡 这个滑块现在是“全能”的：数值越小，柱子越粗且挨得越紧
            ui_g_gap = st.slider("调整公司间距与柱子粗细", 0.1, 0.8, 0.3, key="gap_fin_v5")
        with col_c3:
            ui_p_size = st.slider("涨幅字体大小", 8, 24, 11, key="psize_fin_v5")

        # --- 渲染三个图表 ---
        titles = {
            '净利润': '净利润 (亏损) 变动趋势',
            '其他综合收益': '其他综合收益 (OCI) 变动趋势',
            '综合收益总额': '综合收益 (亏损) 总额变动趋势'
        }

        # 确保 target_metrics 顺序正确
        target_metrics_list = ['净利润', '其他综合收益', '综合收益总额']

        for metric in target_metrics_list:
            # 💡 这里的参数必须和我们定义的 create_financial_trend_chart_v5 严格对应
            fig_fin = create_financial_trend_chart_v5(
                df_fin_raw, 
                metric, 
                titles[metric], 
                ui_show_lab, 
                ui_p_size, 
                ui_g_gap,
                divisor,     # 💡 传入全局单位除数
                unit_label   # 💡 传入全局单位文本
            )
            st.plotly_chart(fig_fin, use_container_width=True)
            # 如果你有导出PPT的需求，记得加入列表
            figs_to_ppt.append((titles[metric], fig_fin))
 
        
 
    
       # --- 🖼️ 第9页: I9资产端 --- 
        field_map = {
            "AC": "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;债权投资<br>(以摊余成本法计量)",
            "FVOCI": "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;其他债权投资<br>(以公允价值计量入其他综合收益)",
            "FVTPL": "&nbsp;&nbsp;&nbsp;交易性金融资产<br>(以公允价值计量入损益)",
            "指定FVOCI": "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;其他权益工具投资<br>(以公允价值计量入其他综合收益)"
        }
        
        color_map = {
            "AC": "rgb(0, 184, 245)",
            "FVOCI": "rgb(114, 19, 234)",
            "FVTPL": "rgb(253, 52, 156)",
            "指定FVOCI": "rgb(181, 2, 95)"
        }
        
        st.write("资产配置结构分析")
        c1, c2, c3, c4 = st.columns(4)
        with c1: show_lab = st.toggle("显示数据标签", value=True)
        with c2: lab_sz = st.slider("标签字号", 8, 16, 11)
        with c3: bar_w = st.slider("柱子宽度", 0.2, 0.9, 0.6, 0.1)
        with c4: co_fs = st.slider("公司名大小", 10, 20, 14)
        
        fig = create_asset_composition_chart(
            df_raw, selected_cos, field_map, color_map, 
            "各公司资产配置结构变动", show_lab, lab_sz, bar_w, co_fs
        )
        st.plotly_chart(fig, use_container_width=True)
        
        
        
        # --- 🖼️ 第10页PPT调用按钮设置：顺序显示综合收益图 (Tasks 循环) ---
        st.markdown("---")
        st.write("##综合收益变动情况")

        # 💡 1. 抓取全局单位 (这里替换成你 Step7 控制单位的真实变量)
        # 例如：如果你的全局按钮选择存放在 st.session_state['unit']
        unit_divisor = st.session_state.get('unit_divisor', 1000)
        unit_str = st.session_state.get('unit_name', '十亿元')

        # 💡 2. 动态获取源数据中可用的年份 (倒序排序：先展示最新年，后展示上一年)
        all_years = sorted(df_raw['报告年份'].dropna().unique(), reverse=True)
        # 去掉年份中的 .0 并限制只取最近两年
        clean_years = [str(y).replace(".0", "") for y in all_years][:2] 

        # 构造 tasks 列表 -> 格式: (传给函数的年份, 图表标题)
        oci_tasks = [
            (y, f"{y} 各公司综合收益变动情况") for y in clean_years
        ]
        
        # 给控制滑块默认值
        global_oci_gap = 0.15
        global_oci_font = 13

        for i, (year_field, title) in enumerate(oci_tasks):
            st.write(f"### {title}")
            
            # 如果是第一张图 (最新的一年)，渲染所有完整控件
            if i == 0:
                c1, c2, c3 = st.columns([1, 2, 2])
                with c1:
                    label_on = st.toggle("显示数据标签", value=True, key=f"lab_oci_{year_field}")
                with c2:
                    global_oci_gap = st.slider(
                        "调节柱子间距与粗细", min_value=0.0, max_value=0.6, value=0.15, step=0.05, key="oci_gap"
                    )
                with c3:
                    global_oci_font = st.slider(
                        "公司名称字体大小", 8, 20, 13, 1, key="oci_font"
                    )
            # 如果是第二张图 (去年)，只渲染标签控制，间隙和大小沿用第一张图的设置
            else:
                c1, c2 = st.columns([1, 4])
                with c1:
                    label_on = st.toggle("显示数据标签", value=True, key=f"lab_oci_{year_field}")

            # 调用封装好的画图函数，并传入所需的全局变量
            fig = create_oci_chart(
                df_raw, 
                year=year_field, 
                title=title, 
                show_labels=label_on, 
                co_font_size=global_oci_font, 
                bar_gap=global_oci_gap, 
                unit_divisor=divisor,   # 联动单位
                unit_str=unit_label,          # 联动单位名称
                selected_cos=selected_cos
            )
            
            st.plotly_chart(fig, use_container_width=True)
            figs_to_ppt.append((title, fig)) # 如果你有存入PPT的列表需求，加这句
        # ========================================================= 
        
     
      
       # --- 🖼️ 第11页: 资负项变动分析 (表+图) ---

        st.markdown("---")
        st.write("资负项 OCI 变动分析")

        # 💡 1. 单位联动逻辑：直接从 Step 7 的全局设置中获取
        # 请确保 st.session_state['unit_divisor'] 和 'unit_name' 与你顶部按钮一致
        u_divisor = st.session_state.get('unit_divisor', 1000)
        u_name = st.session_state.get('unit_name', '十亿元')

        # 2. ⚙️ 控制面板
        st.write("---")
        col_c1, col_c2, col_c3 = st.columns([1, 1.5, 1.5])
        with col_c1:
            ui_show_label_deep = st.toggle("显示数据标签", value=True, key="lab_deep_oci")
        with col_c2:
            ui_bar_gap_deep = st.slider("调节柱子间距", 0.1, 0.8, 0.3, key="gap_deep_oci")
        with col_c3:
            ui_font_deep = st.slider("公司名字字体大小", 8, 20, 12, key="font_deep_oci")

        # 3. 计算表格 (计算逻辑见上一个回答中的 calculate_oci_analysis_table)
        df_analysis, c_y, p_y = calculate_oci_analysis_table(df_raw, selected_cos)

        if not df_analysis.empty:
            # 展示表格 (支持直接复制)
            st.write(f"资负 OCI 分析指标表 ({p_y}-{c_y})")
            
            display_df = pd.DataFrame({"项目": [
                f"FVOCI债权占比总金融资产 ({p_y})",
                f"FVOCI债权占比总金融资产 ({c_y})",
                f"FVOCI变动 - {c_y}/{p_y} [1]",
                f"负债OCI变动 - {c_y}/{p_y} [2]",
                f"两年资负OCI变动比率 [1]/[2]"
            ]})
            
            for _, row in df_analysis.iterrows():
                display_df[row['公司']] = [
                    f"{row[f'FVOCI占比_{p_y}']*100:.1f}%",
                    f"{row[f'FVOCI占比_{c_y}']*100:.1f}%",
                    f"{row['FVOCI变动']*100:.0f}%",
                    f"{row['负债OCI变动']*100:.0f}%",
                    f"{row['两年资负OCI变动比率']*100:.0f}%"
                ]
            
            st.dataframe(display_df, hide_index=True, use_container_width=True)
            
            # 4. 渲染图表
            st.write("指标变动对比图")
            fig_deep = create_asset_liab_oci_chart(
                df_raw, selected_cos, 
                unit_divisor=divisor,  # 传入step7的数值
                unit_str=unit_label,       # 传入step7的数值
                bar_gap=ui_bar_gap_deep, 
                co_font_size=ui_font_deep, 
                show_labels=ui_show_label_deep
            )
            st.plotly_chart(fig_deep, use_container_width=True)
            
            # 存入 PPT 导出列表
            figs_to_ppt.append(("资负深度分析图", fig_deep))
        else:
            st.warning("数据源中缺少计算所需的 AC/FVOCI/FVTPL/负债OCI 等科目。")
        
     
        
     
        
     
        
     
        
     
        
     
        
     
        
     
        
     
        
     
        
        # --- 🖼️ 第12页PPT:净资产与总资产变动趋势 ---
        st.markdown("---")
        st.write("## 净资产与总资产变动趋势")
        
        # 1. 数据预处理
        asset_fields = ['期末股东权益', '总资产']
        df_asset_raw = df_raw[
            (df_raw['字段名'].isin(asset_fields)) & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        
        # 统一列名以适配绘图函数
        df_asset_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)

        # 2. 独立UI控制按钮（加上专属 key 防止冲突）
        col_a1, col_a2, col_a3 = st.columns([1, 2, 2]) 
        with col_a1:
            ui_show_lab_asset = st.toggle("显示数值", value=True, key="lab_asset")
        with col_a2:
            ui_g_gap_asset = st.slider("调整公司间距与柱子粗细", 0.1, 0.8, 0.3, key="gap_asset")
        with col_a3:
            ui_p_size_asset = st.slider("涨幅字体大小", 8, 24, 11, key="psize_asset")

        # 3. 核心：字段名 -> 展示图表名的映射
        asset_tasks = [
            ('期末股东权益', '净资产变动趋势'), 
            ('总资产', '总资产变动趋势')
        ]

        # 4. 循环出图
        for field, title in asset_tasks:
            fig_asset = create_asset_trend_chart(
                df_asset_raw, 
                field,       # 传入底层真实的字段名（期末股东权益/总资产）
                title,       # 传入想要展示的标题（净资产.../总资产...）
                ui_show_lab_asset, 
                ui_p_size_asset, 
                ui_g_gap_asset,
                divisor,     # 响应顶部的全局单位倍数
                unit_label   # 响应顶部的全局单位文本
            )
            st.plotly_chart(fig_asset, use_container_width=True)
            # 添加到 PPT 导出队列
            figs_to_ppt.append((title, fig_asset))      
            
            
            
        # --- 🖼️ 第13页PPT: 合同服务边际 (CSM) （1/7）---
        st.markdown("---")
        st.write("## 合同服务边际 (CSM) 趋势分析")
        
        # 1. 数据准备
        df_csm_raw = df_raw[
            (df_raw['字段名'] == 'CSM期末余额') & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        
        df_csm_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)

        # 2. UI 控制
        col_s1, col_s2, col_s3 = st.columns([1, 2, 2]) 
        with col_s1:
            ui_show_lab_csm = st.toggle("显示数值", value=True, key="lab_csm")
        with col_s2:
            ui_g_gap_csm = st.slider("调整公司间距与柱子粗细", 0.1, 0.8, 0.35, key="gap_csm")
        with col_s3:
            ui_p_size_csm = st.slider("涨幅字体大小", 8, 24, 12, key="psize_csm")

        # 3. 调用绘图
        fig_csm = create_csm_trend_chart(
            df_csm_raw, 
            'CSM期末余额', 
            '合同服务边际余额趋势变化', 
            ui_show_lab_csm, 
            ui_p_size_csm, 
            ui_g_gap_csm,
            divisor,      # 响应全局单位
            unit_label    # 响应全局文本
        )
        
        st.plotly_chart(fig_csm, use_container_width=True)
        
        # 导出列表
        figs_to_ppt.append(('CSM余额变动趋势', fig_csm))         
        
   
        # --- 🖼️ 第14页PPT: 合同服务边际 (CSM) （2/7）---       

        # --- 🖼️ 第15页PPT: 合同服务边际 (CSM) （3/7）---
        st.markdown("---")
        st.write("摊销前CSM变动分析")
    
        # 控制面板
        c1, c2, c3 = st.columns(3)
        with c1: ui_csm_lab = st.toggle("显示数据标签", value=True, key="csm_lab")
        with c2: ui_csm_sz = st.slider("标签字号", 8, 16, 11, key="csm_sz")
        with c3: ui_csm_bw = st.slider("柱子宽度", 0.2, 0.8, 0.5, key="csm_bw")
    
        # 渲染 2025年
        st.write("2025年度分析")
        fig_2025 = create_csm_composition_chart(df_raw, selected_cos, 2025, ui_csm_lab, ui_csm_sz, ui_csm_bw)
        if fig_2025:
            st.plotly_chart(fig_2025, use_container_width=True)
            # figs_to_ppt.append(("CSM占比_2025", fig_2025))
        else:
            st.info("暂无 2025 年 CSM 数据")
    
        st.write("---")
    
        # 渲染 2024年
        st.write("2024年度分析")
        fig_2024 = create_csm_composition_chart(df_raw, selected_cos, 2024, ui_csm_lab, ui_csm_sz, ui_csm_bw)
        if fig_2024:
            st.plotly_chart(fig_2024, use_container_width=True)
            # figs_to_ppt.append(("CSM占比_2024", fig_2024))
        else:
            st.info("暂无 2024 年 CSM 数据")

        # --- 🖼️ 第16页PPT: CSM 效率与持续率 ---
        st.markdown("---")
        st.write("## CSM摊销比率与持续率趋势分析")

        # 1. 紧凑型颜色配置 (一行显示)
        color_mapping = get_company_color_mapping(selected_cos)
        
        # 2. 🎮 新增：多维度调节按钮区
        st.write("### ⚙️ 图表参数微调")
        c1, c2, c3, c4 = st.columns([1, 1, 1.5, 1.5])
        with c1:
            ui_show_labels_13 = st.toggle("显示数值", value=True, key="lab_13")
        with c2:
            ui_marker_size_13 = st.slider("节点圆点大小", 4, 15, 8, key="mk_13")
        with c3:
            ui_legend_size_13 = st.slider("图例文字大小", 8, 16, 11, key="lg_13")
        with c4:
            # 也可以在这里调 entrywidth，如果名字长短不一
            ui_entry_w_13 = st.slider("图例间距宽度", 80, 200, 115, key="ew_13")

        # 3. 数据处理逻辑
        target_fields = ['CSM摊销', 'CSM期末余额', '新业务CSM（集团口径）']
        df_sub = df_raw[df_raw['字段名'].isin(target_fields)].copy()
        df_pivot = df_sub.pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').reset_index().fillna(0)
        
        df_pivot['摊销比率'] = -df_pivot['CSM摊销'] / (df_pivot['CSM期末余额'] - df_pivot['CSM摊销'])
        df_pivot['持续率'] = -df_pivot['新业务CSM（集团口径）'] / df_pivot['CSM摊销']
        df_pivot.replace([np.inf, -np.inf], np.nan, inplace=True)
        df_pivot['报告年份'] = df_pivot['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"

        # 4. 渲染两张图
        col_l, col_r = st.columns(2)
        
        with col_l:
            df_p1 = df_pivot[['公司', '报告年份', '摊销比率']].rename(columns={'摊销比率': 'value'})
            fig1 = create_ratio_line_chart_v3(
                df_p1, "CSM摊销比率趋势", color_mapping, 
                ui_show_labels_13, ui_marker_size_13, ui_legend_size_13
            )
            # 动态更新 entrywidth
            fig1.update_layout(legend=dict(entrywidth=ui_entry_w_13))
            st.plotly_chart(fig1, use_container_width=True)
            
        with col_r:
            df_p2 = df_pivot[['公司', '报告年份', '持续率']].rename(columns={'持续率': 'value'})
            fig2 = create_ratio_line_chart_v3(
                df_p2, "CSM持续率趋势", color_mapping, 
                ui_show_labels_13, ui_marker_size_13, ui_legend_size_13
            )
            fig2.update_layout(legend=dict(entrywidth=ui_entry_w_13))
            st.plotly_chart(fig2, use_container_width=True)






        # --- 🖼️ 第 页PPT: 新业务CSM ---
        st.markdown("---")
        st.write("## 新业务盈利合同 (CSM) 对比")
        
        # 1. 提取并清理数据
        df_nb_csm_raw = df_raw[
            (df_raw['字段名'] == '新业务CSM（集团口径）') & 
            (df_raw['公司'].isin(selected_cos))
        ].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        
        df_nb_csm_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)

        # 2. 专属UI控制按钮 (配置唯一的 ID key)
        col_n1, col_n2, col_n3 = st.columns([1, 2, 2]) 
        with col_n1:
            ui_show_lab_nb = st.toggle("显示数值", value=True, key="lab_nb_csm")
        with col_n2:
            ui_g_gap_nb = st.slider("调整公司间距与柱子粗细", 0.1, 0.8, 0.3, key="gap_nb_csm")
        with col_n3:
            ui_p_size_nb = st.slider("涨幅字体大小", 8, 24, 12, key="psize_nb_csm")

        # 3. 生成并展示图表
        # 这里传入底层名字：新业务CSM（集团口径），以及想要展示的名字：新业务盈利合同（CSM）对比
        fig_nb_csm = create_new_biz_csm_chart(
            df_nb_csm_raw, 
            '新业务CSM（集团口径）', 
            '新业务盈利合同（CSM）对比', 
            ui_show_lab_nb, 
            ui_p_size_nb, 
            ui_g_gap_nb,
            divisor,      # 获取页面顶部的全局单位
            unit_label    # 获取页面顶部的单位标签
        )
        
        st.plotly_chart(fig_nb_csm, use_container_width=True)
        
        # 加入 PPT 导出列表
        figs_to_ppt.append(('新业务盈利合同CSM对比', fig_nb_csm))
            
        # --- 📂 PPT 一键导出区 ---
        st.markdown("---")
        if st.button("🌞 一键将以上图表导出至 PPT", use_container_width=True):
            with st.spinner("正在排版 PPT..."):
                from pptx import Presentation
                from pptx.util import Inches
                prs = Presentation()
                for title_text, fig in figs_to_ppt:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    img_buf = io.BytesIO()
                    fig.write_image(img_buf, format="png", width=1000, height=600)
                    img_buf.seek(0)
                    slide.shapes.add_picture(img_buf, Inches(0.5), Inches(1.2), width=Inches(9))
                
                ppt_buf = io.BytesIO()
                prs.save(ppt_buf)
                st.session_state['pptx_output'] = ppt_buf.getvalue()
                st.success("🎉 PPT 已生成！")

        if 'pptx_output' in st.session_state:
            st.download_button("📥 下载 PPT 报告", st.session_state['pptx_output'], 
                             file_name="保险关键指标分析报告.pptx", use_container_width=True)

# ==================== 页脚 ====================
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 13px; letter-spacing: 1px; margin-top: 50px; padding: 20px; border-top: 1px solid #CBD5E1;">
    Actuarial Data Intelligence · Built for KPMG Actuary Team
</div>
""", unsafe_allow_html=True)
