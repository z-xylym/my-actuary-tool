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
    <div style="line-height: 1.6; font-size: 0.85rem; color: #333333;">
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
                "20年合同服务边际",
                "投资收益率",
                "综合偿付能力充足率"
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

                base_cols = ["公司类型", "公司", "类别", "字段名", "字段类型"]
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
                    
                    final_cols = ["公司类型","公司", "类别", "字段名", "字段类型", "报告年份", "(百万)原币", "汇率", "(百万)人民币"]
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

    # ── 数据源 ──────────────────────────────────────────────────────────
    if 'integrated_data' not in st.session_state:
        st.session_state['integrated_data'] = None

    source_choice = st.radio(
        "数据源选择", ["直接引用集成后的数据", "上传集成表 Excel"], horizontal=True
    )
    df_raw = None

    if source_choice == "上传集成表 Excel":
        viz_file = st.file_uploader("上传行业集成目标表", type=["xlsx"])
        if viz_file:
            df_raw = pd.read_excel(viz_file)
            st.session_state['integrated_data'] = df_raw
    else:
        df_raw = st.session_state.get('integrated_data')

    if df_raw is None:
        st.info("💡 请先完成数据集成或上传目标底稿。")
        st.stop()

    # ── 数据清洗 ─────────────────────────────────────────────────────────
    df_clean = df_raw.copy()
    df_clean.columns = (
        df_clean.columns.astype(str)
        .str.strip()
        .str.replace('\n', '', regex=False)
        .str.replace('\r', '', regex=False)
        .str.replace('\ufeff', '', regex=False)
    )
    df_clean['报告年份'] = df_clean['报告年份'].astype(str).str.replace('.0', '', regex=False)
    val_col = "(百万)人民币" if "(百万)人民币" in df_clean.columns else df_clean.columns[-1]

    # ── 去重 ─────────────────────────────────────────────────────────────
    dedup_cols = ['公司', '报告年份', '字段名']
    if '类别' in df_clean.columns:
        dedup_cols.append('类别')
    df_clean = df_clean.drop_duplicates(subset=dedup_cols, keep='first')

    # ── 透视表 ───────────────────────────────────────────────────────────
    df_pivot = (
        df_clean.groupby(['公司', '报告年份', '字段名'])[val_col]
        .sum().unstack('字段名').reset_index()
    )

    # ── 字段 / 类别 / 公司类型 ───────────────────────────────────────────
    all_fields   = sorted(df_clean['字段名'].unique().tolist())
    all_types    = sorted(df_clean['类别'].dropna().astype(str).unique().tolist())    if '类别'    in df_clean.columns else []
    all_co_types = sorted(df_clean['公司类型'].dropna().astype(str).unique().tolist()) if '公司类型' in df_clean.columns else []

    # ── 动态年份颜色 ─────────────────────────────────────────────────────
    KPMG_COLOR_MAP = {
        "Lightest": "#BFE8FF",     # Pacific Blue (light)
        "Light": "#76D2FF",      # Blue (original Light Blue)
        "Primary": "#1E49E2",     # Cobalt Blue
        "Dark": "#00338D",       # KPMG Blue (original Primary Blue)
    }
    
    # 提取已排序的年份
    all_years_sorted = sorted(
        [y for y in df_clean['报告年份'].unique() if str(y).isdigit()],
        key=lambda x: int(x)
    )
    n_years = len(all_years_sorted)
    dynamic_year_colors = {}
    
    # ── 根据年份数量选择颜色策略 ─────────────────────────
    target_colors = [] # 最终用于映射的颜色列表
    
    if n_years == 3:
        # 3年：按要求，从 Light Blue (#76D2FF) 开始，依次是 Cobalt Blue (#1E49E2)，KPMG Blue (#00338D)
        target_colors = [
            KPMG_COLOR_MAP["Light"],
            KPMG_COLOR_MAP["Primary"],
            KPMG_COLOR_MAP["Dark"]
        ]
    elif n_years == 2:
        # 2年：按要求，前一个年份 (2024) 用 Cobalt Blue (#1E49E2)，后一个年份 (2025) 用 KPMG Blue (#00338D)
        target_colors = [
            KPMG_COLOR_MAP["Primary"],
            KPMG_COLOR_MAP["Dark"]
        ]
    else:
        # 1年 或 4年以上：使用默认的浅到深渐变 (Lightest -> Dark)
        target_colors = [
            KPMG_COLOR_MAP["Lightest"],
            KPMG_COLOR_MAP["Light"],
            KPMG_COLOR_MAP["Primary"],
            KPMG_COLOR_MAP["Dark"]
        ]
    
    # ── 映射年份到颜色 ─────────────────────────
    num_colors_to_use = len(target_colors)
    
    if n_years > 0:
        for idx, year in enumerate(all_years_sorted):
            color_index = 0 # 默认值，处理 n_years=1 的情况
    
            if n_years > 1:
                # 计算当前年份在所有年份中的比例 (0 到 1)
                proportion = idx / (n_years - 1)
                # 将比例映射到目标颜色数组的索引范围 (0 到 num_colors_to_use - 1)
                # 例如：3年数据，3个颜色 -> idx=0/2=0 -> prop=0 -> color_idx=0; idx=1/2=0.5 -> prop=0.5 -> color_idx=1; idx=2/2=1 -> prop=1 -> color_idx=2
                # 例如：5年数据，4个颜色 -> idx=0/4=0 -> prop=0 -> color_idx=0; idx=1/4=0.25 -> prop=0.25 -> color_idx=int(0.25*3)=0; idx=2/4=0.5 -> prop=0.5 -> color_idx=int(0.5*3)=1 ...
                color_index = int(proportion * (num_colors_to_use - 1))
    
            # 确保 color_index 不会超出 target_colors 的有效索引范围
            color_index = min(color_index, num_colors_to_use - 1)
            dynamic_year_colors[str(year)] = target_colors[color_index]

    # ── KPMG 色卡 ────────────────────────────────────────────────────────
    # ✅ 关键修复：HTML 必须写成单行紧凑字符串，不能有多余换行/缩进
    with st.expander("🎨 查看 KPMG 官方色卡"):
        for cat_name, cat_colors in KPMG_CATEGORIES.items():
            st.markdown(f"**{cat_name}**")
            html_str = "".join([
                f'<div style="display:inline-block;margin-right:15px;margin-bottom:8px;">'
                f'<div style="width:14px;height:14px;background-color:{c};display:inline-block;'
                f'border-radius:3px;vertical-align:middle;border:1px solid #ddd;"></div>'
                f'<span style="font-size:13px;vertical-align:middle;"> {n} <b>({c})</b></span></div>'
                for n, c in cat_colors.items()
            ])
            st.markdown(html_str, unsafe_allow_html=True)

    st.divider()

    # ── 核心配置面板 ─────────────────────────────────────────────────────
    with st.expander("🛠️ 核心配置面板", expanded=True):
        r1c1, r1c2, r1c3 = st.columns([1.5, 1, 1])
        is_pct_stack_mode = False

        with r1c1:
            chart_type = st.selectbox(
                "📈 1. 图表类型",
                ["簇状柱状图", "堆积柱状图", "折线对比图", "饼图", "内外环结构对比图"]
            )

            # ── 饼图 / 环图 ──────────────────────────────────────────────
            if "环" in chart_type or chart_type == "饼图":
                calc_mode = "结构分析"
                selected_multi_fields = st.multiselect(
                    "🎯 选择构成指标", all_fields,
                    default=all_fields[:min(3, len(all_fields))]
                )
                selected_cos = st.selectbox("🏗️ 选择展示公司", sorted(df_pivot['公司'].unique().tolist()))
                plot_df_base = (
                    df_clean[
                        df_clean['字段名'].isin(selected_multi_fields) &
                        (df_clean['公司'] == selected_cos)
                    ].copy().rename(columns={val_col: 'final_val'})
                )

            else:
                # ── 堆积图子模式 ─────────────────────────────────────────
                if chart_type == "堆积柱状图":
                    stack_sub = st.radio(
                        "堆积模式", ["单指标 / 公式", "多指标占比"],
                        horizontal=True, key="stack_sub_mode"
                    )
                    is_pct_stack_mode = (stack_sub == "多指标占比")

                # ── 多指标占比 ───────────────────────────────────────────
                if is_pct_stack_mode:
                    calc_mode = "多指标占比"
                    if all_types:
                        pct_type_f = st.selectbox(
                            "🗂️ 按类别筛选指标", ["全部类别"] + all_types, key="pct_type_filter"
                        )
                        pct_field_pool = (
                            sorted(df_clean[df_clean['类别'] == pct_type_f]['字段名'].unique().tolist())
                            if pct_type_f != "全部类别" else all_fields
                        )
                    else:
                        pct_field_pool = all_fields

                    selected_pct_fields = st.multiselect(
                        "🎯 选择堆积指标", pct_field_pool,
                        default=pct_field_pool[:min(3, len(pct_field_pool))],
                        key="pct_stack_fields"
                    )
                    if selected_pct_fields:
                        plot_df_base = (
                            df_clean[df_clean['字段名'].isin(selected_pct_fields)]
                            [['公司', '报告年份', '字段名', val_col]]
                            .copy().rename(columns={val_col: 'final_val'})
                        )
                    else:
                        plot_df_base = pd.DataFrame(columns=['公司', '报告年份', '字段名', 'final_val'])

                # ── 普通单指标 / 公式 ────────────────────────────────────
                else:
                    calc_mode = st.radio("数据模式", ["单指标直显", "自定义公式运算"], horizontal=True)

                    if calc_mode == "单指标直显":
                        if all_types:
                            type_f = st.selectbox(
                                "🗂️ 按类别筛选指标", ["全部类别"] + all_types, key="type_filter_single"
                            )
                            filtered_fields = (
                                sorted(df_clean[df_clean['类别'] == type_f]['字段名'].unique().tolist())
                                if type_f != "全部类别" else all_fields
                            )
                        else:
                            filtered_fields = all_fields

                        target_field = st.selectbox("🎯 选择显示指标", filtered_fields)
                        plot_df_base = (
                            df_pivot[['公司', '报告年份', target_field]]
                            .rename(columns={target_field: 'final_val'}).fillna(0)
                        )

                    else:  # 公式模式
                        v1, v2 = st.columns(2)
                        var_a = v1.selectbox("变量 A", all_fields, index=0)
                        var_b = v2.selectbox("变量 B", ["无"] + all_fields, index=1 if len(all_fields) > 1 else 0)
                        formula_input = st.text_input("✏️ 自定义公式 (使用 A 和 B)", value="A - B")
                        try:
                            A_data = df_pivot[var_a].fillna(0)
                            B_data = df_pivot[var_b].fillna(0) if var_b != "无" else 0
                            res = pd.eval(formula_input, local_dict={'A': A_data, 'B': B_data})
                            df_pivot['final_val'] = (
                                pd.Series(res).replace([np.inf, -np.inf], 0).fillna(0).values
                            )
                            plot_df_base = df_pivot[['公司', '报告年份', 'final_val']].copy()
                        except Exception as e:
                            st.error(f"公式无效: {e}")
                            plot_df_base = pd.DataFrame(columns=['公司', '报告年份', 'final_val'])

                # ── 公司类型筛选 ─────────────────────────────────────────
                all_cos_list = sorted(df_pivot['公司'].unique().tolist())
                if all_co_types:
                    co_type_sel = st.selectbox(
                        "🏢 按公司类型快速选择", ["不按公司类型筛选"] + all_co_types, key="co_type_sel"
                    )
                    if (co_type_sel != "不按公司类型筛选" and
                            st.session_state.get('_prev_co_type_sel') != co_type_sel):
                        st.session_state['selected_cos_ms'] = sorted(
                            df_clean[df_clean['公司类型'] == co_type_sel]['公司'].unique().tolist()
                        )
                        st.session_state['_prev_co_type_sel'] = co_type_sel

                selected_cos = st.multiselect(
                    "🏗️ 选择对比公司", all_cos_list,
                    default=all_cos_list[:min(2, len(all_cos_list))],
                    key="selected_cos_ms"
                )

        with r1c2:
            x_axis_mode = (
                st.radio("🔍 布局视角", ["以公司为横轴", "以年份为横轴"])
                if "环" not in chart_type else "结构视角"
            )
            decimals   = st.number_input("🔢 小数位数", min_value=0, max_value=4, value=0)
            show_value = st.toggle("✅ 显示数据标签", value=True)

        with r1c3:
            unit_options  = {"原始数值": 1.0, "亿元": 0.01, "十亿元": 0.001, "百分比(%)": 100.0}
            selected_unit = st.selectbox("📏 数值单位换算", list(unit_options.keys()))
            multiplier    = unit_options[selected_unit]
            y_axis_title  = st.text_input("📝 Y轴单位显示修改", value=f"单位: {selected_unit.split(' ')[0]}")
            is_transparent = st.toggle("🌈 开启透明背景模式")
            show_avg  = st.checkbox("平均值线") if "柱" in chart_type else False
            avg_color = st.color_picker("基准线颜色", value="#ED2124") if show_avg else "#ED2124"

    # ── 绘图数据准备 ─────────────────────────────────────────────────────
    if plot_df_base.empty:
        st.warning("暂无数据，请检查筛选条件。")
        st.stop()

    if calc_mode in ("结构分析", "多指标占比"):
        cos_filter = [selected_cos] if isinstance(selected_cos, str) else selected_cos
        plot_df  = plot_df_base[plot_df_base['公司'].isin(cos_filter)].copy()
        color_val = "字段名"
    else:
        plot_df  = plot_df_base[plot_df_base['公司'].isin(selected_cos)].copy()
        color_val = "报告年份" if x_axis_mode == "以公司为横轴" else "公司"

    plot_df['绘制金额'] = plot_df['final_val'] * multiplier

    # ── 图例颜色配置 ─────────────────────────────────────────────────────
    st.markdown("#### 🎨 自定义图例标签与颜色")
    unique_items = sorted(plot_df[color_val].unique().tolist())
    rename_map, color_map = {}, {}
    c_cols = st.columns(4)
    for i, item in enumerate(unique_items):
        with c_cols[i % 4]:
            st.caption(f"原始值: {item}")
            new_label = st.text_input("显示名称", value=str(item), key=f"rename_{item}")
            default_c = dynamic_year_colors.get(str(item), DEFAULT_COLORS[i % len(DEFAULT_COLORS)])
            new_color = st.color_picker("选择颜色", value=default_c, key=f"color_{item}")
            rename_map[item]     = new_label
            color_map[new_label] = new_color

    plot_df[color_val] = plot_df[color_val].map(rename_map)

    # ── 绘图引擎 ─────────────────────────────────────────────────────────
    fig       = go.Figure()
    fmt       = f'.{decimals}f'
    suffix    = "%" if selected_unit == "百分比(%)" else ""
    label_fmt = f'%{{y:{fmt}}}{suffix}'

    if "柱状图" in chart_type:
        barmode = 'group' if "簇状" in chart_type else 'relative'

        if is_pct_stack_mode:
            plot_df['x_label'] = plot_df['公司'] + ' ' + plot_df['报告年份']
            plot_df['_total']  = plot_df.groupby('x_label')['绘制金额'].transform('sum')
            plot_df['绘制占比'] = (plot_df['绘制金额'] / plot_df['_total'] * 100).fillna(0)
            for item in rename_map.values():
                d = plot_df[plot_df[color_val] == item]
                fig.add_trace(go.Bar(
                    x=d['x_label'], y=d['绘制占比'], name=str(item),
                    marker_color=color_map[item],
                    text=(
                        d.apply(lambda r: f"{r['绘制占比']:.{decimals}f}%<br>{r['绘制金额']:.{decimals}f}", axis=1)
                        if show_value else None
                    ),
                    textposition='inside'
                ))
            fig.update_yaxes(range=[0, 105])
            barmode = 'relative'

        else:
            for item in rename_map.values():
                d = plot_df[plot_df[color_val] == item]
                fig.add_trace(go.Bar(
                    x=d["公司" if color_val == "报告年份" else "报告年份"],
                    y=d["绘制金额"], name=str(item),
                    marker_color=color_map[item],
                    text=d["绘制金额"] if show_value else None,
                    texttemplate=label_fmt if show_value else None,
                    textposition='outside', textangle=0
                ))

        fig.update_layout(barmode=barmode)

    elif "折线对比图" in chart_type:
        for item in rename_map.values():
            d = plot_df[plot_df[color_val] == item]
            fig.add_trace(go.Scatter(
                x=d["公司" if color_val == "报告年份" else "报告年份"],
                y=d["绘制金额"], name=str(item),
                mode='lines+markers+text' if show_value else 'lines+markers',
                marker_color=color_map[item],
                text=d["绘制金额"], texttemplate=label_fmt, textposition="top center"
            ))

    elif chart_type == "饼图":
        latest = plot_df['报告年份'].max()
        d = plot_df[plot_df['报告年份'] == latest]
        fig = px.pie(d, values='绘制金额', names=color_val, hole=0.4,
                     color=color_val, color_discrete_map=color_map)
        fig.update_traces(textinfo='percent+label' if show_value else 'percent')

    elif chart_type == "内外环结构对比图":
        years_ring = sorted(plot_df['报告年份'].unique().tolist())
        if len(years_ring) < 2:
            st.warning("环形图对比需要至少两年的数据。")
        else:
            d_outer = plot_df[plot_df['报告年份'] == years_ring[-1]]
            d_inner = plot_df[plot_df['报告年份'] == years_ring[0]]
            fig.add_trace(go.Pie(
                labels=d_outer[color_val], values=d_outer['绘制金额'],
                hole=0.7, name=years_ring[-1],
                marker=dict(colors=[color_map[f] for f in d_outer[color_val]]),
                textinfo='percent+label' if show_value else 'percent'
            ))
            fig.add_trace(go.Pie(
                labels=d_inner[color_val], values=d_inner['绘制金额'],
                hole=0.4, name=years_ring[0],
                domain={'x': [0.15, 0.85], 'y': [0.15, 0.85]},
                marker=dict(colors=[color_map[f] for f in d_inner[color_val]]),
                textinfo='percent' if show_value else 'none'
            ))
            fig.update_layout(annotations=[dict(
                text=f'内:{years_ring[0]}<br>外:{years_ring[-1]}',
                x=0.5, y=0.5, font_size=12, showarrow=False
            )])

    # ── 全局样式 ─────────────────────────────────────────────────────────
    bg_color = "rgba(0,0,0,0)" if is_transparent else "white"

    if show_avg and "柱" in chart_type and not is_pct_stack_mode:
        avg_v = plot_df['绘制金额'].mean()
        fig.add_hline(y=avg_v, line_dash="dash", line_color=avg_color,
                      annotation_text=f"平均: {avg_v:{fmt}}{suffix}",
                      annotation_font=dict(color=avg_color))

    fig.update_layout(
        font_family="Microsoft YaHei",
        plot_bgcolor=bg_color, paper_bgcolor=bg_color,
        margin=dict(t=120, l=10, r=10, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        annotations=[dict(
            x=0, y=1.18, xref='paper', yref='paper',
            text=f"<b>{y_axis_title}</b>", showarrow=False,
            font=dict(size=14, color="#333333"), xanchor='left'
        )]
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(showgrid=True, gridcolor="#f0f0f0")

    st.plotly_chart(fig, use_container_width=True)

    with st.expander("📄 查看底层数据明细"):
        st.dataframe(plot_df, use_container_width=True)
        
        
        

# ─────────── Step 7 固定图的展示和 PPT ───────────
with tab7:
    st.markdown("### 🖼️ 报告生成版面")
    st.markdown("""
    <style>
    @media print {
        /* 1. 隐藏侧边栏、顶部导航和所有输入控件（按钮、滑块、上传器） */
        section[data-testid="stSidebar"], 
        .stActionButton, 
        header, 
        .stFileUploader, 
        .stSelectbox, 
        .stSlider, 
        .stCheckbox, 
        .stToggle,
        button {
            display: none !important;
        }

        /* 2. 强制内容占满全屏 */
        .main .block-container {
            max-width: 100% !important;
            padding: 0 !important;
        }

        /* 3. 防止图表、分析框跨页断开 */
        div.element-container, .plotly-graph-div {
            page-break-inside: avoid !important;
            margin-bottom: 20px !important;
        }
        
        /* 4. 强制标题换页（可选，如果你想让每个大章节都从新页面开始） */
        h2 {
            page-break-before: always !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    # ==========================================
    # 🌟 新增：内容与注释文件上传区
    # ==========================================
    st.write("#### 📥 1. 上传图表分析与注释")
    st.info("💡 提示：请上传包含【模块ID】、【分析内容】、【注释内容】这三列的 Excel 文件。如果图表不需要文字，可在 Excel 中留空。")
    notes_file = st.file_uploader("上传 Excel 分析注释表", type=['xlsx', 'xls'])
    
    notes_dict = {}
    if notes_file is not None:
        try:
            df_notes = pd.read_excel(notes_file)
            if set(['模块ID', '分析内容', '注释内容']).issubset(df_notes.columns):
                for _, row in df_notes.iterrows():
                    m_id = str(row['模块ID']).strip()
                    notes_dict[m_id] = {
                        'analysis': str(row['分析内容']) if pd.notna(row['分析内容']) else "",
                        'note': str(row['注释内容']) if pd.notna(row['注释内容']) else ""
                    }
                st.success("✅ 注释文件加载成功！内容将自动投射到对应图表中。")
            else:
                st.error("⚠️ 上传的 Excel 缺少必要的列：请确保包含 '模块ID', '分析内容', '注释内容'")
        except Exception as e:
            st.error(f"读取注释文件出错: {e}")

    # ==========================================
    # 🛠️ 核心辅助模块：自动带出分析与注释的渲染函数
    # ==========================================
    def display_notes(module_id):
        """根据模块 ID 渲染上方分析框和下方注释框，并返回文本供 PPT 使用"""
        mod_data = notes_dict.get(module_id, {})
        analysis_text = mod_data.get('analysis', "")
        note_text = mod_data.get('note', "")
        
        if analysis_text:
            st.markdown(f"""
            <div style="background-color: #F0F4FA; border-left: 4px solid #00338D; padding: 15px; margin-bottom: 10px; border-radius: 4px;">
                <p style="margin: 0; color: #1E3A8A; font-size: 14px;"><b>💡 业务分析：</b> {analysis_text}</p>
            </div>
            """, unsafe_allow_html=True)
            
        return analysis_text, note_text

    def display_bottom_note(note_text):
        if note_text:
            st.markdown(f"""
            <div style="background-color: #FFFDF0; border: 1px dashed #D4A373; padding: 10px; margin-top: 10px; margin-bottom: 20px; border-radius: 4px;">
                <p style="margin: 0; color: #666; font-size: 12px;"><b>📌 附注：</b> {note_text}</p>
            </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 检查数据
    # ==========================================
    if 'integrated_data' not in st.session_state or st.session_state['integrated_data'] is None:
        st.warning("⚠️ 请先在 Step 6 完成数据集成。")
    else:
        df_raw = st.session_state['integrated_data'].copy()
        
        # 确保包含公司类型字段，如果没有则伪造一个用于防止报错
        if '公司类型' not in df_raw.columns:
            df_raw['公司类型'] = "未分类"
            
        # ==========================================
        # ⚙️ 全局设置区
        # ==========================================
        st.write("#### ⚙️ 2. 全局图表设置")
        c0, c1, c2, c3 = st.columns([1.5, 2, 1, 1])
        
        with c0:
            # 🌟 新增：公司类型筛选联动
            all_types = ["全部"] + sorted([str(x) for x in df_raw['公司类型'].unique() if str(x) != 'nan'])
            selected_type = st.selectbox("筛选公司类型", options=all_types, index=0)
            
            # 核心数据过滤
            if selected_type != "全部":
                df_filtered = df_raw[df_raw['公司类型'] == selected_type].copy()
            else:
                df_filtered = df_raw.copy()
                
        with c1:
            all_cos = sorted(df_filtered['公司'].unique())
            selected_cos = st.multiselect("选择展示的公司（按顺序）", options=all_cos, default=all_cos)
        with c2:
            # 🌟 全局单位选择器
            unit_label = st.selectbox("全局显示单位", ["十亿元", "亿元", "百万元","十万元"], index=0)
            unit_map = {"十亿元": 1000, "亿元": 100, "百万元": 1, "十万元": 0.1}
            divisor = unit_map[unit_label]
        with c3:
            font_size = st.slider("全局基础字号", 10, 20, 12)

        figs_to_ppt = [] # PPT 导出队列
        COMMON_TITLE_FONT = dict(size=18, color="#00338D", family="Microsoft YaHei")

        # ==========================================
        # 📈 核心绘图逻辑函数区 
        # ==========================================
        # --- 1.柱状图 ---
        def create_kpmg_chart(df, field_name, title_prefix, show_labels, pct_font_size, global_gap):
            d = df[df['公司'].isin(selected_cos)].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            if field_name == "保险服务业绩":
                df_inc = d[d['字段名'] == "保险服务收入合计"].set_index(['公司', '报告年份'])['(百万)人民币']
                df_exp = d[d['字段名'] == "保险服务费用合计"].set_index(['公司', '报告年份'])['(百万)人民币']
                res_series = df_inc - df_exp
                df_plot = res_series.reset_index(); df_plot.columns = ['公司', '报告年份', 'value']
            else:
                df_plot = d[d['字段名'] == field_name].copy().rename(columns={'(百万)人民币': 'value'})

            df_plot['value'] = df_plot['value'] / divisor
            years = sorted(df_plot['报告年份'].unique())
            if len(years) < 2: return go.Figure().add_annotation(text="数据年份不足", showarrow=False)
            
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            color_map = {y_old: "#1E49E2", y_new: "#00338D"}
            all_max = df_plot['value'].max() if not df_plot['value'].empty else 100
            fixed_offset = all_max * 0.1 

            for yr in [y_old, y_new]:
                df_yr = df_plot[df_plot['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=color_map[yr], 
                    text=[f"{v:.1f}" if pd.notna(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside', textfont=dict(size=font_size)
                ))

            df_old, df_new = df_plot[df_plot['报告年份'] == y_old].set_index('公司'), df_plot[df_plot['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v_old, v_new = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v_old) and pd.notna(v_new) and v_old != 0:
                        pct = (v_new - v_old) / abs(v_old)
                        color, arrow = ("#FD349C", "↗") if pct >= 0 else ("#269924", "↘")
                        fig.add_annotation(
                            x=co, y=max(v_old, v_new) + fixed_offset, text=f"<b>{arrow} {pct:.1%}</b>",
                            showarrow=False, font=dict(color=color, size=pct_font_size), xshift=15
                        )
            fig.update_layout(
                title=dict(text=f"<b>{title_prefix}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT, x=0.01),
                barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                bargroupgap=0.0, bargap=global_gap, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                margin=dict(t=100, b=40, l=20, r=20), height=500
            )
            fig.update_xaxes(showgrid=False, zeroline=False, showline=False)
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor="#E0E0E0", zerolinewidth=1,showline=False)
            return fig

        # --- 2.保险业务构成 ---
        def create_kpmg_composition_chart(df, fields, title_prefix, show_labels, label_size, bar_width, co_font_size):
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            if not available_cos: return go.Figure().add_annotation(text="无对应数据", showarrow=False)

            fig = make_subplots(rows=1, cols=len(available_cos), shared_yaxes=True,
                                column_titles=[f"<span style='font-size:{co_font_size}px;'>{co}</span>" for co in available_cos],
                                horizontal_spacing=0.015)
            short_names = {fields[0]: "采用保费分配法", fields[1]: "未采用保费分配法"}
            for i, co in enumerate(available_cos):
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                d_co = d_co.reindex(sorted(d_co.index.tolist()))
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                paa_val, non_paa_val = d_co[fields[0]] / d_co['Total'] * 100, d_co[fields[1]] / d_co['Total'] * 100
                
                fig.add_trace(go.Bar(
                    x=d_co.index, y=non_paa_val, name=short_names[fields[1]] if i == 0 else None, marker_color="rgb(227,207,251)",
                    text=[f"{v:.0f}%" if v > 0 else "" for v in non_paa_val] if show_labels else None,
                    textposition='inside', insidetextanchor='middle', textfont=dict(size=label_size), constraintext='none',
                    width=bar_width, showlegend=(i == 0), legendgroup="group2", hoverinfo="none"
                ), row=1, col=i+1)
                
                fig.add_trace(go.Bar(
                    x=d_co.index, y=paa_val, name=short_names[fields[0]] if i == 0 else None, marker_color="#510DBC",
                    text=[f"{v:.0f}%" if v > 0 else "" for v in paa_val] if show_labels else None,
                    textposition='inside', insidetextanchor='middle', textfont=dict(size=label_size), constraintext='none',
                    width=bar_width, showlegend=(i == 0), legendgroup="group1", hoverinfo="none"
                ), row=1, col=i+1)

            fig.update_layout(title=dict(text=f"<b>{title_prefix}</b>", font=COMMON_TITLE_FONT),
                              barmode='stack', height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=100, b=80, l=20, r=20), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5))
            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(showgrid=False, zeroline=False, linecolor="#E0E0E0", showline=True, tickfont=dict(size=10), row=1, col=i)
                fig.update_yaxes(showgrid=False, zeroline=False, range=[0, 100], showline=False, showticklabels=(i==1), row=1, col=i)
            for ann in fig.layout.annotations:
                ann.y = ann.y + 0.03  # 数字调大就会往上走，比如 0.05, 0.08 等
            return fig
        
        # --- 3.保险业务构成2 ---
        def create_kpmg_multi_composition_chart(df, field_map, color_map, title_prefix, show_labels, label_size, bar_width, co_font_size):
            fields = list(field_map.keys())
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            
            # 注意：这里现在返回两个值，如果没数据就返回空的 图 和 表
            if not available_cos: return go.Figure(), pd.DataFrame()

            # 🌟 新增逻辑：计算各公司的平均占比表格 🌟
            # 1. 整理数据为透视表
            d_pivot = d_struct.pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').fillna(0)
            for f in fields:
                if f not in d_pivot.columns: d_pivot[f] = 0
            d_pivot['Total'] = d_pivot[fields].sum(axis=1).replace(0, 1)
            
            # 2. 计算各指标的百分比
            pct_cols = []
            for f in fields:
                col_name = field_map[f] # 使用展示名称作为列名
                d_pivot[col_name] = d_pivot[f] / d_pivot['Total'] * 100
                pct_cols.append(col_name)
                
            # 3. 按年份求均值，并转置(T)让指标做行，年份做列
            df_avg = d_pivot[pct_cols].groupby('报告年份').mean().T
            # 🌟 结束表格计算 🌟

            fig = make_subplots(rows=1, cols=len(available_cos), shared_yaxes=True,
                                column_titles=[f"<span style='font-size:{co_font_size}px;'>{co}</span>" for co in available_cos],
                                horizontal_spacing=0.015)

            for i, co in enumerate(available_cos):
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                d_co = d_co.reindex(sorted(d_co.index.tolist()))
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                for field_name, display_name in field_map.items():
                    val_pct = d_co[field_name] / d_co['Total'] * 100
                    txt_c = "white" if "30, 73, 226" in color_map[field_name] or "114, 19, 234" in color_map[field_name] else "black"
                    fig.add_trace(go.Bar(
                        x=d_co.index, y=val_pct, name=display_name if i == 0 else None, marker_color=color_map[field_name],
                        text=[f"{v:.0f}%" if v >= 1 else "" for v in val_pct] if show_labels else None,
                        textposition='inside', insidetextanchor='middle', textfont=dict(size=label_size, color=txt_c),
                        constraintext='none', width=bar_width, showlegend=(i == 0), legendgroup=field_name, hoverinfo="none"
                    ), row=1, col=i+1)

            fig.update_layout(title=dict(text=f"<b>{title_prefix}</b>", y=0.98, font=COMMON_TITLE_FONT),
                              barmode='stack', height=600, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=100, b=100, l=20, r=20), legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5, font=dict(size=12)))
            
            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(showgrid=False, zeroline=False, showline=True, linecolor='#E0E0E0', tickfont=dict(size=10), row=1, col=i)
                fig.update_yaxes(showgrid=False, zeroline=False, range=[0, 100], showline=False, showticklabels=(i==1), row=1, col=i)
            
            # 注意：最后返回了 fig 和计算好的 df_avg 表格
            for ann in fig.layout.annotations:
                ann.y = ann.y + 0.03  # 数字调大就会往上走，比如 0.05, 0.08 等
            return fig, df_avg   
        
        # --- 4.利润贡献拆解 ---
        def create_profit_composition_chart(df, selected_cos, target_year, show_labels, label_size, bar_width, co_font_size):
            source_fields = ["合同服务边际的摊销", "非金融风险调整的变动", "亏损部分的确认及转回", "采用保费分配法计量的保险合同保险业绩", "保险利润", "再保净损益"]
            year_str = str(target_year).replace(".0", "")
            d_sub = df[(df['报告年份'].astype(str).str.contains(year_str)) & (df['公司'].isin(selected_cos))].copy()
            d_sub = d_sub[d_sub['字段名'].isin(source_fields)]
            if d_sub.empty: return None, None
                
            d_pivot = d_sub.pivot(index='公司', columns='字段名', values='(百万)人民币')
            contrib_val = (d_pivot['合同服务边际的摊销'] / d_pivot['保险利润'] * 100)
            contribution_df = pd.DataFrame(contrib_val).transpose()
            contribution_df.index = ["合同服务边际释放的贡献"]            
            for f in source_fields:
                if f not in d_pivot.columns: d_pivot[f] = 0.0
            d_pivot = d_pivot.reindex(selected_cos).fillna(0)
        
            d_pivot['营运偏差及其他_展示'] = (d_pivot['保险利润'] - d_pivot['合同服务边际的摊销'] - d_pivot['非金融风险调整的变动'] + d_pivot['亏损部分的确认及转回'] - d_pivot['采用保费分配法计量的保险合同保险业绩'] - d_pivot['再保净损益'])
            d_pivot['合同服务边际的释放_展示'] = d_pivot['合同服务边际的摊销']
            d_pivot['非金融风险调整的变动_展示'] = d_pivot['非金融风险调整的变动']
            d_pivot['亏损部分的确认及转回_展示'] = -d_pivot['亏损部分的确认及转回']
            d_pivot['保费分配法业务净损益_展示'] = d_pivot['采用保费分配法计量的保险合同保险业绩']
            d_pivot['再保净损益_展示'] = d_pivot['再保净损益']
        
            display_mapping = [
                ("亏损部分的确认及转回_展示", "亏损部分的确认及转回", "rgb(190, 190, 190)"),
                ("合同服务边际的释放_展示", "合同服务边际的释放", "rgb(30, 73, 226)"),
                ("非金融风险调整的变动_展示", "非金融风险调整的变动", "rgb(118, 210, 255)"),
                ("营运偏差及其他_展示", "营运偏差及其他", "rgb(114, 19, 214)"),
                ("保费分配法业务净损益_展示", "保费分配法业务净损益", "rgb(253, 52, 156)"),
                ("再保净损益_展示", "再保净损益", "rgb(9, 142, 126)")
            ]
            d_pivot['Total'] = d_pivot[[item[0] for item in display_mapping]].abs().sum(axis=1).replace(0, 1)
        
            fig = go.Figure()
            for col_name, legend_name, color in display_mapping:
                vals_pct = (d_pivot[col_name] / d_pivot['Total']) * 100
                txt_color = "white" if "30, 73, 226" in color or "114, 19, 214" in color or "9, 142, 126" in color else "black"
                fig.add_trace(go.Bar(
                    name=legend_name, x=d_pivot.index, y=vals_pct, width=bar_width, marker_color=color,
                    text=[f"{v:.0f}%" if abs(v) >= 2 else "" for v in vals_pct] if show_labels else None,
                    textposition='inside', insidetextanchor='middle', textfont=dict(size=label_size, color=txt_color), constraintext='none'
                ))
        
            fig.update_layout(
                title=dict(text=f"<b>{year_str}年保险利润构成</b>", x=0.5, y=0.95, font=COMMON_TITLE_FONT),
                barmode='relative', height=600, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=100, b=80, l=220, r=80),
                legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=-0.05, traceorder="reversed", font=dict(size=11, color="#00338D")),
                yaxis=dict(side='right', showgrid=False, range=[-45, 105],tickmode='array',showticklabels=True, tickvals=[-40, -20, 0, 20, 40, 60, 80, 100],
                           ticktext=["-40%", "-20%", "0%", "20%", "40%", "60%", "80%", "100%"], zeroline=True, zerolinecolor="#F7860C")
            )
            x_labels = [f"<span style='font-size:{co_font_size}px;color:#00338D;'>{co}</span>" for co in d_pivot.index]
            fig.update_xaxes(showgrid=False, zeroline=False, tickvals=d_pivot.index, ticktext=x_labels, side="top")
            return fig, contribution_df

        # 数据预处理（图5、6用）
        target_fields = ['净投资回报', '承保财务损益', '分出再保险财务损益']
        df_clean = df_filtered[df_filtered['字段名'].isin(target_fields)].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first')
        df_pivot_5 = df_clean.pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').fillna(0).reset_index()
        df_pivot_5['承保财务净损益'] = -df_pivot_5['承保财务损益'] - df_pivot_5['分出再保险财务损益']
        df_pivot_5['投资利润'] = df_pivot_5['净投资回报'] - df_pivot_5['承保财务净损益']
        df_plot_final = df_pivot_5.melt(id_vars=['公司', '报告年份'], value_vars=['净投资回报', '承保财务净损益', '投资利润'], var_name='指标名称', value_name='value')

        df_profit_raw = df_filtered[(df_filtered['字段名'].isin(['投资利润', '保险利润'])) & (df_filtered['公司'].isin(selected_cos))].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        df_profit_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)

        # --- 5.简单柱状图 ---
        def create_simple_kpmg_chart(df_source, field_name, title, show_labels, p_size, g_gap):
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            color_map = {y_old: "#FD349C", y_new: "#97014F"}
            fig = go.Figure()
            for yr in [y_old, y_new]:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=color_map[yr],
                    text=[f"{v:.1f}" if pd.notna(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside', textfont=dict(size=font_size)
                ))
            all_max = d['value'].max() if not d['value'].empty else 100
            off = all_max * 0.12
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr, col = ("↗", "#FD349C") if pct >= 0 else ("↘", "#269924")
                        fig.add_annotation(x=co, y=max(v0, v1) + off, text=f"<b>{arr} {pct:.1%}</b>", showarrow=False, font=dict(color=col, size=p_size), xshift=10)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT),
                              barmode='group', bargroupgap=0.0, bargap=g_gap, margin=dict(t=100, b=40, l=20, r=20), height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            fig.update_yaxes(showgrid=False, zeroline=True,zerolinecolor="#E0E0E0",zerolinewidth=1)
            fig.update_xaxes(showgrid=False, zeroline=False)
            return fig
        
        # --- 6.利润对比图 ---
        def create_profit_chart_v4(df_source, field_name, title, color_dict, show_labels, p_size, g_gap):
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            for yr in [y_old, y_new]:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=color_dict[yr], 
                    text=[f"{v:.1f}" if pd.notna(v) else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside', textfont=dict(size=font_size)
                ))
            all_max = d['value'].max() if not d['value'].empty else 100
            off = all_max * 0.12
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr, col = ("↗", "#FD349C") if pct >= 0 else ("↘", "#269924")
                        fig.add_annotation(x=co, y=max(v0, v1) + off, text=f"<b>{arr} {pct:.1%}</b>", showarrow=False, font=dict(color=col, size=p_size), xshift=10)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT),
                              barmode='group', bargroupgap=0.0, bargap=g_gap, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=100, b=40, l=20, r=20), height=500, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor="#E0E0E0", zerolinewidth=1.2)
            fig.update_xaxes(showgrid=False, zeroline=False)
            return fig

        # 数据预处理 (税与利润)
        df_tax_sub = df_filtered[(df_filtered['字段名'].isin(['税前利润总额', '净利润'])) & (df_filtered['公司'].isin(selected_cos))].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()
        df_tax_pivot = df_tax_sub.pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').fillna(0).reset_index()
        df_tax_pivot['有效税率'] = np.where(df_tax_pivot['税前利润总额'] != 0, (df_tax_pivot['税前利润总额'] - df_tax_pivot['净利润']) / df_tax_pivot['税前利润总额'], 0)
        df_tax_pivot['报告年份'] = df_tax_pivot['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"

        # --- 7.税前和净利润 ---
        def create_tax_subplot_chart(df_pivot, selected_cos, show_labels, bar_width, label_size, co_font_size, co_y_offset):
            available_cos = [co for co in selected_cos if co in df_pivot['公司'].unique()]
            if not available_cos: 
                return go.Figure()
            
            global_max = (df_pivot[['税前利润总额', '净利润']].max().max() / divisor)
            y_axis_limit = global_max * 1.25 
            tax_label_y = global_max * 1.15 

            fig = make_subplots(
                rows=1, cols=len(available_cos), 
                shared_yaxes=True,
                column_titles=[f"<b>{co}</b>" for co in available_cos],
                horizontal_spacing=0.03
            )

            for i, co in enumerate(available_cos):
                col_idx = i + 1
                d_co = df_pivot[df_pivot['公司'] == co].sort_values('报告年份')
                
                # 1. 税前利润柱子
                fig.add_trace(go.Bar(
                    x=d_co['报告年份'], y=d_co['税前利润总额'] / divisor, 
                    marker_color="#1E49E2",
                    name="税前利润总额",
                    legendgroup="group1",                   # 设置分组名
                    showlegend=(i == 0),                   # 仅当是第一个公司时显示图例
                    text=[f"{v/divisor:.1f}" for v in d_co['税前利润总额']] if show_labels else None,
                    textposition='outside', textfont=dict(size=label_size), 
                    width=bar_width, offsetgroup=1
                ), row=1, col=col_idx)
                
                # 2. 净利润柱子
                fig.add_trace(go.Bar(
                    x=d_co['报告年份'], y=d_co['净利润'] / divisor, 
                    marker_color="#C7A0F7",
                    name="净利润",
                    legendgroup="group2",                   # 设置分组名
                    showlegend=(i == 0),                   # 仅当是第一个公司时显示图例
                    text=[f"{v/divisor:.1f}" for v in d_co['净利润']] if show_labels else None,
                    textposition='outside', textfont=dict(size=label_size), 
                    width=bar_width, offsetgroup=2
                ), row=1, col=col_idx)

                # 3. 绘制有效税率标注
                for idx, row in d_co.iterrows():
                    rate = row['有效税率']
                    t_color = "#97014F" if rate >= 0 else "#269924"
                    fig.add_annotation(
                        x=row['报告年份'], y=tax_label_y, text=f"<b>{rate:.0%}</b>", 
                        showarrow=False, font=dict(size=label_size + 2, color=t_color), 
                        xref=f"x{col_idx}" if col_idx > 1 else "x", yref="y1" 
                    )
                
                # 4. 绘制灰色隔离框
                xref_domain = f"x{col_idx} domain" if col_idx > 1 else "x domain"
                fig.add_shape(
                    type="rect",
                    xref=xref_domain, yref="paper", 
                    x0=-0.05, x1=1.05, y0=0, y1=1,
                    line=dict(color="rgba(200, 200, 200, 0.3)", width=1),
                    fillcolor="rgba(240, 240, 240, 0.1)",
                    layer="below"
                )

            # 5. 统一布局
            fig.update_layout(
                title=dict(
                    text=f"<b>税前利润和净利润变动趋势</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", 
                    font=COMMON_TITLE_FONT, x=0.01, y=0.94
                ),
                height=550, margin=dict(t=120, b=100, l=40, r=40),
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                showlegend=True,
                legend=dict(
                    orientation="h",       
                    yanchor="top", y=-0.15, 
                    xanchor="center", x=0.5
                )
            )

            # 6. 坐标轴更新
            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(
                    showline=True, linecolor='lightgray', linewidth=1, 
                    showgrid=False, tickfont=dict(size=10), zeroline=False, row=1, col=i
                )
                fig.update_yaxes(
                    showline=False, showgrid=False, showticklabels=False, 
                    zeroline=False, range=[0, y_axis_limit], row=1, col=i
                )

            # 7. 公司名称位置微调
            for ann in fig.layout.annotations:
                if ann.text in [f"<b>{c}</b>" for c in available_cos]:
                    ann.y = co_y_offset 
                    ann.font.size = co_font_size

            return fig

        # 预处理 (综合收益)
        df_fin_raw = df_filtered[(df_filtered['字段名'].isin(['净利润', '其他综合收益', '综合收益总额'])) & (df_filtered['公司'].isin(selected_cos))].drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='first').copy()
        df_fin_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)

        # --- 8.综合收益变动趋势 ---
        def create_financial_trend_chart_v5(df_source, field_name, title, show_labels, p_size, g_gap):
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            for yr, col in zip([y_old, y_new], ["rgb(15, 101, 253)", "rgb(0, 51, 141)"]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=col,
                    text=[f"{v:.1f}" if pd.notna(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside', textfont=dict(size=font_size), cliponaxis=False
                ))
            all_vals = d['value'].dropna()
            max_range = all_vals.max() - all_vals.min() if not all_vals.empty else 100
            off = max_range * 0.12 
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color, arr_symbol = ("#FD349C", "↗") if pct >= 0 else ("#269924", "↘")
                        y_pos, v_align = (max(v0, v1) + off, "bottom") if v1 >= 0 else (min(v0, v1) - off, "top")
                        fig.add_annotation(x=co, y=y_pos, text=f"<b>{arr_symbol} {pct:.1%}</b>", showarrow=False, font=dict(color=arr_color, size=p_size), xshift=10, valign=v_align)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT, x=0.01),
                              barmode='group', bargroupgap=0, bargap=g_gap, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=120, b=60, l=40, r=40), height=500, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig.update_yaxes(showgrid=False, zeroline=True)
            fig.update_xaxes(showline=False, showgrid=False, zeroline=False)
            return fig

        # --- 9.资产端分类 ---
        def create_asset_composition_chart(df, selected_cos, field_map, color_map, title_prefix, show_labels, label_size, bar_width, co_font_size):
            fields = list(field_map.keys())
            d = df[df['公司'].isin(selected_cos)].copy()
            d_struct = d[d['字段名'].isin(fields)].copy()
            d_struct['报告年份'] = d_struct['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
            available_cos = [co for co in selected_cos if co in d_struct['公司'].unique()]
            if not available_cos: return go.Figure()
            fig = make_subplots(rows=1, cols=len(available_cos), shared_yaxes=True,
                                column_titles=[f"<span style='font-size:{co_font_size}px; color:#00338D;'>{co}</span>" for co in available_cos], horizontal_spacing=0.015)
            for i, co in enumerate(available_cos):
                d_co = d_struct[d_struct['公司'] == co].pivot(index='报告年份', columns='字段名', values='(百万)人民币')
                for f in fields:
                    if f not in d_co.columns: d_co[f] = 0
                d_co = d_co.reindex(sorted(d_co.index.tolist()))
                d_co['Total'] = d_co.sum(axis=1).replace(0, 1)
                for field_name in fields:
                    val_pct = d_co[field_name] / d_co['Total'] * 100
                    txt_color = "white" if field_name in ["FVOCI", "指定FVOCI", "AC"] else "black"
                    fig.add_trace(go.Bar(
                        x=d_co.index, y=val_pct, name=field_map[field_name] if i == 0 else None, marker_color=color_map[field_name],
                        text=[f"{v:.0f}%" for v in val_pct] if show_labels else None, textposition='inside', insidetextanchor='middle',
                        textfont=dict(size=label_size, color=txt_color), constraintext='none', width=bar_width, showlegend=(i == 0), legendgroup=field_name, hoverinfo="skip"
                    ), row=1, col=i+1)
            fig.update_layout(title=dict(text=f"<b>{title_prefix}</b>", font=COMMON_TITLE_FONT), barmode='stack', height=600,
                              paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', uniformtext=dict(minsize=label_size, mode='show'),
                              margin=dict(t=120, b=120, l=40, r=40), legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="center", x=0.5, font=dict(size=11), itemsizing="constant"))
            for ann in fig.layout.annotations: ann.y = ann.y + 0.05 
            for i in range(1, len(available_cos) + 1):
                fig.update_xaxes(showgrid=False, showline=False, zeroline=False, tickfont=dict(size=10, color="#00338D"), row=1, col=i)
                fig.update_yaxes(showgrid=False, range=[0, 101], showline=False, zeroline=False, tickvals=[0, 25, 50, 75, 100], ticktext=["0%", "25%", "50%", "75%", "100%"], showticklabels=(i==1), row=1, col=i)
            return fig
        
        # --- 10.综合收益拆解图 ---
        def create_oci_chart(df_raw, year, title, show_labels, co_font_size, bar_gap, selected_cos):
            oci_fields = ['可转损益OCI合计', '不可转损益OCI合计', '净利润', '综合收益总额']
            df_raw_copy = df_raw.copy()
            df_raw_copy['报告年份'] = df_raw_copy['报告年份'].astype(str).str.replace(".0", "", regex=False)
            target_year = str(year).replace(".0", "")
            df_sub = df_raw_copy[(df_raw_copy['字段名'].isin(oci_fields)) & (df_raw_copy['报告年份'] == target_year)].copy().drop_duplicates(subset=['公司', '报告年份', '字段名'], keep='last')
            if df_sub.empty: return go.Figure().update_layout(title=f"<b style='color:#00338D;'>{title}</b><br>无数据", paper_bgcolor='rgba(0,0,0,0)')
        
            df_pivot = df_sub.pivot_table(index='公司', columns='字段名', values='(百万)人民币', aggfunc='sum').reset_index().fillna(0)
            for f in oci_fields:
                if f not in df_pivot.columns: df_pivot[f] = 0.0
        
            df_pivot['净利润'] = df_pivot['净利润'] / divisor
            df_pivot['可转损益OCI合计'] = df_pivot['可转损益OCI合计'] / divisor
            df_pivot['不可转损益OCI合计'] = df_pivot['不可转损益OCI合计'] / divisor
            df_pivot['综合收益总额'] = df_pivot['综合收益总额'] / divisor
            df_pivot['其他'] = df_pivot['综合收益总额'] - df_pivot['净利润'] - df_pivot['可转损益OCI合计'] - df_pivot['不可转损益OCI合计']
        
            metrics_config = {
                "净利润": {"color": "rgb(0,176,240)", "name": "净利润"},
                "可转损益OCI合计": {"color": "rgb(253,52,156)", "name": "可转损益OCI变动+FVOCI债券公允价值变动"},
                "不可转损益OCI合计": {"color": "rgb(114,19,234)", "name": "不可转损益负债OCI变动+FVOCI股权公允价值变动"},
                "其他": {"color": "rgb(127,127,127)", "name": "其他"},
                "综合收益总额": {"color": "rgb(172,234,255)", "name": "综合收益总额"}
            }
        
            df_pivot = df_pivot.set_index('公司')
            valid_companies = [c for c in selected_cos if c in df_pivot.index]
            if not valid_companies: return go.Figure()
            all_vals = df_pivot[['净利润', '可转损益OCI合计', '不可转损益OCI合计', '其他', '综合收益总额']].values.flatten()
            max_val, min_val = np.nanmax(all_vals), np.nanmin(all_vals)
            buffer = (max_val - min_val) * 0.25
            y_range = [min_val - buffer if min_val < 0 else 0, max_val + buffer]        
            
            fig = make_subplots(rows=1, cols=len(valid_companies), shared_yaxes=True, horizontal_spacing=0.015, subplot_titles=[""] * len(valid_companies))
            for col_idx, co in enumerate(valid_companies):
                co_data = df_pivot.loc[co]
                for m_key, m_info in metrics_config.items():
                    val = co_data.get(m_key, 0)
                    text_label = f"{val:.0f}" if (show_labels and pd.notna(val) and val != 0) else ""
                    fig.add_trace(go.Bar(
                        name=m_info["name"], x=[m_key], y=[val], text=[text_label], textposition='outside', textfont=dict(size=11, color='#00338D'),
                        marker_color=m_info["color"], width=0.8, legendgroup=m_key, showlegend=(col_idx == 0)
                    ), row=1, col=col_idx + 1)
                fig.update_xaxes(title_text=co, title_font=dict(size=co_font_size, color="#333"), showticklabels=False, showline=False, zeroline=False, showgrid=False, row=1, col=col_idx + 1)
        
            fig.update_layout(title=dict(text=f"<b style='color:#00338D;'>{title}</b>", font=COMMON_TITLE_FONT), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              barmode='group', bargap=bar_gap, height=450, margin=dict(t=60, b=40, l=40, r=20),
                              legend=dict(orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5, font=dict(size=11)),
                              annotations=[dict(text=f"单位：{unit_label}人民币", xref="paper", yref="paper", x=1.0, y=1.12, showarrow=False, font=dict(size=12, color="#666"), xanchor="right")])
            fig.update_yaxes(range=y_range, showline=False, zeroline=False, showgrid=False, row=1, col="all")
            return fig
        
        # --- 11.负债OCI与FVOCI对比 ---
        def create_asset_liab_oci_chart(df_raw, selected_cos, bar_gap, co_font_size, show_labels):
            fields = ['可转损益的负债OCI', 'FVOCI债券公允价值']
            df_sub = df_raw[df_raw['字段名'].isin(fields)].copy()
            df_sub['报告年份'] = df_sub['报告年份'].astype(str).str.replace(".0", "", regex=False)
            years = sorted(df_sub['报告年份'].unique(), reverse=True)[:2]
            years = sorted(years) 
            fig = make_subplots(rows=1, cols=len(selected_cos), shared_yaxes=True, horizontal_spacing=0.01, subplot_titles=[""] * len(selected_cos))
            colors = {"可转损益的负债OCI": "rgb(0, 184, 245)", "FVOCI债券公允价值": "rgb(253, 52, 156)"}
        
            all_vals = df_sub[df_sub['公司'].isin(selected_cos)]['(百万)人民币'] / divisor
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
                        val = val_row['(百万)人民币'].iloc[0] / divisor if not val_row.empty else 0
                        f_vals.append(val)
                        f_texts.append(f"{val:.0f}" if (show_labels and val != 0) else "")
                        f_x.append(f"{y}")
                    fig.add_trace(go.Bar(name=f_name, x=f_x, y=f_vals, text=f_texts, textposition='outside', textfont=dict(size=11, color='#00338D'), marker_color=colors[f_name], legendgroup=f_name, showlegend=(idx == 0), width=0.4), row=1, col=idx + 1)
                fig.update_xaxes(title_text=co, title_font=dict(size=co_font_size), row=1, col=idx+1, showline=False, zeroline=False, showgrid=False)
                fig.update_yaxes(range=y_range, row=1, col=idx+1, showline=False, zeroline=False, showgrid=False)
        
            fig.update_layout(title=dict(text="<b>OCI变动分析 (2/2) - 可转损益负债OCI与FVOCI债权公允价值变动比对</b>", font=COMMON_TITLE_FONT), barmode='group', bargap=bar_gap, height=400,
                              paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=80, b=80, l=50, r=20),
                              legend=dict(orientation="h", yanchor="top", y=-0.28, x=0.5, xanchor="center"),
                              annotations=[dict(text=f"单位：{unit_label}人民币", xref="paper", yref="paper", x=1, y=1.18, showarrow=False, font=dict(size=12, color="#666"))])
            return fig
        
        # --- 12.计算 OCI 分析表 ---
        def calculate_oci_analysis_table(df_raw, selected_cos):
            fields = ['AC', 'FVOCI', 'FVTPL', '指定FVOCI', 'FVOCI债券公允价值', '可转损益的负债OCI']
            df = df_raw[df_raw['字段名'].isin(fields)].copy()
            df['报告年份'] = df['报告年份'].astype(str).str.replace(".0", "", regex=False)
            years = sorted(df['报告年份'].unique(), reverse=True)[:2]
            if len(years) < 2: return pd.DataFrame(), "", ""
            curr_y, prev_y = years[0], years[1]
        
            results = []
            for co in selected_cos:
                data = {}
                for y in [curr_y, prev_y]:
                    for f in fields:
                        val = df[(df['公司'] == co) & (df['报告年份'] == y) & (df['字段名'] == f)]['(百万)人民币']
                        data[f"{f}_{y}"] = val.iloc[0] if not val.empty else 0
                def get_ratio(y):
                    denom = data[f"AC_{y}"] + data[f"FVOCI_{y}"] + data[f"FVTPL_{y}"] + data[f"指定FVOCI_{y}"]
                    return data[f"FVOCI_{y}"] / denom if denom != 0 else 0
                ratio_fvoci = data[f"FVOCI债券公允价值_{curr_y}"] / data[f"FVOCI债券公允价值_{prev_y}"] if data[f"FVOCI债券公允价值_{prev_y}"] != 0 else 0
                ratio_liab = data[f"可转损益的负债OCI_{curr_y}"] / data[f"可转损益的负债OCI_{prev_y}"] if data[f"可转损益的负债OCI_{prev_y}"] != 0 else 0
                ratio_final = ratio_fvoci / ratio_liab if ratio_liab != 0 else 0
                results.append({
                    "公司": co, f"FVOCI占比_{prev_y}": get_ratio(prev_y), f"FVOCI占比_{curr_y}": get_ratio(curr_y),
                    "FVOCI变动": ratio_fvoci, "负债OCI变动": ratio_liab, "两年资负OCI变动比率": ratio_final
                })
            return pd.DataFrame(results), curr_y, prev_y       
        
        # --- 13.净资产与总资产 ---
        def create_asset_trend_chart(df_source, field_name, title, show_labels, p_size, g_gap):
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            for yr, col in zip([y_old, y_new], ["rgb(15, 101, 253)", "rgb(0, 51, 141)"]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=col,
                                     text=[f"{v:.1f}" if pd.notna(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None, textposition='outside', textfont=dict(size=font_size), cliponaxis=False))
            all_vals = d['value'].dropna()
            off = (all_vals.max() - all_vals.min()) * 0.12 if not all_vals.empty else 12
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color, arr_symbol = ("#FD349C", "↗") if pct >= 0 else ("#269924", "↘")
                        y_pos, v_align = (max(v0, v1) + off, "bottom") if v1 >= 0 else (min(v0, v1) - off, "top")
                        fig.add_annotation(x=co, y=y_pos, text=f"<b>{arr_symbol} {pct:.1%}</b>", showarrow=False, font=dict(color=arr_color, size=p_size), xshift=10, valign=v_align)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT, x=0.01),
                              barmode='group', bargroupgap=0, bargap=g_gap, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=120, b=60, l=40, r=40), height=500, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig.update_yaxes(showgrid=False, zeroline=False)
            fig.update_xaxes(showline=False, showgrid=False, zeroline=False)
            return fig
        
        # --- 14.CSM 趋势 ---
        def create_csm_trend_chart(df_source, field_name, title, show_labels, p_size, g_gap):    
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            color_map = {years[-1]: "rgb(0, 51, 141)", years[-2]: "rgb(15, 101, 253)"}
            if len(years) > 2: color_map[years[-3]] = "rgb(173, 216, 230)"
            fig = go.Figure()
            for yr in years:
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(name=f"{yr}YE", x=df_yr['公司'], y=df_yr['value'], marker_color=color_map.get(yr, "gray"),
                                     text=[f"{v:.0f}" if pd.notna(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None, textposition='outside', textfont=dict(size=font_size), cliponaxis=False))
            y_old, y_new = years[-2], years[-1]
            all_vals = d['value'].dropna()
            off = (all_vals.max() - all_vals.min()) * 0.12 if not all_vals.empty else 12
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color, arr_symbol = ("#FD349C", "↗") if pct >= 0 else ("#269924", "↘")
                        y_pos, v_align = (max(v0, v1) + off, "bottom") if v1 >= 0 else (min(v0, v1) - off, "top")
                        fig.add_annotation(x=co, y=y_pos, text=f"<b>{arr_symbol} {pct:.1%}</b>", showarrow=False, font=dict(color=arr_color, size=p_size), xshift=15, valign=v_align)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT, x=0.01),
                              barmode='group', bargroupgap=0, bargap=g_gap, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=120, b=60, l=40, r=40), height=550, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            fig.update_yaxes(showgrid=False, zeroline=False)
            fig.update_xaxes(showline=False, showgrid=False, zeroline=False)
            return fig

        # --- 15. CSM概览表 ---
        def show_csm_summary_table(df, target_year):
            field_map = {"CSM期初余额": "CSM余额", "新业务CSM（集团口径）": "新业务CSM", "CSM计息": "CSM计息", "CSM调整": "CSM调整", "CSM摊销": "CSM摊销", "CSM其他": "其他变动", "CSM期末余额": "CSM期末余额"}
            year_str = str(target_year).replace(".0", "")
            d_sub = df[(df['报告年份'].astype(str).str.contains(year_str)) & (df['字段名'].isin(field_map.keys()))].copy()
            if d_sub.empty: return None
            d_pivot = d_sub.pivot(index='公司', columns='字段名', values='(百万)人民币').rename(columns=field_map).reindex(columns=list(field_map.values())).fillna(0)
            d_pivot = d_pivot / divisor
            def format_number(val):
                if pd.isna(val) or abs(val) < 0.05: return "-"
                elif val < 0: return f"({abs(val):.1f})"
                else: return f"{val:.1f}"
            for col in d_pivot.columns: d_pivot[col] = d_pivot[col].apply(format_number)
            return d_pivot

        # --- 16.摊销前 CSM ---
        def create_csm_composition_chart(df, selected_cos, target_year, show_labels, label_size, bar_width):
            field_map = {"新业务CSM（集团口径）": "新业务 CSM", "CSM计息": "CSM 计息", "CSM调整": "CSM 调整"}
            color_map = {"新业务CSM（集团口径）": "rgb(0, 51, 140)", "CSM计息": "rgb(147, 157, 253)", "CSM调整": "rgb(253, 52, 156)"}
            fields = list(field_map.keys())
            year_str = str(target_year).replace(".0", "")
            d_sub = df[(df['报告年份'].astype(str).str.contains(year_str)) & (df['公司'].isin(selected_cos)) & (df['字段名'].isin(fields))].copy()
            if d_sub.empty: return None
            d_pivot = d_sub.pivot(index='公司', columns='字段名', values='(百万)人民币')
            for f in fields:
                if f not in d_pivot.columns: d_pivot[f] = 0
            d_pivot = d_pivot.reindex(selected_cos).fillna(0)
            d_pivot['Total'] = d_pivot[fields].abs().sum(axis=1).replace(0, 1)
            fig = go.Figure()
            for f_name in fields:
                vals_pct = (d_pivot[f_name] / d_pivot['Total']) * 100
                t_color = "white" if f_name == "新业务CSM（集团口径）" else "black"
                fig.add_trace(go.Bar(name=field_map[f_name], x=d_pivot.index, y=vals_pct, width=bar_width, marker_color=color_map[f_name],
                                     text=[f"{v:.0f}%" if abs(v) >= 1 else "" for v in vals_pct] if show_labels else None,
                                     textposition='inside', insidetextanchor='middle', textfont=dict(size=label_size, color=t_color), constraintext='none'))
            fig.update_layout(title=dict(text=f"<b>CSM核心分析 (2/4) - {year_str}年摊销前CSM变动项占比</b>", x=0.5, y=0.95, font=COMMON_TITLE_FONT),
                              barmode='relative', height=550, margin=dict(t=80, b=100, l=60, r=40), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5),
                              yaxis=dict(showgrid=False, range=[-25, 105], tickvals=[-20, 0, 20, 40, 60, 80, 100], ticktext=["-20%", "0%", "20%", "40%", "60%", "80%", "100%"], zeroline=False), xaxis=dict(showgrid=False, zeroline=False))
            return fig         
        
        # --- 17. CSM 比率折线图 ---  
        def get_company_color_mapping(companies):
            PRESET_COLORS_FIXED = ["#C00000", "#0865EE", "#FEAED7", "#92D050", "#7030A0", "#EF9867", "#61CBF4", "#C7A0F7"]
            cols = st.columns(len(companies))
            mapping = {}
            for i, co in enumerate(companies):
                with cols[i]:
                    st.caption(f"**{co}**")
                    mapping[co] = st.color_picker(label=co, value=PRESET_COLORS_FIXED[i % len(PRESET_COLORS_FIXED)], key=f"cp_compact_{co}", label_visibility="collapsed")
            return mapping

        def create_ratio_line_chart_v3(df_plot, title, color_map, show_labels, marker_size, legend_font_size, y_axis_format=".1%"):
            fig = go.Figure()
            years = sorted(df_plot['报告年份'].unique())
            if not years: return fig
            latest_year = years[-1]
            latest_data = df_plot[df_plot['报告年份'] == latest_year].set_index('公司')['value'].dropna()
            max_co, min_co = latest_data.idxmax() if not latest_data.empty else None, latest_data.idxmin() if not latest_data.empty else None
            
            for co in selected_cos:
                d_co = df_plot[df_plot['公司'] == co].sort_values('报告年份')
                if d_co.empty: continue
                is_highlight = (co == max_co or co == min_co)
                fig.add_trace(go.Scatter(
                    x=d_co['报告年份'], y=d_co['value'], name=co, mode='lines+markers+text' if show_labels else 'lines+markers',
                    line=dict(color=color_map.get(co, "#333"), width=3 if is_highlight else 2, dash="solid" if is_highlight else "dot"),
                    marker=dict(size=marker_size, symbol='circle'),
                    text=[f"{v*100:.1f}%" if (is_highlight or i==0 or i==len(d_co)-1) else "" for i, v in enumerate(d_co['value'])] if show_labels else None,
                    textposition="top center", cliponaxis=False
                ))
            fig.update_layout(title=dict(text=f"<b>{title}</b>", font=COMMON_TITLE_FONT, x=0.01), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=150, b=40, l=60, r=40), height=600, showlegend=True,
                              legend=dict(orientation="h", yanchor="top", y=1.15, xanchor="right", x=1, font=dict(size=legend_font_size), entrywidthmode="pixels", traceorder="normal", valign="middle"),
                              xaxis=dict(showgrid=False, showline=False, zeroline=False), yaxis=dict(showgrid=False, zeroline=False, tickformat=y_axis_format))
            return fig

        # --- 18.新业务盈利合同 ---   
        def create_new_biz_csm_chart(df_source, field_name, title, show_labels, p_size, g_gap):
            d = df_source[df_source['指标名称'] == field_name].copy()
            d['报告年份'] = d['报告年份'].astype(str).str.replace(".0", "", regex=False)
            d['value'] = d['value'] / divisor 
            years = sorted(d['报告年份'].unique())
            if len(years) < 2: return go.Figure()
            y_old, y_new = years[-2], years[-1]
            fig = go.Figure()
            for yr, col in zip([y_old, y_new], ["rgb(15, 101, 253)", "rgb(0, 51, 141)"]):
                df_yr = d[d['报告年份'] == yr].set_index('公司').reindex(selected_cos).reset_index()
                fig.add_trace(go.Bar(
                    name=f"{yr}年新业务盈利合同（CSM）", x=df_yr['公司'], y=df_yr['value'], marker_color=col,
                    text=[f"{v:.1f}" if pd.notna(v) and v != 0 else "" for v in df_yr['value']] if show_labels else None,
                    textposition='outside', textfont=dict(size=font_size), cliponaxis=False
                ))
            all_vals = d['value'].dropna()
            off = (all_vals.max() - all_vals.min()) * 0.12 if not all_vals.empty else 12
            df_old, df_new = d[d['报告年份'] == y_old].set_index('公司'), d[d['报告年份'] == y_new].set_index('公司')
            for co in selected_cos:
                if co in df_old.index and co in df_new.index:
                    v0, v1 = df_old.loc[co, 'value'], df_new.loc[co, 'value']
                    if pd.notna(v0) and v0 != 0:
                        pct = (v1 - v0) / abs(v0)
                        arr_color, arr_symbol = ("#FD349C", "↗") if pct >= 0 else ("#269924", "↘")
                        y_pos, v_align = (max(v0, v1) + off, "bottom") if v1 >= 0 else (min(v0, v1) - off, "top")
                        fig.add_annotation(x=co, y=y_pos, text=f"<b>{arr_symbol} {pct:.1%}</b>", showarrow=False, font=dict(color=arr_color, size=p_size), xshift=10, valign=v_align)
            fig.update_layout(title=dict(text=f"<b>{title}</b><br><span style='font-size:12px;'>单位：({unit_label}人民币)</span>", font=COMMON_TITLE_FONT, x=0.01),
                              barmode='group', bargroupgap=0, bargap=g_gap, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                              margin=dict(t=120, b=60, l=40, r=40), height=500, legend=dict(orientation="v", yanchor="top", y=1.1, xanchor="right", x=1.0))
            fig.update_yaxes(showgrid=False, zeroline=False)
            fig.update_xaxes(showline=False, showgrid=False, zeroline=False)
            return fig

        # --- 19.新业务指标 ---   
        def create_new_business_metrics_charts(df_raw, selected_cos, show_lab, lab_sz, bar_width, co_sz):
            df_clean = df_raw.copy()
            df_clean['报告年份'] = df_clean['报告年份'].astype(str).str.replace('.0', '', regex=False)
            val_col = "(百万)人民币" if "(百万)人民币" in df_clean.columns else df_clean.columns[-1]
            required_fields = ["新业务RA", "新业务亏损合同（LC）——非PAA", "新业务未来现金流入现值", "新业务未来现金流入现值（盈利）", "新业务未来现金流入现值（亏损）", "新业务CSM（集团口径）"]
            df_base = df_clean[df_clean['字段名'].isin(required_fields)].copy()
            df_pivot = df_base.pivot_table(index=['公司', '报告年份'], columns='字段名', values=val_col, aggfunc='sum').reset_index()
            if selected_cos: df_pivot = df_pivot[df_pivot['公司'].isin(selected_cos)]
            if df_pivot.empty: return None, None, None
            df_pivot['新业务CSM利润率'] = np.where(df_pivot.get('新业务未来现金流入现值（盈利）', 0) == 0, np.nan, df_pivot.get('新业务CSM（集团口径）', 0) / df_pivot.get('新业务未来现金流入现值（盈利）', 1))
            df_pivot['新业务LC亏损率'] = np.where(df_pivot.get('新业务未来现金流入现值（亏损）', 0) == 0, np.nan, df_pivot.get('新业务亏损合同（LC）——非PAA', 0) / df_pivot.get('新业务未来现金流入现值（亏损）', 1))
            df_pivot['新业务RA率'] = np.where(df_pivot.get('新业务未来现金流入现值', 0) == 0, np.nan, df_pivot.get('新业务RA', 0) / df_pivot.get('新业务未来现金流入现值', 1))
            years = sorted(df_pivot['报告年份'].unique().tolist())
            if len(years) < 2: return None, None, None
            latest_year, prev_year = years[-1], years[-2]
            configs = [("新业务CSM利润率", "新业务价值分析 (2/4) - 新业务CSM利润率", "rgb(30, 73, 225)", "rgb(149, 229, 255)"), 
                       ("新业务LC亏损率", "新业务价值分析 (3/4) - 新业务LC亏损率", "rgb(253, 52, 156)", "rgb(255, 214, 235)"), 
                       ("新业务RA率", "新业务价值分析 (4/4) - 新业务RA率", "rgb(114, 19, 234)", "rgb(227, 207, 251)")]
            figs = []
            plot_bargap = max(0, 1.0 - bar_width) 
            for metric, title, c_latest, c_prev in configs:
                df_lat, df_pre = df_pivot[df_pivot['报告年份'] == latest_year].dropna(subset=[metric]), df_pivot[df_pivot['报告年份'] == prev_year].dropna(subset=[metric])
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_pre['公司'], y=df_pre[metric], name=f"{prev_year}年", marker_color=c_prev, text=[f"{x*100:.1f}%" if pd.notnull(x) else "" for x in df_pre[metric]] if show_lab else None, textposition='outside', textfont=dict(size=lab_sz)))
                fig.add_trace(go.Bar(x=df_lat['公司'], y=df_lat[metric], name=f"{latest_year}年", marker_color=c_latest, text=[f"{x*100:.1f}%" if pd.notnull(x) else "" for x in df_lat[metric]] if show_lab else None, textposition='outside', textfont=dict(size=lab_sz)))
                margin_b, leg_y, leg_anchor = 20, 1.02, "bottom"
                if not df_lat.empty:
                    avg_val = df_lat[metric].mean()
                    if avg_val < 0: leg_y, leg_anchor, margin_b = -0.15, "top", 60
                    fig.add_hline(y=avg_val, line_dash="dash", line_color=c_latest, line_width=1.5)
                    fig.add_annotation(x=1.0, xref="paper", y=avg_val, yref="y", xanchor="left", xshift=10, text=f"<b>{str(latest_year)[-2:]}年平均 {avg_val*100:.1f}%</b>", showarrow=False, bgcolor="white", bordercolor=c_latest, borderwidth=1.5, borderpad=6, font=dict(color=c_latest, size=lab_sz))
                fig.update_layout(title=dict(text=f"<b>{title}</b>", x=0.5, xanchor='center', y=0.95, font=COMMON_TITLE_FONT), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", barmode='group', bargap=plot_bargap, margin=dict(t=80, b=margin_b, l=20, r=120), yaxis=dict(tickformat=".0%", showgrid=False, zeroline=False), xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=co_sz)), legend=dict(orientation="h", yanchor=leg_anchor, y=leg_y, xanchor="right", x=1))
                figs.append(fig)
            return figs[0], figs[1], figs[2]

        # --- 20.综合净资产指标 --- 
        def create_csm_equity_analysis(df, selected_cos, show_labels, label_size, bar_width, co_font_size, pct_height_adjust):
            years = [2024, 2025]
            d_sub = df[(df['报告年份'].isin(years)) & (df['公司'].isin(selected_cos)) & (df['字段名'].isin(["CSM期末余额", "期末股东权益"]))].copy()
            d_sub['value'] = d_sub['(百万)人民币'] / divisor   
            _all_totals = [d_sub[(d_sub['公司'] == _co) & (d_sub['报告年份'] == _yr)]['value'].sum() for _co in selected_cos for _yr in years]
            global_max_bar = max(_all_totals + [1.0])
            global_base_y, global_y_top = global_max_bar * (1.0 + pct_height_adjust / 100.0), global_max_bar * (1.0 + pct_height_adjust / 100.0) * 1.15
            fig = make_subplots(rows=1, cols=len(selected_cos), horizontal_spacing=0.025)
            for leg_name, leg_color in [("再保前净CSM", "rgb(0, 184, 245)"), ("净资产", "rgb(30, 73, 226)"), ("再保前税后CSM/净资产", "rgb(114, 19, 214)")]:
                fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", marker=dict(symbol="square", size=12, color=leg_color), name=leg_name, showlegend=True), row=1, col=1)
            for i, co in enumerate(selected_cos):
                col_idx, cd = i + 1, d_sub[d_sub['公司'] == co]
                v24b, v25b = cd[(cd['报告年份']==2024) & (cd['字段名']=="CSM期末余额")]['value'].sum(), cd[(cd['报告年份']==2025) & (cd['字段名']=="CSM期末余额")]['value'].sum()
                v24e, v25e = cd[(cd['报告年份']==2024) & (cd['字段名']=="期末股东权益")]['value'].sum(), cd[(cd['报告年份']==2025) & (cd['字段名']=="期末股东权益")]['value'].sum()
                v24c, v25c = v24b*0.75, v25b*0.75
                p24, p25 = (v24c / v24e * 100) if v24e != 0 else 0, (v25c / v25e * 100) if v25e != 0 else 0
                x_axis = ["2024YE", "2025YE"]
                fig.add_trace(go.Bar(x=x_axis, y=[v24c, v25c], marker_color="rgb(0, 184, 245)", width=bar_width, showlegend=False, text=[f"{v:.0f}" for v in [v24c, v25c]] if show_labels else None, textposition='inside', textfont=dict(size=label_size, color="white"), constraintext='none'), row=1, col=col_idx)
                fig.add_trace(go.Bar(x=x_axis, y=[v24e, v25e], marker_color="rgb(30, 73, 226)", width=bar_width, showlegend=False, text=[f"{v:.0f}" for v in [v24e, v25e]] if show_labels else None, textposition='inside', textfont=dict(size=label_size, color="white"), constraintext='none'), row=1, col=col_idx)
                pct_max, delta = max(abs(p24), abs(p25), 1.0), global_max_bar * 0.04
                y24_line, y25_line = global_base_y + (p24 / pct_max) * delta * 0.5, global_base_y + (p25 / pct_max) * delta * 0.5
                fig.add_trace(go.Scatter(x=x_axis, y=[y24_line, y25_line], mode="lines", line=dict(color="rgb(114, 19, 214)", width=1.2), showlegend=False, hoverinfo="skip"), row=1, col=col_idx)
                fig.add_trace(go.Scatter(x=x_axis, y=[y24_line, y25_line], mode="markers+text", marker=dict(symbol="triangle-up", size=7, color="rgb(114, 19, 214)"), text=[f"{p24:.0f}%", f"{p25:.0f}%"], textposition="top center", textfont=dict(color="rgb(114, 19, 214)", size=label_size), showlegend=False, hoverinfo="skip"), row=1, col=col_idx)
                fig.update_yaxes(range=[0, global_y_top], showticklabels=False, showgrid=False, zeroline=False, row=1, col=col_idx)
                xref_d, yref_d = "x domain" if col_idx == 1 else f"x{col_idx} domain", "y domain" if col_idx == 1 else f"y{col_idx} domain"
                fig.add_shape(type="rect", xref=xref_d, yref=yref_d, x0=0, x1=1, y0=0, y1=1, line=dict(color="#CCCCCC", width=1), fillcolor="rgba(200, 200, 200, 0.18)" if i % 2 == 1 else "rgba(255,255,255,0)", layer="below")
                fig.add_annotation(text=f"<b>{co}</b>", xref=xref_d, yref=yref_d, x=0.5, y=1.08, showarrow=False, font=dict(size=co_font_size, color="#00338D"), xanchor="center", yanchor="bottom")
            fig.add_annotation(text=f"(单位：{unit_label})", xref="paper", yref="paper", x=1.0, y=1.0, showarrow=False, font=dict(size=12, color="#00338D"), xanchor="right", yanchor="bottom")
            fig.update_xaxes(type="category", showgrid=False, zeroline=False)       
            fig.update_layout(title=dict(text=f"<b>CSM核心分析 (4/4) - CSM与净资产综合对比</b>", font=COMMON_TITLE_FONT), barmode='stack', height=680, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=130, b=100, l=30, r=30), legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, font=dict(size=12), itemsizing="constant"))
            return fig

        # --- 21. 费用结构 --- 
        def create_expense_breakdown_chart(df, selected_cos, show_labels, label_size, bar_width, co_font_size):
            years, source_fields = [2024, 2025], ["获取费用", "维持费用", "非履约费用"]
            d_sub = df[(df['报告年份'].isin(years)) & (df['公司'].isin(selected_cos)) & (df['字段名'].isin(source_fields))].copy()
            d_sub['value'] = d_sub['(百万)人民币'] / divisor
            avg_acq, avg_maint, avg_non = d_sub[d_sub['字段名'] == "获取费用"]['value'].sum(), d_sub[d_sub['字段名'] == "维持费用"]['value'].sum(), d_sub[d_sub['字段名'] == "非履约费用"]['value'].sum()
            avg_total = avg_acq + avg_maint + avg_non
            p_acq, p_maint, p_non = (avg_acq/avg_total*100) if avg_total>0 else 0, (avg_maint/avg_total*100) if avg_total>0 else 0, (avg_non/avg_total*100) if avg_total>0 else 0
            global_max = max([d_sub[(d_sub['公司']==co) & (d_sub['报告年份']==yr)]['value'].sum() for co in selected_cos for yr in years] + [1.0])
            fig = make_subplots(rows=1, cols=len(selected_cos), horizontal_spacing=0.02)
            for leg_name, leg_color in [("获取费用", "rgb(30, 73, 226)"), ("维持费用", "rgb(118, 210, 255)"), ("非履约费用", "rgb(114, 19, 234)")]:
                fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", marker=dict(symbol="square", size=12, color=leg_color), name=leg_name, showlegend=True), row=1, col=1)
            for i, co in enumerate(selected_cos):
                col_idx = i + 1
                xref_key, yref_key, xref_d, yref_d = f"x{col_idx}" if col_idx>1 else "x", f"y{col_idx}" if col_idx>1 else "y", f"x{col_idx} domain" if col_idx>1 else "x domain", f"y{col_idx} domain" if col_idx>1 else "y domain"
                cd = d_sub[d_sub['公司'] == co]
                gv = lambda yr, field: cd[(cd['报告年份']==yr) & (cd['字段名']==field)]['value'].sum()
                v24a, v25a, v24m, v25m, v24n, v25n = gv(2024,"获取费用"), gv(2025,"获取费用"), gv(2024,"维持费用"), gv(2025,"维持费用"), gv(2024,"非履约费用"), gv(2025,"非履约费用")
                t24, t25 = v24a+v24m+v24n, v25a+v25m+v25n
                lbl = lambda v, t: f"{v/t*100:.0f}%, {v:.0f}" if t>0 and v>0 else ""
                x_axis = ["2024YE", "2025YE"]
                fig.add_trace(go.Bar(x=x_axis, y=[v24a, v25a], marker_color="rgb(30, 73, 226)", width=bar_width, showlegend=False, text=[lbl(v24a,t24), lbl(v25a,t25)] if show_labels else None, textposition='inside', textfont=dict(size=label_size, color="white"), constraintext='none'), row=1, col=col_idx)
                fig.add_trace(go.Bar(x=x_axis, y=[v24m, v25m], marker_color="rgb(118, 210, 255)", width=bar_width, showlegend=False, text=[lbl(v24m,t24), lbl(v25m,t25)] if show_labels else None, textposition='inside', textfont=dict(size=label_size, color="#1a1a2e"), constraintext='none'), row=1, col=col_idx)
                fig.add_trace(go.Bar(x=x_axis, y=[v24n, v25n], marker_color="rgb(114, 19, 234)", width=bar_width, showlegend=False, text=[lbl(v24n,t24), lbl(v25n,t25)] if show_labels else None, textposition='inside', textfont=dict(size=label_size, color="white"), constraintext='none'), row=1, col=col_idx)
                for y0, y1 in [(v24a, v25a), (v24a+v24m, v25a+v25m), (t24, t25)]:
                    fig.add_shape(type="line", xref=xref_key, yref=yref_key, x0=bar_width/2, y0=y0, x1=1-bar_width/2, y1=y1, line=dict(color="rgba(170,170,170,0.85)", width=0.9), layer="above")
                for x_cat, total_v in [("2024YE", t24), ("2025YE", t25)]:
                    fig.add_annotation(x=x_cat, y=total_v+global_max*0.02, xref=xref_key, yref=yref_key, text=f"<b>{total_v:.0f}</b>", showarrow=False, font=dict(size=label_size, color="#222"), bgcolor="white", bordercolor="#BBB", borderwidth=1, xanchor="center", yanchor="bottom")
                pct = ((t25-t24)/t24*100) if t24>0 else 0
                base_arrow_y, slope = max(t24, t25) + global_max*0.12, global_max*0.01
                if pct > 0: y_start, y_end, arr_clr, sign = base_arrow_y-slope, base_arrow_y+slope, "rgb(253, 52, 155)", "+"
                elif pct < 0: y_start, y_end, arr_clr, sign = base_arrow_y+slope, base_arrow_y-slope, "rgb(0, 180, 100)", ""
                else: y_start, y_end, arr_clr, sign = base_arrow_y, base_arrow_y, "rgb(253, 52, 155)", ""
                fig.add_annotation(x=0.65, y=y_end, ax=0.35, ay=y_start, xref=xref_key, yref=yref_key, axref=xref_key, ayref=yref_key, text="", showarrow=True, arrowhead=2, arrowsize=0.7, arrowwidth=1.5, arrowcolor=arr_clr)
                fig.add_annotation(x=0.5, xref=xref_d, y=max(y_start, y_end)+global_max*0.02, yref=yref_key, text=f"<b>{sign}{pct:.0f}%</b>", showarrow=False, font=dict(size=label_size+1, color=arr_clr), xanchor="center", yanchor="bottom")
                fig.add_annotation(text=f"<b>{co}</b>", xref=xref_d, yref=yref_d, x=0.5, y=1.12, showarrow=False, font=dict(size=co_font_size, color="#00338D"), xanchor="center", yanchor="bottom")
                fig.add_shape(type="rect", xref=xref_d, yref=yref_d, x0=0, x1=1, y0=0, y1=1, line=dict(color="#CCCCCC", width=1), fillcolor="rgba(200,200,200,0.18)" if i%2==1 else "rgba(255,255,255,0)", layer="below")
                fig.update_yaxes(range=[0, global_max*1.35], showticklabels=False, showgrid=False, zeroline=False, row=1, col=col_idx)
                fig.update_xaxes(type="category", showgrid=False, zeroline=False, row=1, col=col_idx)
            fig.add_annotation(text=f"(单位：{unit_label})", xref="paper", yref="paper", x=1.0, y=1.05, showarrow=False, font=dict(size=12, color="#00338D"), xanchor="right", yanchor="bottom")
            fig.add_annotation(text=f"<b>全行业平均占比：</b>&nbsp;&nbsp;<span style='color:rgb(30,73,226)'>■ 获取费用 {p_acq:.0f}%</span>&nbsp;&nbsp;|&nbsp;&nbsp;<span style='color:rgb(118,210,255)'>■ 维持费用 {p_maint:.0f}%</span>&nbsp;&nbsp;|&nbsp;&nbsp;<span style='color:rgb(114,19,234)'>■ 非履约费用 {p_non:.0f}%</span>", xref="paper", yref="paper", x=0.5, y=-0.2, showarrow=False, font=dict(size=13, color="#333333"), xanchor="center", yanchor="top")        
            fig.update_layout(title=dict(text=f"<b>运营效能分析 - 费用结构深度剖析</b>", font=COMMON_TITLE_FONT), barmode='stack', height=650, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', margin=dict(t=130, b=80, l=30, r=30), legend=dict(orientation="h", yanchor="top", y=-0.08, xanchor="center", x=0.5, font=dict(size=12), itemsizing="constant"))
            return fig

        # --- 22. 六维图 ---
        def create_six_dimensional_charts(df_raw, target_year, label_size, show_labels, dot_size):
            df = df_raw.copy()
            df['字段名'] = df['字段名'].fillna('未知').astype(str)
            df['报告年份'] = df['报告年份'].astype(str).str.replace('.0', '', regex=False)
            df_curr = df[df['报告年份'] == str(target_year)].copy()
            df_pivot = df_curr.pivot_table(index='公司', columns='字段名', values='(百万)人民币', aggfunc='sum').fillna(0)
            morandi_colors = ['#577590', '#4D908E', '#43AA8B', '#90BE6D', '#F9C74F', '#F9844A', '#F8961E', '#F3722C', '#F94144', '#277DA1', '#B5838D', '#6D597A', '#E5989B', '#B7B7A4', '#A5A58D']
            companies = df_pivot.index.tolist()
            color_map = {cos: morandi_colors[i % len(morandi_colors)] for i, cos in enumerate(companies)}
            def get_col(name): return df_pivot[name] if name in df_pivot.columns else pd.Series(0, index=df_pivot.index)
            plot_data = pd.DataFrame(index=df_pivot.index)
            plot_data['利润率'] = get_col('净利润') / ((get_col('期初股东权益') + get_col('期末股东权益')) / 2).replace(0, np.nan)
            plot_data['CSM增长率'] = (get_col('CSM期末余额') - get_col('CSM期初余额')) / get_col('CSM期初余额').replace(0, np.nan)
            plot_data['财务杠杆率'] = get_col('期末股东权益') / get_col('总资产').replace(0, np.nan)
            plot_data['投资收益率'] = get_col('投资收益率')
            plot_data['净资产增长率'] = (get_col('期末股东权益') - get_col('期初股东权益')) / get_col('期初股东权益').replace(0, np.nan)
            plot_data['综合偿付能力充足率'] = get_col('综合偿付能力充足率')
            plot_data = plot_data.fillna(0)
            configs = [("股东回报", "利润率", "净利润", ".1%", "Y轴=利润率，X轴净利润"), ("盈利潜力", "CSM增长率", "CSM期末余额", ".1%", "Y轴=CSM增长率，X轴期末CSM"), ("财务杠杆", "财务杠杆率", "期末股东权益", ".1%", "Y轴=杠杆率，X轴净资产"), ("投资能力", "投资收益率", "期末股东权益", ".1%", "Y轴投资收益率，X轴净资产"), ("财务稳定", "净资产增长率", "期末股东权益", ".1%", "Y轴净资产增长率，X轴净资产"), ("偿付能力", "综合偿付能力充足率", "期末股东权益", ".0%", "Y轴综合偿付能力，X轴净资产")]
            figs = []
            for title, y_col, x_col, y_fmt, subtitle in configs:
                fig = go.Figure()
                y_vals = plot_data[y_col]
                x_raw = get_col(x_col) if x_col in df_pivot.columns else plot_data.get(x_col, pd.Series(0, index=plot_data.index))
                x_vals = x_raw / divisor
                for cos in companies:
                    y_val, x_val = y_vals.loc[cos], x_vals.loc[cos]
                    pos = "top center" if y_val < y_vals.max() else "bottom center"
                    fig.add_trace(go.Scatter(x=[x_val], y=[y_val], mode='markers+text' if show_labels else 'markers', name=cos, text=[f"<b>{cos}</b><br>{y_val*100:.1f}%"], textposition=pos, textfont=dict(size=label_size, color='#333'), cliponaxis=False, marker=dict(size=dot_size, color=color_map[cos], line=dict(width=1.5, color='rgba(255,255,255,0.9)'), opacity=0.9)))
                if len(y_vals) > 0:
                    avg_y = y_vals.mean()
                    fig.add_hline(y=avg_y, line_dash="dash", line_color="#FF69B4", line_width=1.5)
                    fig.add_annotation(xref="paper", x=1.02, y=avg_y, text=f"平均:{avg_y*100:.1f}%", showarrow=False, font=dict(color="#FF69B4", size=10, weight='bold'), xanchor="left")
                fig.update_layout(title=dict(text=f"<span style='font-size:16px;color:#222'><b>{title}</b></span><br><span style='font-size:10px;color:#888'>{subtitle}（单位：{unit_label}）</span>", x=0.05, y=0.92, xanchor='left'), plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=100, b=50, l=50, r=80), showlegend=False, height=480, xaxis=dict(showgrid=False, zeroline=False, tickfont=dict(size=9)), yaxis=dict(showgrid=False, zeroline=False, tickformat=y_fmt, range=[y_vals.min() - abs(y_vals.min()*0.1), y_vals.max() + abs(y_vals.max()*0.2)] if not y_vals.empty else None))
                fig.update_xaxes(showline=False, zeroline=False)
                fig.update_yaxes(showline=False, zeroline=False)
                figs.append(fig)
            return figs

        # --- 23. 业绩明细表 ---
        def create_financial_report_table(df_raw, target_year, selected_cos):
            df = df_raw.copy()
            df['报告年份'] = df['报告年份'].astype(str).str.replace('.0', '', regex=False)
            df_year = df[(df['报告年份'] == str(target_year)) & (df['公司'].isin(selected_cos))]
            val_col = "(百万)人民币" if "(百万)人民币" in df.columns else df.columns[-1]
            data_map = df_year.set_index(['公司', '字段名'])[val_col].to_dict()
            def get_val(co, field): return data_map.get((co, field), 0)
            rows_config = [
                ("未采用保费分配法的保险合同", None, "header"), ("保险服务收入", "未采用保费分配法计量的保险合同保险服务收入", "data"), ("合同服务边际的释放", "合同服务边际的摊销", "data"), ("非金融风险调整的变动", "非金融风险调整的变动", "data"), ("预期当期发生的保险服务费用", "预计当期发生的保险服务费用", "data"), ("保险获取现金流的摊销", "保险获取现金流的摊销（保险服务收入）", "data"), ("其他", "其他收入调整", "data"),
                ("保险服务费用", "未采用保费分配法计量的保险合同保险服务费用", "neg_data"), ("保险获取现金流的摊销 ", "保险获取现金流的摊销（保险服务费用）", "neg_data"), ("亏损部分的确认及转回", "亏损部分的确认及转回", "neg_data"), ("当期发生的赔款及其他相关费用", "当期发生的赔款及其他相关费用", "neg_data"), ("已发生赔款负债相关的履约现金流量变动", "已发生赔款负债相关的履约现金流量变动", "neg_data"), ("其他项", "FIXED_ZERO", "data"),
                ("保险服务业绩", "FORMULA_KPI_1", "subtotal"), 
                ("采用保费分配法的保险合同", None, "header"), ("保险服务收入 ", "采用保费分配法计量的保险合同保险服务收入", "data"), ("保险服务费用 ", "采用保费分配法计量的保险合同保险服务费用", "neg_data"), ("保险服务业绩 ", "采用保费分配法计量的保险合同保险业绩", "subtotal"),
                ("集团/业务条线-汇总（见附注）", None, "header"), ("保险服务收入  ", "保险服务收入合计", "data"), ("保险服务费用  ", "保险服务费用合计", "neg_data"), ("保险服务业绩（不含再保收支净额）", "FORMULA_TOTAL_KPI", "total"), ("CSM释放在保险服务业绩中占比", "FORMULA_CSM_RATIO", "percent"),
            ]
            report_data = []
            for row_name, field, row_type in rows_config:
                row_values = {f"项目名称 (单位: 百万人民币)": row_name}
                for co in selected_cos:
                    val = np.nan
                    if row_type == "header": pass
                    elif row_type == "data": val = get_val(co, field) if field != "FIXED_ZERO" else 0
                    elif row_type == "neg_data": val = -abs(get_val(co, field))
                    elif row_type == "subtotal": val = get_val(co, "未采用保费分配法计量的保险合同保险服务收入") - get_val(co, "未采用保费分配法计量的保险合同保险服务费用") if row_name == "保险服务业绩" else get_val(co, "采用保费分配法计量的保险合同保险业绩")
                    elif row_type == "total": val = get_val(co, "保险服务收入合计") - get_val(co, "保险服务费用合计")
                    elif row_type == "percent":
                        denom = get_val(co, "未采用保费分配法计量的保险合同保险服务收入") - get_val(co, "未采用保费分配法计量的保险合同保险服务费用")
                        val = (get_val(co, "合同服务边际的摊销") / denom) if denom != 0 else 0
                    row_values[co] = val
                report_data.append(row_values)
            return pd.DataFrame(report_data)

        def format_financial_df(df):
            def _style_val(v):
                if pd.isna(v): return "-"
                if isinstance(v, float) and 0 < abs(v) < 1.5: return f"{v*100:.0f}%"
                if v < 0: return f"({abs(v):,.0f})"
                if v == 0: return "0"
                return f"{v:,.0f}"
            styled_df = df.copy()
            for col in styled_df.columns[1:]: styled_df[col] = styled_df[col].apply(_style_val)
            return styled_df

        # ==========================================
        # 🚀 图表调用与渲染区 (完全对标咨询报告命名)
        # ==========================================

        # ---------------- Section 1 ----------------
        st.write("---")
        st.write("## 1. 整体业绩表现分析")

        tasks_1 = [("inc_total", "保险服务收入合计", "保险服务收入变动趋势"), 
                   ("exp_total", "保险服务费用合计", "保险服务费用变动趋势"), 
                   ("perf_total", "保险服务业绩", "保险服务业绩变动趋势")]
        ctrl1, ctrl2, ctrl3 = st.columns([1.5, 2, 3])
        with ctrl1: ui_t1_lab = st.toggle(f"显示柱状图数据标签", value=True, key="t1_lab")
        with ctrl2: ui_t1_psize = st.slider("涨幅文字大小", 8, 24, 14, key="t1_psize")
        with ctrl3: ui_t1_gap = st.slider("调整公司间距与柱子粗细", 0.1, 0.8, 0.3, key="t1_gap")

        for m_id, field, title in tasks_1:
            an, nt = display_notes(m_id)
            fig = create_kpmg_chart(df_filtered, field, title, ui_t1_lab, ui_t1_psize, ui_t1_gap)
            if fig:
                st.plotly_chart(fig, use_container_width=True)
                figs_to_ppt.append((title, fig, an, nt))
            display_bottom_note(nt)

        st.markdown("---")
        an, nt = display_notes("comp_1")
        comp_title, comp_fields = "保险业务收入构成 (1/2) - PAA与非PAA对比", ["采用保费分配法计量的保险合同保险服务收入", "未采用保费分配法计量的保险合同保险服务收入"]
        c_ctrl1, c_ctrl2, c_ctrl3, c_ctrl4 = st.columns([1, 1.5, 1.5, 1.5])
        with c_ctrl1: label_on_comp = st.toggle("显示数据标签", value=True, key="lab_comp")
        with c_ctrl2: l_size = st.slider("标签大小", 5, 20, 12, key="l_size_comp")
        with c_ctrl3: b_width = st.slider("柱子宽度", 0.1, 1.0, 0.6, 0.05, key="b_width_comp")
        with c_ctrl4: c_f_size = st.slider("公司名称大小", 10, 20, 14, key="c_f_size_comp")
        fig_comp = create_kpmg_composition_chart(df_filtered, comp_fields, comp_title, label_on_comp, l_size, b_width, c_f_size)
        st.plotly_chart(fig_comp, use_container_width=True)
        figs_to_ppt.append((comp_title, fig_comp, an, nt))
        display_bottom_note(nt)

        an, nt = display_notes("comp_2")
        title_2 = "保险业务收入构成 (2/2) - 详细收入构成"
        field_map_2 = {"合同服务边际的摊销": "合同服务边际的释放", "非金融风险调整的变动": "非金融风险调整的变动", "预计当期发生的保险服务费用": "预期当期发生的保险服务费用", "保险获取现金流的摊销（保险服务收入）": "保险获取现金流的摊销", "与当期服务或过去服务相关得保费经验调整": "与当期服务或过去服务相关的保费经验调整", "其他收入调整": "其他"}
        color_map_2 = {"合同服务边际的摊销": "rgb(30, 73, 226)", "非金融风险调整的变动": "rgb(254, 174, 215)", "预计当期发生的保险服务费用": "rgb(0, 163, 161)", "保险获取现金流的摊销（保险服务收入）": "rgb(1, 184, 245)", "与当期服务或过去服务相关得保费经验调整": "rgb(0, 219, 214)", "其他收入调整": "rgb(114, 19, 234)"}
        
        # 接收图表(fig_multi)和均值表格(df_multi_avg)
        fig_multi, df_multi_avg = create_kpmg_multi_composition_chart(df_filtered, field_map_2, color_map_2, title_2, label_on_comp, l_size, b_width, c_f_size)
        
        if fig_multi: 
            # 1. 如果有数据，先在页面展示算出来的均值表格
            if not df_multi_avg.empty:
                st.write("##### 📊 各公司平均占比情况")
                # 使用 .style.format 给表格加上百分号
                st.dataframe(df_multi_avg.style.format("{:.1f}%"), use_container_width=True)
                
            # 2. 展示图表
            st.plotly_chart(fig_multi, use_container_width=True)
            
            # 3. 将图表加入 PPT
            figs_to_ppt.append((title_2, fig_multi, an, nt))
            # 4. 可选：把算出来的这个表格也一并加进 PPT 导出队列
            if not df_multi_avg.empty:
                figs_to_ppt.append((title_2 + " 平均占比", df_multi_avg, "", ""))
                
        display_bottom_note(nt)

        # 📍 [预留位置：如果后续要在这里插入图表，请加代码]

        # ---------------- Section 2 ----------------
        st.markdown("---")
        st.write("## 2. 利润与综合收益分析")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1: ui_prof_lab = st.toggle("显示利润标签", value=False, key="prof_lab")
        with c2: ui_prof_sz = st.slider("利润标签字号", 8, 16, 11, key="prof_sz")
        with c3: ui_prof_bw = st.slider("利润柱子宽度", 0.2, 0.8, 0.4, key="prof_bw")
        with c4: ui_prof_co = st.slider("利润公司字号", 10, 20, 14, key="prof_co")

        for year_val, mod_id in [(2025, "prof_2025"), (2024, "prof_2024")]:
            an, nt = display_notes(mod_id)
            fig_p, contrib_p = create_profit_composition_chart(df_filtered, selected_cos, year_val, ui_prof_lab, ui_prof_sz, ui_prof_bw, ui_prof_co)
            if fig_p:
                st.write(f"### 📅 利润表现分析 - {year_val}年保险利润构成")
                st.dataframe(contrib_p.style.format("{:.0f}%"), use_container_width=True)
                st.plotly_chart(fig_p, use_container_width=True)
                figs_to_ppt.append((f'{year_val}年保险利润构成', fig_p, an, nt))
            display_bottom_note(nt)

        st.markdown("---")
        display_tasks = [("inv_return", "净投资回报", "净投资回报变动趋势", "直接提取"), 
                         ("uw_profit", "承保财务净损益", "承保财务净损益变动趋势", "公式计算"), 
                         ("inv_profit", "投资利润", "投资利润变动趋势", "公式计算")]
        c1, c2, c3 = st.columns([2, 2, 3])
        with c1: lab_inv = st.toggle("显示趋势图标签", True)
        with c2: psz_inv = st.slider("趋势图文字大小", 8, 20, 12)
        with c3: gap_inv = st.slider("趋势图柱子粗细", 0.1, 0.8, 0.3)
        for mod_id, field, title, _ in display_tasks:
            an, nt = display_notes(mod_id)
            fig = create_simple_kpmg_chart(df_plot_final, field, f"投资表现分析 - {title}", lab_inv, psz_inv, gap_inv)
            st.plotly_chart(fig, use_container_width=True)
            figs_to_ppt.append((title, fig, an, nt))
            display_bottom_note(nt)

        st.markdown("---")
        available_years = sorted(df_profit_raw['报告年份'].astype(str).str.replace(".0", "", regex=False).unique())
        if len(available_years) >= 2:
            yr_old, yr_new = available_years[-2], available_years[-1]
            for field, title, c_dict, mod_id in [("投资利润", "投资利润变动对比分析", {yr_old: "#FC349B", yr_new: "#97014F"}, "trend_inv"), 
                                                 ("保险利润", "保险利润变动对比分析", {yr_old: "#1E49E2", yr_new: "#00338C"}, "trend_ins")]:
                an, nt = display_notes(mod_id)
                fig = create_profit_chart_v4(df_profit_raw, field, title, c_dict, lab_inv, psz_inv, gap_inv)
                st.plotly_chart(fig, use_container_width=True)
                figs_to_ppt.append((title, fig, an, nt))
                display_bottom_note(nt)

        st.markdown("---")
        an, nt = display_notes("tax_profit")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: ui_show_lab = st.toggle("税图标签", value=True)
        with c2: ui_bar_w = st.slider("税图柱宽", 0.1, 0.8, 0.4)
        with c3: ui_fs = st.slider("税图字号", 8, 16, 10)
        with c4: ui_cfs = st.slider("公司名大小", 10, 24, 14)
        with c5: ui_co_y = st.slider("税图高度", 1.0, 1.2, 1.05, step=0.01)
        fig_tax = create_tax_subplot_chart(df_tax_pivot, selected_cos, ui_show_lab, ui_bar_w, ui_fs, ui_cfs, ui_co_y)
        st.plotly_chart(fig_tax, use_container_width=True)
        figs_to_ppt.append(("税前利润和净利润变动趋势", fig_tax, an, nt))
        display_bottom_note(nt)

        # 综合收益变动趋势 (三图共用一组按钮) ----------------
        st.markdown("---")
        st.write("### 综合收益变动趋势概览")

        # 1. 紧凑排布的专属 UI 控制面板 (放在循环最前面)
        ui_f1, ui_f2, ui_f3 = st.columns(3)
        with ui_f1: lab_fin = st.toggle("显示收益图标签", True, key="fin_lab")
        with ui_f2: psz_fin = st.slider("涨幅文字字号", 8, 24, 12, key="fin_psz")
        with ui_f3: gap_fin = st.slider("柱子间距(粗细)", 0.05, 0.8, 0.3, key="fin_gap")

        # 2. 定义这三张图的配置项
        fin_targets = [
            ('净利润', '综合收益概览 (1/3) - 净利润(亏损)趋势', "net_profit"), 
            ('其他综合收益', '综合收益概览 (2/3) - OCI变动趋势', "oci_profit"), 
            ('综合收益总额', '综合收益概览 (3/3) - 综合收益总额趋势', "total_profit")
        ]

        # 3. 循环渲染，并统一传入上面定义的专属按钮参数
        for metric, title, m_id in fin_targets:
            an, nt = display_notes(m_id)
            # 👇 核心：传入 lab_fin, psz_fin, gap_fin
            fig_fin = create_financial_trend_chart_v5(df_fin_raw, metric, title, lab_fin, psz_fin, gap_fin)
            
            if fig_fin:
                st.plotly_chart(fig_fin, use_container_width=True)
                figs_to_ppt.append((title, fig_fin, an, nt))
            display_bottom_note(nt)

        # 📍 [预留位置：利润分析补充图表]


        # ---------------- Section 3 ----------------
        st.markdown("---")
        st.write("## 3. 资产负债与 OCI 结构分析")
        
        ui_ast1, ui_ast2, ui_ast3 = st.columns(3)
        with ui_ast1: lab_ast = st.toggle("显示资产标签", True, key="ast_l")
        with ui_ast2: bw_ast = st.slider("资产柱宽", 0.1, 1.0, 0.6, key="ast_w")
        with ui_ast3: sz_ast = st.slider("资产字号", 8, 20, 11, key="ast_s")

        an, nt = display_notes("asset_struct")
        field_map_asset = {"AC": "债权投资", "FVOCI": "其他债权投资", "FVTPL": "交易性金融资产", "指定FVOCI": "其他权益工具投资"}
        color_map_asset = {"AC": "rgb(0, 184, 245)", "FVOCI": "rgb(114, 19, 234)", "FVTPL": "rgb(253, 52, 156)", "指定FVOCI": "rgb(181, 2, 95)"}
        fig_asset_comp = create_asset_composition_chart(df_filtered, selected_cos, field_map_asset, color_map_asset, "资产配置结构分析 - 各公司配置变动", lab_ast, sz_ast, bw_ast, 14)
        st.plotly_chart(fig_asset_comp, use_container_width=True)
        figs_to_ppt.append(("资产配置结构分析", fig_asset_comp, an, nt))
        display_bottom_note(nt)

        st.markdown("---")
        # 共用的年份OCI控制
        ui_oci1, ui_oci2, ui_oci3 = st.columns(3)
        with ui_oci1: lab_oci = st.toggle("显示OCI标签", True, key="oci_l")
        with ui_oci2: gap_oci = st.slider("OCI柱间距", 0.05, 0.5, 0.15, key="oci_g")
        with ui_oci3: sz_oci = st.slider("OCI公司字号", 8, 20, 13, key="oci_s")

        clean_years = [str(y).replace(".0", "") for y in sorted(df_filtered['报告年份'].dropna().unique(), reverse=True)][:2] 
        oci_tasks = [(y, f"OCI变动分析 - {y}年综合收益变动情况", "oci_year_lat" if i==0 else "oci_year_pre") for i, y in enumerate(clean_years)]
        for year_field, title, m_id in oci_tasks:
            an, nt = display_notes(m_id)
            fig_oci = create_oci_chart(df_filtered, year_field, title, lab_oci, sz_oci, gap_oci, selected_cos)
            st.plotly_chart(fig_oci, use_container_width=True)
            figs_to_ppt.append((title, fig_oci, an, nt))
            display_bottom_note(nt)

        an, nt = display_notes("oci_deep")
        df_analysis, c_y, p_y = calculate_oci_analysis_table(df_filtered, selected_cos)
        if not df_analysis.empty:
            fig_deep = create_asset_liab_oci_chart(df_filtered, selected_cos, gap_oci, sz_oci, lab_oci)
            st.plotly_chart(fig_deep, use_container_width=True)
            figs_to_ppt.append(("资负深度分析图", fig_deep, an, nt))
        display_bottom_note(nt)

        #OCI指标表
        an, nt = display_notes("oci_deep")
        df_res, c_y, p_y = calculate_oci_analysis_table(df_filtered, selected_cos)

        if not df_res.empty:
            st.write(f"#####  资负 OCI 变动分析表 ({p_y}YE - {c_y}YE)")
            df_t = df_res.set_index("公司").T
            df_t = df_t.rename(index={
                f"FVOCI占比_{p_y}": f"FVOCI占比({p_y})", 
                f"FVOCI占比_{c_y}": f"FVOCI占比({c_y})",
                "FVOCI变动": "FVOCI变动", 
                "负债OCI变动": "负债OCI变动", 
                "两年资负OCI变动比率": "资负匹配率"
            })
            st.dataframe(df_t.style.format("{:.2%}"), use_container_width=True)
            df_ppt = df_t.reset_index().rename(columns={"index": "指标名称"})
            for col in df_ppt.columns[1:]:
                df_ppt[col] = df_ppt[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else "-")               
            figs_to_ppt.append((f"资负OCI匹配分析-{c_y}", df_ppt, an, nt))            
        display_bottom_note(nt)



        st.markdown("---")
        ui_n1, ui_n2, ui_n3 = st.columns(3)
        with ui_n1: lab_ast_t = st.toggle("显示净资产标签", True, key="n_l")
        with ui_n2: psz_ast_t = st.slider("净资产趋势字号", 8, 24, 11, key="n_s")
        with ui_n3: gap_ast_t = st.slider("净资产趋势柱距", 0.1, 0.8, 0.3, key="n_g")

        df_asset_raw = df_filtered[df_filtered['字段名'].isin(['期末股东权益', '总资产']) & df_filtered['公司'].isin(selected_cos)].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()
        df_asset_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)
        for field, title, m_id in [('期末股东权益', '规模趋势分析 (1/2) - 净资产变动', "equity_trend"), ('总资产', '规模趋势分析 (2/2) - 总资产变动', "asset_trend")]:
            an, nt = display_notes(m_id)
            fig_ast = create_asset_trend_chart(df_asset_raw, field, title, lab_ast_t, psz_ast_t, gap_ast_t)
            st.plotly_chart(fig_ast, use_container_width=True)
            figs_to_ppt.append((title, fig_ast, an, nt))
            display_bottom_note(nt)

        # 📍 [预留位置：资产负债补充图表]


        # ---------------- Section 4 ----------------
        st.markdown("---")
        st.write("## 4. CSM (合同服务边际) 核心表现追踪")

        ui_csm1, ui_csm2, ui_csm3 = st.columns(3)
        with ui_csm1: lab_csm_t = st.toggle("显示CSM标签", True, key="c_l")
        with ui_csm2: psz_csm_t = st.slider("CSM趋势字号", 8, 24, 12, key="c_s")
        with ui_csm3: gap_csm_t = st.slider("CSM趋势柱宽", 0.1, 0.8, 0.35, key="c_g")

        df_csm_raw = df_filtered[(df_filtered['字段名'] == 'CSM期末余额') & df_filtered['公司'].isin(selected_cos)].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()
        df_csm_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)
        
        an, nt = display_notes("csm_bal")
        fig_csm = create_csm_trend_chart(df_csm_raw, 'CSM期末余额', 'CSM核心分析 (1/4) - CSM余额变动趋势', lab_csm_t, psz_csm_t, gap_csm_t)
        st.plotly_chart(fig_csm, use_container_width=True)
        figs_to_ppt.append(('CSM余额变动趋势', fig_csm, an, nt))
        display_bottom_note(nt)

        an, nt = display_notes("csm_table")
        table_df_2025 = show_csm_summary_table(df_filtered, 2025)
        if table_df_2025 is not None:
            st.write(f"### 2025年度 CSM 概览明细表 (单位:{unit_label})")
            st.dataframe(table_df_2025, use_container_width=True)
            figs_to_ppt.append((f'2025年CSM概览表_单位{unit_label}', table_df_2025, an, nt))
        display_bottom_note(nt)

        ui_cc1, ui_cc2, ui_cc3 = st.columns(3)
        with ui_cc1: lab_csm_c = st.toggle("摊销前占比标签", True, key="cc_l")
        with ui_cc2: bw_csm_c = st.slider("摊销前柱宽", 0.1, 1.0, 0.5, key="cc_w")
        with ui_cc3: sz_csm_c = st.slider("摊销前字号", 8, 20, 11, key="cc_s")

        for y, m_id in [(2025, "csm_comp_lat"), (2024, "csm_comp_pre")]:
            an, nt = display_notes(m_id)
            fig_csm_comp = create_csm_composition_chart(df_filtered, selected_cos, y, lab_csm_c, sz_csm_c, bw_csm_c)
            if fig_csm_comp:
                st.plotly_chart(fig_csm_comp, use_container_width=True)
                figs_to_ppt.append((f"CSM占比_{y}", fig_csm_comp, an, nt))
            display_bottom_note(nt)

        ui_cr1, ui_cr2 = st.columns(2)
        with ui_cr1: lab_csm_r = st.toggle("比率折线显示数值", True, key="cr_l")
        with ui_cr2: sz_csm_r = st.slider("比率点大小", 4, 15, 8, key="cr_s")

        an, nt = display_notes("csm_ratio")
        color_mapping = get_company_color_mapping(selected_cos)
        df_csm_sub = df_filtered[df_filtered['字段名'].isin(['CSM摊销', 'CSM期末余额', '新业务CSM（集团口径）'])].pivot_table(index=['公司', '报告年份'], columns='字段名', values='(百万)人民币').reset_index().fillna(0)
        df_csm_sub['摊销比率'] = -df_csm_sub['CSM摊销'] / (df_csm_sub['CSM期末余额'] - df_csm_sub['CSM摊销'])
        df_csm_sub['持续率'] = -df_csm_sub['新业务CSM（集团口径）'] / df_csm_sub['CSM摊销']
        df_csm_sub.replace([np.inf, -np.inf], np.nan, inplace=True)
        df_csm_sub['报告年份'] = df_csm_sub['报告年份'].astype(str).str.replace(".0", "", regex=False) + "YE"
        
        col_l, col_r = st.columns(2)
        with col_l:
            fig_r1 = create_ratio_line_chart_v3(df_csm_sub[['公司', '报告年份', '摊销比率']].rename(columns={'摊销比率': 'value'}), "CSM核心分析 (3/4) - 摊销比率趋势", color_mapping, lab_csm_r, sz_csm_r, 11, ".1%")
            st.plotly_chart(fig_r1, use_container_width=True)
        with col_r:
            fig_r2 = create_ratio_line_chart_v3(df_csm_sub[['公司', '报告年份', '持续率']].rename(columns={'持续率': 'value'}), "CSM核心分析 (3/4) - 持续率趋势", color_mapping, lab_csm_r, sz_csm_r, 11, ".1%")
            st.plotly_chart(fig_r2, use_container_width=True)
        figs_to_ppt.append(("CSM摊销与持续率分析", [fig_r1, fig_r2], an, nt))
        display_bottom_note(nt)

        st.markdown("---")
        ui_ce1, ui_ce2, ui_ce3 = st.columns(3)
        with ui_ce1: lab_csm_e = st.toggle("净资产对比显示标签", True, key="ce_l")
        with ui_ce2: bw_csm_e = st.slider("净资产对比柱宽", 0.1, 1.0, 0.4, key="ce_w")
        with ui_ce3: sz_csm_e = st.slider("净资产对比字号", 8, 20, 11, key="ce_s")

        an, nt = display_notes("csm_equity")
        fig_csm_eq = create_csm_equity_analysis(df_filtered, selected_cos, lab_csm_e, sz_csm_e, bw_csm_e, 16, 20)
        if fig_csm_eq:
            st.plotly_chart(fig_csm_eq, use_container_width=True)
            figs_to_ppt.append(('CSM及净资产对比图', fig_csm_eq, an, nt))
        display_bottom_note(nt)

        # 📍 [预留位置：CSM补充图表]

        # ---------------- Section 5 ----------------
        st.markdown("---")
        st.write("## 5. 新业务价值分析")
        
        ui_nb1, ui_nb2, ui_nb3 = st.columns(3)
        with ui_nb1: lab_nb_c = st.toggle("盈利合同标签", True, key="nbc_l")
        with ui_nb2: psz_nb_c = st.slider("盈利合同涨幅字号", 8, 24, 12, key="nbc_s")
        with ui_nb3: gap_nb_c = st.slider("盈利合同柱间距", 0.1, 0.8, 0.3, key="nbc_g")

        df_nb_csm_raw = df_filtered[(df_filtered['字段名'] == '新业务CSM（集团口径）') & df_filtered['公司'].isin(selected_cos)].drop_duplicates(subset=['公司', '报告年份', '字段名']).copy()
        df_nb_csm_raw.rename(columns={'字段名': '指标名称', '(百万)人民币': 'value'}, inplace=True)
        an, nt = display_notes("nb_csm")
        fig_nb_csm = create_new_biz_csm_chart(df_nb_csm_raw, '新业务CSM（集团口径）', '新业务价值分析 (1/4) - 盈利合同（CSM）对比', lab_nb_c, psz_nb_c, gap_nb_c)
        st.plotly_chart(fig_nb_csm, use_container_width=True)
        figs_to_ppt.append(('新业务盈利合同CSM对比', fig_nb_csm, an, nt))
        display_bottom_note(nt)

        ui_ns1, ui_ns2, ui_ns3 = st.columns(3)
        with ui_ns1: lab_nb_s = st.toggle("价值结构标签", True, key="nbs_l")
        with ui_ns2: bw_nb_s = st.slider("价值结构柱宽", 0.2, 1.0, 0.6, key="nbs_w")
        with ui_ns3: sz_nb_s = st.slider("价值结构字号", 8, 20, 12, key="nbs_s")

        an, nt = display_notes("nb_struct")
        fig_csm_nb, fig_lc, fig_ra = create_new_business_metrics_charts(df_filtered, selected_cos, lab_nb_s, sz_nb_s, bw_nb_s, 14)
        if fig_csm_nb:
            st.plotly_chart(fig_csm_nb, use_container_width=True)
            st.plotly_chart(fig_lc, use_container_width=True)
            st.plotly_chart(fig_ra, use_container_width=True)
            figs_to_ppt.append(('新业务价值结构', [fig_csm_nb, fig_lc, fig_ra], an, nt))
        display_bottom_note(nt)

        # 📍 [预留位置：新业务分析补充图表]

        # ---------------- Section 6 ----------------
        st.markdown("---")
        st.write("## 6. 运营效能分析与六维评价矩阵")
        
        ui_e1, ui_e2, ui_e3 = st.columns(3)
        with ui_e1: lab_exp = st.toggle("显示费用标签", True, key="exp_l")
        with ui_e2: bw_exp = st.slider("费用柱宽", 0.1, 0.8, 0.35, key="exp_w")
        with ui_e3: sz_exp = st.slider("费用字号", 8, 20, 10, key="exp_s")

        an, nt = display_notes("exp_struct")
        fig_exp = create_expense_breakdown_chart(df_filtered, selected_cos, lab_exp, sz_exp, bw_exp, 15)
        if fig_exp:
            st.plotly_chart(fig_exp, use_container_width=True)
            figs_to_ppt.append(('费用结构分析', fig_exp, an, nt))
        display_bottom_note(nt)

        ui_6d1, ui_6d2, ui_6d3 = st.columns(3)
        with ui_6d1: lab_6d = st.toggle("显示六维雷达点标签", True, key="6d_l")
        with ui_6d2: dot_6d = st.slider("圆点大小", 5, 20, 13, key="6d_d")
        with ui_6d3: sz_6d = st.slider("六维标签字号", 8, 20, 10, key="6d_s")

        an, nt = display_notes("six_dim")
        st.write("### 综合评价矩阵 - 财务结果六维表现")
        years_list = sorted([y for y in df_filtered['报告年份'].unique() if str(y).lower() != 'nan'])
        figs_6d = create_six_dimensional_charts(df_filtered, years_list[-1] if years_list else 2025, sz_6d, lab_6d, dot_6d)
        if figs_6d:
            for row in range(2):
                cols = st.columns(3)
                for col_idx in range(3):
                    idx = row * 3 + col_idx
                    if idx < len(figs_6d):
                        with cols[col_idx]: st.plotly_chart(figs_6d[idx], use_container_width=True, config={'displayModeBar': False})
            figs_to_ppt.append(("财务六维分析", figs_6d, an, nt))
        display_bottom_note(nt)

        # --- 表格导出 ---
        an, nt = display_notes("fin_detail")
        st.write("### 📋 附录：保险服务业绩明细表")
        df_l_table = create_financial_report_table(df_filtered, years_list[-1] if years_list else 2025, selected_cos)
        if df_l_table is not None and not df_l_table.empty:
            st.dataframe(format_financial_df(df_l_table), use_container_width=True)
            figs_to_ppt.append(("保险服务业绩明细表", format_financial_df(df_l_table), an, nt))
        display_bottom_note(nt)

        # ==========================================
        # 📂 报告输出区：一键生成 PPT 与 PDF (含自动依赖补救)
        # ==========================================
        st.markdown("---")
        st.write("### 📤报告导出区")
        
        export_col1, export_col2 = st.columns(2)
        
        with export_col1:
            if st.button("🌞 一键生成专业 PPT 报告", use_container_width=True):
                with st.spinner("正在将所有图表与分析排版至 PPT，请稍候..."):
                    try:
                        from pptx import Presentation
                        from pptx.util import Inches, Pt
                        from pptx.enum.text import PP_ALIGN
                        import io
                        
                        prs = Presentation()
                        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9 画幅
                        
                        for title_text, content, an_txt, nt_txt in figs_to_ppt:
                            contents = content if isinstance(content, list) else [content]
                            
                            for single_content in contents:
                                slide = prs.slides.add_slide(prs.slide_layouts[6]) 
                                
                                # 主标题
                                txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12), Inches(0.8))
                                p = txBox.text_frame.paragraphs[0]
                                p.text, p.font.size, p.font.bold = title_text, Pt(24), True
                                
                                # 分析内容 (顶部)
                                if an_txt:
                                    anBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(12), Inches(0.8))
                                    p_an = anBox.text_frame.paragraphs[0]
                                    p_an.text = "💡 业务洞察：\n" + an_txt
                                    p_an.font.size = Pt(13)
                                
                                # 内容放置高度自适应
                                content_top = Inches(1.8) if an_txt else Inches(1.0)
                                
                                # 插入图表
                                if hasattr(single_content, 'write_image'):
                                    img_buf = io.BytesIO()
                                    # Plotly kaleido engine is triggered here
                                    single_content.write_image(img_buf, format="png", width=1200, height=500)
                                    img_buf.seek(0)
                                    slide.shapes.add_picture(img_buf, Inches(1.5), content_top, width=Inches(10))
                                    
                                # 插入表格
                                elif isinstance(single_content, pd.DataFrame):
                                    rows, cols = single_content.shape[0] + 1, single_content.shape[1]
                                    table = slide.shapes.add_table(rows, cols, Inches(0.5), content_top, Inches(12.0), Inches(3.0)).table
                                    for j, col_name in enumerate(single_content.columns):
                                        table.cell(0, j).text = str(col_name)
                                    for i in range(rows - 1):
                                        for j in range(cols):
                                            table.cell(i+1, j).text = str(single_content.iloc[i, j])
                                    for r in range(rows):
                                        for c in range(cols):
                                            for pr in table.cell(r, c).text_frame.paragraphs:
                                                pr.font.size = Pt(10)

                                # 注释内容 (底部)
                                if nt_txt:
                                    ntBox = slide.shapes.add_textbox(Inches(0.5), Inches(6.8), Inches(12), Inches(0.5))
                                    p_nt = ntBox.text_frame.paragraphs[0]
                                    p_nt.text = "📌 附注：" + nt_txt
                                    p_nt.font.size = Pt(11)

                        ppt_buf = io.BytesIO()
                        prs.save(ppt_buf)
                        st.session_state['pptx_output'] = ppt_buf.getvalue()
                        st.success("🎉 PPT 生成成功！")
                        
                    except Exception as e:
                        if "kaleido" in str(e).lower():
                            st.error("⚠️ 导出 PPT 失败！原因：缺少将图表转为图片的必须组件。\n\n**解决方法：** 请在运行此代码的命令提示符（Terminal/CMD）中输入并运行以下命令，安装完毕后再点击按钮即可：\n\n`pip install -U kaleido`")
                        else:
                            st.error(f"生成 PPT 时发生未知错误: {e}")

            if 'pptx_output' in st.session_state:
                st.download_button("📥 下载PPT", st.session_state['pptx_output'], file_name="保险财务分析报告.pptx", use_container_width=True)

        with export_col2:
            st.components.v1.html(
                """
                <div style="display: flex; justify-content: center; align-items: center; height: 100%;">
                    <button onclick="window.parent.print()" 
                            style="width: 100%; padding: 12px; background-color: #F0F2F6; color: #31333F; 
                                   border: 1px solid #D0D4DC; border-radius: 8px; font-size: 16px; cursor: pointer;
                                   font-family: 'Source Sans Pro', sans-serif; font-weight: bold;">
                        🖨️ 网页一键存为高清 PDF (Ctrl+P)
                    </button>
                </div>
                """,
                height=50
            )
            st.caption("提示：这是保存完美排版的最清晰方式。点击后请在打印设置中将【布局】选为“横向”，并在更多设置中勾选【背景图形】。")

# ==================== 页脚 ====================
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 13px; letter-spacing: 1px; margin-top: 50px; padding: 20px; border-top: 1px solid #CBD5E1;">
    Actuarial Data Intelligence · Built for KPMG Actuary Team
</div>
""", unsafe_allow_html=True)
