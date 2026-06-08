import streamlit as st
import requests
from bs4 import BeautifulSoup
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt
from plotly.subplots import make_subplots
import streamlit.components.v1 as components  # 确保引入 components
import streamlit as st
import streamlit.components.v1 as components
import uuid

def show_step_8_content():
    """行业统计分析 - 完全参照 Step 7 的三层架构"""
    import pandas as pd

    # ─────────────────────────────────────────────
    # 工具函数 1：解析 bins 字符串 → float 列表
    # ─────────────────────────────────────────────
    def parse_bins_input(bins_str):
        """
        把 '-5,0,30,50,70,100,500,1000' 解析成 [-5.0, 0.0, 30.0, ..., 1000.0]
        不合法时返回 None
        """
        if bins_str is None:
            return None
        bins_str = str(bins_str).strip()
        if not bins_str:
            return None
        try:
            bins = [float(x.strip()) for x in bins_str.split(",") if str(x).strip() != ""]
            bins = sorted(list(dict.fromkeys(bins)))   # 去重 + 排序
            return bins if len(bins) >= 2 else None
        except Exception:
            return None
    
    # ─────────────────────────────────────────────
    # 工具函数 2：读取 Excel，返回 custom_bins_map
    # ─────────────────────────────────────────────
    def load_custom_bins_excel(uploaded_file):
        df = pd.read_excel(uploaded_file)
        custom_map = {}
        for _, row in df.iterrows():
            m_id = str(row["m_id"]).strip()
            if not m_id or m_id.lower() == "nan":
                continue
            if "enable" in df.columns and str(row["enable"]).strip() in ["0", "False", "false", "否"]:
                continue
            x_bins = parse_bins_input(row["x_bins"])
            y_bins = parse_bins_input(row["y_bins"])
            if x_bins is None and y_bins is None:
                continue
            custom_map[m_id] = {"x_bins_custom": x_bins, "y_bins_custom": y_bins}
        return custom_map
    
    

    # ==========================================
    # 第1层：样式注入与数据准备
    # ==========================================
    
    # 1.1 样式注入
    COMMON_TITLE_FONT = dict(size=18, color="#00338D", family="Microsoft YaHei")

    # ==========================================
    # 1. 样式与前端注入 (紧凑版 CSS & JS)
    # ==========================================

    st.markdown("""
    <style>
    [data-testid="stSidebar"] { background: rgba(255,255,255,0.95) !important; border-right: 1px solid #EAEAEA !important; box-shadow: 2px 0px 15px rgba(0,0,0,0.08) !important; }
    .nav-floating-sign-s8 { position: fixed; left: 0; top: 50%; transform: translateY(-50%); background: rgba(0, 133, 120, 0.85); color: white; padding: 20px 8px; border-radius: 0 12px 12px 0; writing-mode: vertical-rl; text-orientation: mixed; font-size: 22px; font-weight: bold; letter-spacing: 3px; z-index: 9999; cursor: pointer; box-shadow: 3px 3px 12px rgba(0,0,0,0.25); transition: all 0.2s; }
    .nav-floating-sign-s8:hover { background: rgba(0, 133, 120, 1); padding-left: 15px; }
    
    /* 放在 @media print { } 块的外面 */
    .stPlotlyChart {
        width: 100% !important;
        min-width: 0 !important;
    }
    .print-only { display: none !important; }
    /* 🌟 新增：专门给封面封底用的样式类 */
    .cover-page {
        position: relative !important;
        width: 338.67mm !important;
        height: 190.5mm !important;
        margin: 0 !important;
        padding: 0 !important;
        page-break-after: always !important;
        overflow: hidden !important;
        background: transparent !important;
    }
    .cover-page img { width: 100% !important; height: 100% !important; object-fit: cover !important; display: block !important; }
    /* Streamlit网页模式去掉container两侧padding */
    .block-container {
        padding-top:0 !important;
        padding-right:10px !important;
        padding-left:10px !important;
        margin-top:0 !important;
    }
    /* 封面封底强制颜色不被浏览器覆盖 */
    .cover-text {
        forced-color-adjust: none !important;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color: white !important;
        -webkit-text-fill-color: white !important;
    }
    .element-container:first-child{
        margin-top:0 !important;
        padding-top:0 !important;
    }
    @media print {
        .print-only { display: block !important; }
        html,body{
            width:338.67mm!important;
            height:190.5mm!important;
            overflow:hidden!important;
            zoom:100%!important;
        }
        .main .block-container{
            max-width:100%!important;
            padding-top:0!important;
            padding-bottom:0!important;
        }   
        .block-container {
            padding-top: 0rem !important;
        }
        /* ===== 隐藏所有交互元素 ===== */
        .no-print, h1, .nav-floating-sign-s8,
        [data-testid="collapsedControl"], header, footer,
        [data-testid="stHeader"], [data-testid="stSidebar"],
        section[data-testid="stSidebar"],
        [data-testid="stToolbar"], button[kind="secondary"],
        input, .stSlider, [data-testid="stSelectbox"],
        [data-testid="stRadio"], [data-testid="stExpander"],
        .stAlert,
        button[role="tab"],
        div[role="tablist"],
        [data-baseweb="tab-list"],
        hr { display: none !important; }

        /* ===== 页面布局撑满纸张 ===== */
        .page-break-container {
            break-inside: avoid !important;
            margin: 0 !important;              /* 清掉容器 margin，防溢出变空白页 */
            padding: 0 !important;
            padding-bottom: 0mm !important;    /* 只留很小的底部间距 */
        }
        .stApp {
            max-width: 100% !important;
            width: 100% !important;
        }
        
        /* 保留两列并排布局的例外 */
        /* ⭐ PDF打印时强制双列 */
        
        .keep-columns [data-testid="stHorizontalBlock"]{
            display:flex!important;
            flex-wrap:nowrap!important;
            align-items:flex-start!important;
            justify-content:space-between!important;
            gap:0!important;
            width:100%!important;
        }
        .keep-columns [data-testid="stHorizontalBlock"]>div{
            width:49%!important;
            min-width:49%!important;
            max-width:49%!important;
            flex:0 0 49%!important;
            overflow:hidden!important;
            page-break-inside:avoid!important;
            break-inside:avoid!important;
        }

        /* ===== 分页标题 ===== */
        .page-break-title {
            break-before: page !important;      /* 用新语法替代 page-break-before */
            padding-top: 10px !important;       /* 原来 40px，缩小防撑出空白 */
            margin-top: 0 !important;
            text-align: left !important;
        }

        /* ===== 标题 ===== */
        h2 {
            display: block !important;
            text-align: left !important;
            color: #00338D !important;
            font-size: 30px !important;
            font-weight: bold !important;
            border-bottom: 2px solid #00338D !important;
            padding-bottom: 6px !important;
            margin: 14px 0 10px 0 !important;
        }
        h3:not(.no-print) {
            display: block !important;
            text-align: left !important;
            color: #00338D !important;
            font-size: 30px !important;
            font-weight: bold !important;
            margin: 10px 0 8px 0 !important;
            page-break-after: avoid !important;
        }

        /* ===== 图表 ===== */
        .plotly-graph-div,
        .stPlotlyChart {
            width: 100% !important;
            max-width: 100% !important;
            height: auto !important;
            page-break-inside: avoid !important;
            display: block !important;
        }

        /* ===== 表格 ===== */
        div[data-testid="stDataFrame"],
        div[data-testid="stTable"] {
            zoom: 0.65 !important;
            margin: 0 auto 20px auto !important;
            max-width: 100% !important;
            page-break-inside: auto !important;
        }
        div[data-testid="stTable"] tr {
            page-break-inside: avoid !important;
        }

        .element-container {
            page-break-inside: avoid !important;
            width: 100% !important;
        }
        .pdf-page-break {
                break-before: page !important;
                page-break-before: always !important;
                height: 0 !important;
                margin: 0 !important;
                padding: 0 !important;
            }
        
            table {
                page-break-inside: auto !important;
            }
            tr {
                page-break-inside: avoid !important;
                page-break-after: auto !important;
            }
            td, th {
                page-break-inside: avoid !important;
            }
            thead {
                display: table-header-group !important;
            }
        }
    }

    /* 纵向打印微调 */
    @media print and (orientation: portrait) {
        .stPlotlyChart { margin-bottom: 10mm !important; }
    }
    /* 横向打印微调 */
    @media print and (orientation: landscape) {
        .stPlotlyChart { margin-bottom: 6mm !important; }
    }

    .stPlotlyChart, div[data-testid="stDataFrame"] { display: flex !important; justify-content: center !important; }
    .highlight-blue-box { border: 1.5px solid rgba(0,51,141,0.85) !important; border-radius: 12px !important; padding: 10px !important; background: rgba(0,51,141,0.02) !important; box-shadow: 0px 4px 12px rgba(0,51,141,0.12) !important; margin-bottom: 25px !important; }
    </style>
    <div class="nav-floating-sign-s8" id="custom-nav-trigger-s8">展开行业导航栏 </div>
    """, unsafe_allow_html=True)

    components.html("""<script>let t = setInterval(() => { const d = window.parent.document; const b = d.getElementById("custom-nav-trigger-s8"); const c = d.querySelector('[data-testid="collapsedControl"]') || d.querySelector('button[kind="header"]'); if(b && c) { b.onclick = () => c.click(); clearInterval(t); } }, 500);</script>""", height=0, width=0)
    
    # 1.2 获取数据
    if 'integrated_data' not in st.session_state or st.session_state['integrated_data'] is None:
        st.warning("⚠️ 请先在 Step 6 完成数据集成。")
        return
    
    df_raw = st.session_state['integrated_data'].copy()
    
    # 1.3 提取年份
    valid_years = sorted([
        int(y) for y in df_raw['报告年份'].dropna().astype(str).str.replace(".0", "", regex=False).unique()
        if y.isdigit()
    ])
    if len(valid_years) < 2:
        st.warning("⚠️ 年份数量不足，至少需要两个年份。")
        return
    
    latest_year = valid_years[-1]
    prev_year = valid_years[-2]
    year_list = [prev_year, latest_year]
    
    st.markdown("### 📈 行业统计分析")

    # ==========================================
    # 2. 注释表加载区（适配 Step 7 格式）
    # ==========================================
    notes_dict_8, ordered_modules, first_levels = {}, [], []
    df_notes = None
    
    def generate_custom_analysis(m_id, df, cos, cy, py):
            """根据m_id和数据自动生成行业分析内容，返回字符串"""
            import numpy as np
     
            # ── 工具函数1：取某公司类型某字段某年的所有公司均值 ──
            def gv(field, year, ct):
                yr = str(int(year)) if year == int(year) else str(year)
                mask = (
                    (df['公司类型'] == ct) &
                    (df['报告年份'].astype(str).str.replace('.0', '', regex=False) == yr) &
                    (df['字段名'] == field)
                )
                vals = df[mask]['(百万)人民币'].dropna()
                return vals.mean() if not vals.empty else None
     
            # ── 工具函数2：涨跌型（各公司类型均值同比）──
            def rise_fall(field, year_c, year_p, unit="亿元", scale=100,
                          actual_field=None):
                f = actual_field or field
                rises, falls = [], []
                for ct in cos:
                    c = gv(f, year_c, ct)
                    p = gv(f, year_p, ct)
                    if c is None or p is None or p == 0:
                        continue
                    chg  = (c - p) / abs(p)
                    val_c = c / scale
                    (rises if chg >= 0 else falls).append((ct, chg, val_c))
                parts = []
                if rises:
                    rises.sort(key=lambda x: -x[1])
                    parts.append(
                        f"{len(rises)}类机构均值上升，其中{rises[0][0]}增幅最大"
                        f"（+{rises[0][1]:.1%}，均值约{rises[0][2]:.0f}{unit}）"
                    )
                if falls:
                    falls.sort(key=lambda x: x[1])
                    parts.append(
                        f"{len(falls)}类机构均值下降，其中{falls[0][0]}降幅最大"
                        f"（{falls[0][1]:.1%}，均值约{falls[0][2]:.0f}{unit}）"
                    )
                return "；".join(parts) + "。" if parts else ""
     
            # ── 工具函数3：构成型 ──
            def composition_by_type(fields_display, year):
                lines = []
                for ct in cos:
                    totals = {f: abs(gv(f, year, ct) or 0) for f in fields_display}
                    grand  = sum(totals.values())
                    if grand == 0:
                        continue
                    main_f   = max(totals, key=totals.get)
                    main_pct = totals[main_f] / grand * 100
                    lines.append(
                        f"{ct}以{fields_display[main_f]}为主（占比约{main_pct:.1f}%）"
                    )
                return "；".join(lines) + "。" if lines else ""
     
            # ── 工具函数4：计算字段的涨跌描述（两个字段相减）──
            def calc_two_field_diff(field_a, field_b, year_c, year_p,
                                    unit="亿元", scale=100):
                """net = field_a - field_b，对每个公司类型计算均值并同比"""
                lines = []
                for ct in cos:
                    a_c = gv(field_a, year_c, ct)
                    b_c = gv(field_b, year_c, ct)
                    a_p = gv(field_a, year_p, ct)
                    b_p = gv(field_b, year_p, ct)
                    if a_c is None:
                        continue
                    net_c = ((a_c or 0) - (b_c or 0)) / scale
                    net_p = ((a_p or 0) - (b_p or 0)) / scale \
                            if a_p is not None else None
                    if net_p and net_p != 0:
                        chg       = (net_c - net_p) / abs(net_p)
                        direction = "升" if chg >= 0 else "降"
                        lines.append(
                            f"{ct}均值约{net_c:.0f}{unit}（同比{direction}{abs(chg):.1%}）"
                        )
                    else:
                        lines.append(f"{ct}均值约{net_c:.0f}{unit}")
                return "；".join(lines) + "。" if lines else ""
     
            # ── 工具函数5：分布图通用描述（各类型参与公司数量）──
            def dist_count_summary(year):
                lines = []
                for ct in cos:
                    yr   = str(int(year)) if year == int(year) else str(year)
                    mask = (
                        (df['公司类型'] == ct) &
                        (df['报告年份'].astype(str)
                         .str.replace('.0', '', regex=False) == yr)
                    )
                    n = df[mask]['公司'].nunique()
                    if n > 0:
                        lines.append(f"{ct}{n}家")
                return "样本：" + "、".join(lines) + "。" if lines else ""
     
            try:
                # ============================================================
                # 散点图 ── SCATTER_AXIS_META 的 15 个 m_id
                # ============================================================
     
                # 1. 保险服务收入
                if m_id == "industry_inc_total":
                    return rise_fall("保险服务收入", cy, py,
                                     actual_field="保险服务收入合计")
     
                # 2. 净投资回报
                if m_id == "industry_inv_return":
                    return rise_fall("净投资回报", cy, py,
                                     actual_field="净投资回报")
     
                # 3. 保险服务费用
                if m_id == "industry_exp_total":
                    return rise_fall("保险服务费用", cy, py,
                                     actual_field="保险服务费用合计")
     
                # 4. 承保财务净损益（= 承保财务损益 - 分出再保险财务损益）
                if m_id == "industry_uw_profit":
                    return calc_two_field_diff(
                        "承保财务损益", "分出再保险财务损益", cy, py
                    )
     
                # 5. 保险服务业绩（= 保险服务收入合计 - 保险服务费用合计）
                if m_id == "industry_perf_total":
                    return calc_two_field_diff(
                        "保险服务收入合计", "保险服务费用合计", cy, py
                    )
     
                # 6. 投资利润（= 净投资回报 - 承保财务净损益）
                if m_id == "industry_inv_profit":
                    lines = []
                    for ct in cos:
                        nr_c = gv("净投资回报",        cy, ct)
                        uw_c = gv("承保财务损益",       cy, ct)
                        re_c = gv("分出再保险财务损益", cy, ct)
                        nr_p = gv("净投资回报",        py, ct)
                        uw_p = gv("承保财务损益",       py, ct)
                        re_p = gv("分出再保险财务损益", py, ct)
                        if nr_c is None:
                            continue
                        net_uw_c = (uw_c or 0) - (re_c or 0)
                        net_uw_p = (uw_p or 0) - (re_p or 0)
                        inv_c    = ((nr_c or 0) - net_uw_c) / 100
                        inv_p    = ((nr_p or 0) - net_uw_p) / 100 \
                                   if nr_p is not None else None
                        if inv_p and inv_p != 0:
                            chg       = (inv_c - inv_p) / abs(inv_p)
                            direction = "升" if chg >= 0 else "降"
                            lines.append(
                                f"{ct}均值约{inv_c:.0f}亿元（同比{direction}{abs(chg):.1%}）"
                            )
                        else:
                            lines.append(f"{ct}均值约{inv_c:.0f}亿元")
                    return "；".join(lines) + "。" if lines else ""
     
                # 7. 净利润
                if m_id == "industry_net_profit":
                    return rise_fall("净利润", cy, py, actual_field="净利润")
     
                # 8. 税前利润
                if m_id == "industry_tax_profit":
                    return rise_fall("税前利润", cy, py, actual_field="税前利润总额")
     
                # 9. 其他综合收益
                if m_id == "industry_oci_profit":
                    return rise_fall("其他综合收益", cy, py, actual_field="其他综合收益")
     
                # 10. 综合收益总额
                if m_id == "industry_total_profit":
                    return rise_fall("综合收益总额", cy, py, actual_field="综合收益总额")
     
                # 11. CSM余额
                if m_id == "industry_csm_bal":
                    return rise_fall("CSM余额", cy, py, actual_field="CSM期末余额")
     
                # 12. 净资产
                if m_id == "industry_equity_trend":
                    return rise_fall("净资产", cy, py, actual_field="期末股东权益")
     
                # 13. CSM/净资产占比（= CSM期末余额 × 0.75 / 期末股东权益 × 100）
                if m_id == "industry_csm_equity":
                    lines = []
                    for ct in cos:
                        csm_c = gv("CSM期末余额",  cy, ct)
                        eq_c  = gv("期末股东权益", cy, ct)
                        csm_p = gv("CSM期末余额",  py, ct)
                        eq_p  = gv("期末股东权益", py, ct)
                        if csm_c is None or not eq_c:
                            continue
                        ratio_c = csm_c * 0.75 / eq_c * 100
                        if csm_p and eq_p:
                            ratio_p = csm_p * 0.75 / eq_p * 100
                            chg     = ratio_c - ratio_p   # 百分点变化
                            sign    = "+" if chg >= 0 else ""
                            lines.append(
                                f"{ct}均值约{ratio_c:.1f}%（同比{sign}{chg:.1f}个百分点）"
                            )
                        else:
                            lines.append(f"{ct}均值约{ratio_c:.1f}%")
                    return "；".join(lines) + "。" if lines else ""
     
                # 14. 新业务CSM
                if m_id == "industry_nb_csm":
                    return rise_fall("新业务CSM", cy, py,
                                     actual_field="新业务CSM（集团口径）")
     
                # 15. 新业务LC亏损率（= 新业务亏损合同 / 新业务未来现金流入现值（亏损）× 100）
                if m_id == "industry_lc_loss_ratio":
                    lines = []
                    for ct in cos:
                        lc_c  = gv("新业务亏损合同（LC）——非PAA",  cy, ct)
                        pv_c  = gv("新业务未来现金流入现值（亏损）", cy, ct)
                        lc_p  = gv("新业务亏损合同（LC）——非PAA",  py, ct)
                        pv_p  = gv("新业务未来现金流入现值（亏损）", py, ct)
                        if lc_c is None or not pv_c:
                            continue
                        ratio_c = lc_c / pv_c * 100
                        if lc_p and pv_p:
                            ratio_p = lc_p / pv_p * 100
                            chg     = ratio_c - ratio_p
                            sign    = "+" if chg >= 0 else ""
                            lines.append(
                                f"{ct}均值约{ratio_c:.1f}%（同比{sign}{chg:.1f}个百分点）"
                            )
                        else:
                            lines.append(f"{ct}均值约{ratio_c:.1f}%")
                    return "；".join(lines) + "。" if lines else ""
     
                # ============================================================
                # 堆叠分布图 ── STACK_DIST_META 的 4 个 m_id
                # ============================================================
     
                # 16. CSM/BEL占比分布（= CSM期末余额 / BEL期末余额 × 100）
                if m_id == "industry_csm_ratio":
                    lines = []
                    for ct in cos:
                        csm_c = gv("CSM期末余额", cy, ct)
                        bel_c = gv("BEL期末余额", cy, ct)
                        if csm_c is None or not bel_c:
                            continue
                        ratio = csm_c / bel_c * 100
                        lines.append(f"{ct}行业均值约{ratio:.1f}%")
                    count_str = dist_count_summary(cy)
                    ratio_str = "；".join(lines) + "。" if lines else ""
                    return count_str + ratio_str
     
                # 17. CSM持续率分布（= 新业务CSM / CSM摊销 × 100）
                if m_id == "industry_csm_continuity_ratio":
                    lines = []
                    for ct in cos:
                        nb_c    = gv("新业务CSM（集团口径）", cy, ct)
                        amort_c = gv("CSM摊销",               cy, ct)
                        if nb_c is None or not amort_c:
                            continue
                        ratio = -nb_c / amort_c * 100
                        lines.append(f"{ct}行业均值约{ratio:.1f}%")
                    count_str = dist_count_summary(cy)
                    ratio_str = "；".join(lines) + "。" if lines else ""
                    return count_str + ratio_str
     
                # 18. CSM摊销比率分布（= CSM摊销 / (期末CSM - CSM摊销) × 100）
                if m_id == "industry_csm_amort_ratio":
                    lines = []
                    for ct in cos:
                        amort_c = gv("CSM摊销",     cy, ct)
                        end_c   = gv("CSM期末余额", cy, ct)
                        if amort_c is None or end_c is None:
                            continue
                        denom = end_c - amort_c
                        if denom == 0:
                            continue
                        ratio = -amort_c / denom * 100
                        lines.append(f"{ct}行业均值约{ratio:.1f}%")
                    count_str = dist_count_summary(cy)
                    ratio_str = "；".join(lines) + "。" if lines else ""
                    return count_str + ratio_str
     
                # 19. 新业务IFRS利润率分布（= (新业务CSM - LC) / 未来现金流入现值 × 100）
                if m_id == "industry_nb_ifrs_margin":
                    lines = []
                    for ct in cos:
                        nb_c  = gv("新业务CSM（集团口径）",       cy, ct)
                        lc_c  = gv("新业务亏损合同（LC）——非PAA", cy, ct)
                        pv_c  = gv("新业务未来现金流入现值",       cy, ct)
                        if nb_c is None or not pv_c:
                            continue
                        ratio = ((nb_c or 0) - (lc_c or 0)) / pv_c * 100
                        lines.append(f"{ct}行业均值约{ratio:.1f}%")
                    count_str = dist_count_summary(cy)
                    ratio_str = "；".join(lines) + "。" if lines else ""
                    return count_str + ratio_str
     
                # ============================================================
                # 构成类图表（13 个 m_id）
                # ============================================================
     
                if m_id == "industry_comp_1":
                    return composition_by_type({
                        "采用保费分配法计量的保险合同保险服务收入": "PAA合同收入",
                        "未采用保费分配法计量的保险合同保险服务收入": "非PAA合同收入",
                    }, cy)
     
                if m_id == "industry_comp_2":
                    return composition_by_type({
                        "合同服务边际的摊销": "合同服务边际释放",
                        "非金融风险调整的变动": "非金融风险调整变动",
                        "预计当期发生的保险服务费用": "预期保险服务费用",
                        "保险获取现金流的摊销（保险服务收入）": "保险获取现金流摊销",
                    }, cy)
     
                if m_id == "industry_exp_1":
                    return composition_by_type({
                        "保险获取现金流的摊销（保险服务费用）": "保险获取现金流摊销",
                        "亏损部分的确认及转回": "亏损部分确认及转回",
                        "当期发生的赔款及其他相关费用": "当期赔款及费用",
                        "已发生赔款负债相关的履约现金流量变动": "已发生赔款负债变动",
                    }, cy)
     
                if m_id == "industry_exp_struct":
                    lines = []
                    for ct in cos:
                        vals = {
                            "获取费用": abs(gv("获取费用",  cy, ct) or 0),
                            "维持费用": abs(gv("维持费用",  cy, ct) or 0),
                            "非履约费用": abs(gv("非履约费用", cy, ct) or 0),
                        }
                        grand = sum(vals.values())
                        if grand == 0:
                            continue
                        main_f = max(vals, key=vals.get)
                        lines.append(
                            f"{ct}以{main_f}为主（占比约{vals[main_f]/grand*100:.1f}%）"
                        )
                    return "；".join(lines) + "。" if lines else ""
     
                if m_id == "industry_prof_2025":
                    csm_lines = [
                        f"{ct}约{gv('合同服务边际的释放', cy, ct):.1f}%"
                        for ct in cos
                        if gv("合同服务边际的释放", cy, ct) is not None
                    ]
                    csm_str  = "合同服务边际的释放均值：" + "、".join(csm_lines) + "。" \
                               if csm_lines else ""
                    comp_str = composition_by_type({
                        "亏损部分的确认及转回": "亏损部分确认及转回",
                        "合同服务边际的释放":   "合同服务边际释放",
                        "非金融风险调整的变动": "非金融风险调整变动",
                        "营运偏差及其他":       "营运偏差及其他",
                        "保费分配法业务净损益": "保费分配法净损益",
                        "再保净损益":           "再保净损益",
                    }, cy)
                    return (comp_str + csm_str).strip()
     
                if m_id == "industry_prof_2024":
                    csm_lines = [
                        f"{ct}约{gv('合同服务边际的释放', py, ct):.1f}%"
                        for ct in cos
                        if gv("合同服务边际的释放", py, ct) is not None
                    ]
                    csm_str  = "合同服务边际的释放均值：" + "、".join(csm_lines) + "。" \
                               if csm_lines else ""
                    comp_str = composition_by_type({
                        "亏损部分的确认及转回": "亏损部分确认及转回",
                        "合同服务边际的释放":   "合同服务边际释放",
                        "非金融风险调整的变动": "非金融风险调整变动",
                        "营运偏差及其他":       "营运偏差及其他",
                        "保费分配法业务净损益": "保费分配法净损益",
                        "再保净损益":           "再保净损益",
                    }, py)
                    return (comp_str + csm_str).strip()
     
                if m_id == "industry_prof_mix":
                    lines = []
                    for ct in cos:
                        ins = gv("保险利润", cy, ct) or 0
                        inv = gv("投资利润", cy, ct) or 0
                        tot = ins + inv
                        if tot == 0:
                            continue
                        dominant = "保险利润" if ins / tot >= 0.5 else "投资利润"
                        lines.append(
                            f"{ct}以{dominant}为主（保险利润占比{ins/tot:.0%}）"
                        )
                    return "；".join(lines) + "。" if lines else ""
     
                if m_id == "industry_oci_deep":
                    lines = []
                    for ct in cos:
                        liab = gv("可转损益的负债OCI",  cy, ct) or 0
                        bond = gv("FVOCI债券公允价值", cy, ct) or 0
                        if liab == 0 and bond == 0:
                            continue
                        dominant = "负债端OCI" if abs(liab) >= abs(bond) \
                                   else "FVOCI债券公允价值变动"
                        lines.append(f"{ct}主要驱动为{dominant}")
                    return "；".join(lines) + "。" if lines else ""
     
                if m_id == "industry_oci_year_lat":
                    return composition_by_type({
                        "净利润":         "净利润",
                        "可转损益OCI合计":  "可转损益OCI",
                        "不可转损益OCI合计": "不可转损益OCI",
                    }, cy)
     
                if m_id == "industry_oci_year_pre":
                    return composition_by_type({
                        "净利润":         "净利润",
                        "可转损益OCI合计":  "可转损益OCI",
                        "不可转损益OCI合计": "不可转损益OCI",
                    }, py)
     
                if m_id == "industry_asset_struct":
                    return composition_by_type({
                        "AC":      "AC（债权投资）",
                        "FVOCI":   "FVOCI（其他债权投资）",
                        "FVTPL":   "FVTPL（交易性金融资产）",
                        "指定FVOCI": "指定FVOCI（其他权益工具）",
                    }, cy)
     
                if m_id == "industry_csm_comp_lat":
                    return composition_by_type({
                        "新业务CSM（集团口径）": "新业务CSM",
                        "CSM计息":             "CSM计息",
                        "CSM调整":             "CSM调整",
                    }, cy)
     
                if m_id == "industry_csm_comp_pre":
                    return composition_by_type({
                        "新业务CSM（集团口径）": "新业务CSM",
                        "CSM计息":             "CSM计息",
                        "CSM调整":             "CSM调整",
                    }, py)
     
            except Exception as e:
                return f"（自动生成失败：{e}）"
     
            return ""   # 无匹配的 m_id 返回空字符串，不报错
    

    with st.expander("📥 行业分析注释输入", expanded=False):
        # 1. 上传自定义坐标轴刻度表 (保持不变)
        uploaded_bins_file = st.file_uploader(
            "📊 上传自定义坐标轴刻度表（Excel）",
            type=["xlsx", "xls"], key="custom_bins_uploader",
            help="表头必须包含：m_id, x_bins, y_bins。可选列：enable（0=禁用）、note"
        )
        if uploaded_bins_file is not None:
            try:
                custom_bins_map = load_custom_bins_excel(uploaded_bins_file)
                st.session_state["custom_bins_map"] = custom_bins_map
                st.success(f"✅ 已加载 {len(custom_bins_map)} 个图的自定义刻度配置")
            except Exception as e:
                st.error(f"❌ 刻度表读取失败：{e}")
    
        if st.session_state.get("custom_bins_map"):
            if st.button("清除自定义刻度配置", key="clear_custom_bins"):
                del st.session_state["custom_bins_map"]
                st.success("已清除自定义刻度配置，恢复默认")
    
        # ==========================================
        # 2. 注释表来源选择
        # 🌟 核心修改 1：把 value=False 改成 value=True，让它默认就是打开状态！
        # ==========================================
        use_default = st.toggle("使用默认注释表", value=True, key="step8_use_default")
    
        df_notes = None

        if use_default:
            try:
                with st.spinner("🚀 正在从云端加载默认注释表，请稍候..."):
                    # 🌟 核心修改 2：替换为 raw.githubusercontent.com 的原生读取链接
                    default_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/RD-%E5%9B%BE%E7%89%87%E5%86%85%E5%AE%B9%E5%88%86%E6%9E%90%E5%92%8C%E6%B3%A8%E9%87%8A%E6%A8%A1%E6%9D%BF-step8.xlsx"
                    df_notes = pd.read_excel(default_url)
                st.success("✅ 内置默认注释表从云端加载成功！")
            except Exception as e:
                st.error(f"❌ 云端下载失败，报错原因: {e}")
                # 贴心加上排错提示
                st.info("💡 排错指南：\n1. 请确认你的 GitHub 仓库是 Public（公开）的，如果是 Private 则代码无法直接读取。\n2. 请确认文件已经成功 Push 到 GitHub，且没有拼写错误。")
        else:
            notes_file = st.file_uploader("上传 Excel 分析注释表", type=['xlsx', 'xls'], key="step8_notes")
            if notes_file:
                try:
                    df_notes = pd.read_excel(notes_file)
                    st.success("✅ 自定义注释表上传成功！")
                except Exception as e:
                    st.error(f"❌ 上传文件解析失败: {e}")

        # 💡 友情提示：记得在后续代码中，使用 df_notes 之前，先判断一下 if df_notes is not None:
    
        if df_notes is not None:
            # ==========================================
            # 🌟 核心防崩溃修复：强制把要塞文本的列变成 object 类型
            # ==========================================
            if '分析内容-自定义' in df_notes.columns:
                df_notes['分析内容-自定义'] = df_notes['分析内容-自定义'].astype(object)
            if '分析内容-默认' in df_notes.columns:
                df_notes['分析内容-默认'] = df_notes['分析内容-默认'].astype(object)

            if '模块ID' in df_notes.columns and '分析内容-自定义' in df_notes.columns:
                for idx, row in df_notes.iterrows():
                    mid = str(row.get('模块ID', '')).strip()
                    if not mid or mid == 'nan':
                        continue
                    raw_val = row.get('分析内容-自定义')
                    # 修复：用 pd.isna 兼容 float NaN，不依赖字符串比较
                    if pd.isna(raw_val) or str(raw_val).strip() in ('', 'nan', 'None'):
                        generated = generate_custom_analysis(
                            mid, df_raw,
                            st.session_state.get('step8_selected_types', []),
                            latest_year, prev_year
                        )
                        if generated:
                            # 穿上防护服后，这里塞入字符串就绝对不会再报 TypeError 了！
                            df_notes.at[idx, '分析内容-自定义'] = generated
            
            # 1. 清洗所有列（去除前后空格）
            for col in df_notes.columns:
                df_notes[col] = df_notes[col].astype(str).str.strip()
            
            # 2. 确保必要的列存在（如果不存在则创建空列）
            required_cols = ['模块ID', '一级分类', '二级分类', '对应图表名称', '分析内容-默认', '分析内容-自定义', '注释内容']
            for col in required_cols:
                if col not in df_notes.columns:
                    df_notes[col] = ''
            
            # 3. 清洗空值（将 nan 转为空字符串）
            for col in ['一级分类', '二级分类', '对应图表名称', '模块ID']:
                if col in df_notes.columns:
                    df_notes[col] = df_notes[col].replace(['nan', 'NaN', 'NAN', 'None'], '')
            
            # 4. 二级分类为空时，强制填入"全部"
            if '二级分类' in df_notes.columns:
                df_notes['二级分类'] = df_notes['二级分类'].apply(
                    lambda x: "全部" if str(x).strip() == "" else str(x).strip()
                )
            
            # 5. 构建 notes_dict_8 字典（合并两个分析字段）
            for _, r in df_notes.iterrows():
                m_id = str(r.get('模块ID', '')).strip()
                if not m_id:
                    continue
                
                # 合并分析内容：优先使用自定义，若自定义为空则用默认
                analysis_default = str(r.get('分析内容-默认', '')).strip()
                analysis_custom = str(r.get('分析内容-自定义', '')).strip()
                
                # 最终显示的分析内容（优先自定义）
                final_analysis = analysis_custom if analysis_custom else analysis_default
                
                notes_dict_8[m_id] = {
                    'title': str(r.get('对应图表名称', '')).strip(),
                    'analysis': final_analysis,           # 合并后的分析内容
                    'analysis_default': analysis_default,  # 保留默认供后续使用
                    'analysis_custom': analysis_custom,    # 保留自定义供后续使用
                    'note': str(r.get('注释内容', '')).strip(),
                    '一级分类': str(r.get('一级分类', '')).strip(),
                    '二级分类': str(r.get('二级分类', '')).strip(),
                    'remark': str(r.get('话术', '')).strip() if '话术' in df_notes.columns else ''
                }
                
                if m_id not in ordered_modules:
                    ordered_modules.append(m_id)
            
            # 6. 提取一级分类列表
            first_levels = [x for x in df_notes['一级分类'].unique() if x and x != '']
            
            # 7. 存入 session_state
            st.session_state['step8_notes_dict'] = notes_dict_8
            st.session_state['step8_ordered_modules'] = ordered_modules
            st.session_state['step8_df_notes'] = df_notes.copy()
    
    # 如果未加载，则尝试从session_state取
    if df_notes is None and 'step8_df_notes' in st.session_state:
        df_notes = st.session_state['step8_df_notes'].copy()
        notes_dict_8 = st.session_state.get('step8_notes_dict', {})
        ordered_modules = st.session_state.get('step8_ordered_modules', [])
        first_levels = [x for x in df_notes['一级分类'].unique() if x and x != ''] if df_notes is not None else []
        
    # ==========================================
    # 3. 侧边栏导航
    # ==========================================
    
    print_mode = False
    active_m_id = None
    active_chart_name = None
    
    with st.sidebar:
        st.markdown("<h3 style='color: #00338D; font-size: 18px;'>行业分析导航</h3>", unsafe_allow_html=True)
    
        if first_levels:
            main_nav = st.radio("📁 一级模块", first_levels + ["🖨️ 一键显示全部 (打印/导出)"], key="step8_main")
    
            if main_nav == "🖨️ 一键显示全部 (打印/导出)":
                print_mode = True
                st.info("💡 点击下方按钮导出 PDF")
                components.html(
                    """<button onclick="window.parent.print()" style="width:100%; padding:12px; background:#00338D; color:white; border:none; border-radius:6px; cursor:pointer;">立即导出 PDF 报告</button>""",
                    height=60
                )
            else:
                df_sub1 = df_notes[df_notes['一级分类'] == main_nav]
                sec_levels = [x for x in df_sub1['二级分类'].unique() if x and x != '']
    
                if len(sec_levels) == 0:
                    charts = [x for x in df_sub1['对应图表名称'].unique() if x and x != '']
                    chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                    row = df_sub1[df_sub1['对应图表名称'] == chart_nav].iloc[0]
                    active_m_id = row['模块ID']
                    active_chart_name = row['对应图表名称']
                else:
                    sub_nav = st.radio("📂 二级模块", ["全部"] + sec_levels, key="step8_sub")
                    if sub_nav != "全部":
                        df_sub2 = df_sub1[df_sub1['二级分类'] == sub_nav]
                        charts = [x for x in df_sub2['对应图表名称'].unique() if x and x != '']
                        chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                        row = df_sub2[df_sub2['对应图表名称'] == chart_nav].iloc[0]
                        active_m_id = row['模块ID']
                        active_chart_name = row['对应图表名称']
                    else:
                        charts = [x for x in df_sub1['对应图表名称'].unique() if x and x != '']
                        chart_nav = st.radio("📊 具体图表", charts, key="step8_chart")
                        row = df_sub1[df_sub1['对应图表名称'] == chart_nav].iloc[0]
                        active_m_id = row['模块ID']
                        active_chart_name = row['对应图表名称']
        else:
            st.warning("⚠️ 请先上传包含层级信息的注释表")
            return
        
    # ==========================================
    # 行业分析配置（公司类型选择等）
    # ==========================================
    
    # 🌟 固定的公司类型显示顺序
    COMPANY_TYPE_ORDER = ["头部", "银行系", "外资系", "养老健康", "小型"]
    
    def sort_company_types(type_list):
        return sorted(type_list, key=lambda x: COMPANY_TYPE_ORDER.index(x) if x in COMPANY_TYPE_ORDER else 999)
    
    # 公司类型色卡
    KPMG_CATEGORIES = globals().get('KPMG_CATEGORIES', {
        "Primary Colors": {
            "KPMG Blue": "#00338D", "Cobalt Blue": "#1E49E2",
            "Light Blue": "#ACEAFF", "Pacific Blue": "#00B8F5", 
            "Purple": "#7213EA", "Pink": "#FD349C"
        },
        "Traffic Light": {"Red": "#ED2124", "Amber": "#F1C44D", "Positive Green": "#269924"}
    })
    primary = KPMG_CATEGORIES["Primary Colors"]
    traffic = KPMG_CATEGORIES["Traffic Light"]
    COMPANY_TYPE_COLORS = {
        "头部": primary["KPMG Blue"],
        "银行系": primary["Pacific Blue"],
        "外资系": traffic["Positive Green"],
        "养老健康": traffic["Amber"],
        "小型": primary["Purple"]
    }
    
    def get_company_color(ct):
        return COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
    
    # 1.6 全局配置
    with st.expander("⚙️ 行业分析配置", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            all_types = sorted([str(x) for x in df_raw['公司类型'].dropna().unique()])
            all_types_sorted = [t for t in COMPANY_TYPE_ORDER if t in all_types]
            for t in all_types:
                if t not in all_types_sorted:
                    all_types_sorted.append(t)
            
            selected_types = st.multiselect(
                "🏢 选择公司类型（可多选）",
                all_types_sorted,
                default=[t for t in COMPANY_TYPE_ORDER if t in all_types_sorted]
            )
            st.session_state['step8_selected_types'] = selected_types  # 👈 存入session
            
        with c2:
            year_1 = st.selectbox("年度1", valid_years, index=len(valid_years)-2)
            year_2 = st.selectbox("年度2", valid_years, index=len(valid_years)-1)
            year_list = [int(year_1), int(year_2)]
    
    if not selected_types:
        st.warning("请至少选择一个公司类型")
        return
    
    # ==========================================
    # 第2层：计算函数库（只负责数据计算，不涉及绘图）
    # ==========================================


    
    def parse_bins_input(bins_str):
        """解析区间输入字符串，返回排序后的列表或 None"""
        if not bins_str or str(bins_str).strip() == "":
            return None
        try:
            bins = sorted(set([float(x.strip()) for x in str(bins_str).split(",") if str(x).strip()]))
            if len(bins) < 2:
                return None
            return bins
        except:
            return None
    
    def calc_scatter_data(df, selected_types, latest_year, prev_year, field_name, display_name):
        """计算散点图数据（支持计算字段）"""
        import pandas as pd
        import numpy as np
        
        df_current = df[df['报告年份'].astype(str) == str(latest_year)].copy()
        df_prev = df[df['报告年份'].astype(str) == str(prev_year)].copy()
        
        # 获取所有公司的类型映射
        company_type_map = {}
        for co in df_current['公司'].unique():
            co_type_rows = df_current[df_current['公司'] == co]['公司类型']
            if not co_type_rows.empty:
                company_type_map[co] = co_type_rows.iloc[0]
        
        scatter_data = []
        
        for co in company_type_map.keys():
            co_type = company_type_map[co]
            if co_type not in selected_types:
                continue
            
            # ========== 根据 field_name 计算当前值和前一年的值 ==========
            current_val = None
            prev_val = None
            
            # 1. 承保财务净损益 = 承保财务损益 - 分出再保险财务损益
            if field_name == "承保财务净损益":
                inv_uw_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '承保财务损益')]['(百万)人民币'].sum()
                inv_re_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '分出再保险财务损益')]['(百万)人民币'].sum()
                current_val = (inv_uw_curr - inv_re_curr) / 100
                
                inv_uw_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '承保财务损益')]['(百万)人民币'].sum()
                inv_re_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '分出再保险财务损益')]['(百万)人民币'].sum()
                prev_val = (inv_uw_prev - inv_re_prev) / 100
            
            # 2. 投资利润 = 净投资回报 - 承保财务净损益
            elif field_name == "投资利润":
                # 先计算承保财务净损益
                uw_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '承保财务损益')]['(百万)人民币'].sum() - \
                          df_current[(df_current['公司'] == co) & (df_current['字段名'] == '分出再保险财务损益')]['(百万)人民币'].sum()
                nr_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '净投资回报')]['(百万)人民币'].sum()
                current_val = (nr_curr - uw_curr) / 100
                
                uw_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '承保财务损益')]['(百万)人民币'].sum() - \
                          df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '分出再保险财务损益')]['(百万)人民币'].sum()
                nr_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '净投资回报')]['(百万)人民币'].sum()
                prev_val = (nr_prev - uw_prev) / 100
            
            # 3. 保险服务业绩 = 保险服务收入合计 - 保险服务费用合计
            elif field_name == "保险服务业绩":
                inc_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '保险服务收入合计')]['(百万)人民币'].sum()
                exp_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '保险服务费用合计')]['(百万)人民币'].sum()
                current_val = (inc_curr - exp_curr) / 100
                
                inc_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '保险服务收入合计')]['(百万)人民币'].sum()
                exp_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '保险服务费用合计')]['(百万)人民币'].sum()
                prev_val = (inc_prev - exp_prev) / 100
            
            # 4. CSM/净资产占比 = (CSM期末余额 * 0.75) / 期末股东权益 * 100
            elif field_name == "CSM/净资产占比":
                csm_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == 'CSM期末余额')]['(百万)人民币'].sum()
                eq_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '期末股东权益')]['(百万)人民币'].sum()
                current_val = (csm_curr * 0.75 / eq_curr * 100) if eq_curr != 0 else 0
                
                csm_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == 'CSM期末余额')]['(百万)人民币'].sum()
                eq_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '期末股东权益')]['(百万)人民币'].sum()
                prev_val = (csm_prev * 0.75 / eq_prev * 100) if eq_prev != 0 else 0
            
            # 5. LC亏损率 = 新业务亏损合同 / 新业务未来现金流入现值（亏损） * 100
            elif field_name == "LC亏损率":
                lc_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务亏损合同（LC）——非PAA')]['(百万)人民币'].sum()
                pv_loss_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务未来现金流入现值（亏损）')]['(百万)人民币'].sum()
                current_val = (lc_curr / pv_loss_curr * 100) if pv_loss_curr != 0 else np.nan
                
                lc_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务亏损合同（LC）——非PAA')]['(百万)人民币'].sum()
                pv_loss_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务未来现金流入现值（亏损）')]['(百万)人民币'].sum()
                prev_val = (lc_prev / pv_loss_prev * 100) if pv_loss_prev != 0 else np.nan
            
            # 6. 新业务CSM利润率 = 新业务CSM / 新业务未来现金流入现值（盈利） * 100
            elif field_name == "新业务CSM利润率":
                nb_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务CSM（集团口径）')]['(百万)人民币'].sum()
                pv_profit_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务未来现金流入现值（盈利）')]['(百万)人民币'].sum()
                current_val = (nb_curr / pv_profit_curr * 100) if pv_profit_curr != 0 else np.nan
                
                nb_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务CSM（集团口径）')]['(百万)人民币'].sum()
                pv_profit_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务未来现金流入现值（盈利）')]['(百万)人民币'].sum()
                prev_val = (nb_prev / pv_profit_prev * 100) if pv_profit_prev != 0 else np.nan
            
            # 7. 新业务LC亏损率 = 新业务亏损合同 / 新业务未来现金流入现值（亏损） * 100
            elif field_name == "新业务LC亏损率":
                lc_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务亏损合同（LC）——非PAA')]['(百万)人民币'].sum()
                pv_loss_curr = df_current[(df_current['公司'] == co) & (df_current['字段名'] == '新业务未来现金流入现值（亏损）')]['(百万)人民币'].sum()
                current_val = (lc_curr / pv_loss_curr * 100) if pv_loss_curr != 0 else np.nan
                
                lc_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务亏损合同（LC）——非PAA')]['(百万)人民币'].sum()
                pv_loss_prev = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == '新业务未来现金流入现值（亏损）')]['(百万)人民币'].sum()
                prev_val = (lc_prev / pv_loss_prev * 100) if pv_loss_prev != 0 else np.nan
           
            # 8. 其他普通字段：直接从数据中读取
            else:
                # 注意：x_field 现在是显示名，需要映射到原始字段名
                original_field_map = {
                    "保险服务收入": "保险服务收入合计",
                    "保险服务费用": "保险服务费用合计",
                    "净利润": "净利润",
                    "税前利润": "税前利润总额",
                    "净投资回报": "净投资回报",
                    "其他综合收益": "其他综合收益",
                    "综合收益总额": "综合收益总额",
                    "CSM余额": "CSM期末余额",
                    "净资产": "期末股东权益",
                    "新业务CSM": "新业务CSM（集团口径）",
                }
                
                original_field = original_field_map.get(field_name, field_name)
                
                current_rows = df_current[(df_current['公司'] == co) & (df_current['字段名'] == original_field)]
                if current_rows.empty:
                    continue
                current_val = current_rows['(百万)人民币'].iloc[0] / 100
                
                prev_rows = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == original_field)]
                if prev_rows.empty:
                    continue
                prev_val = prev_rows['(百万)人民币'].iloc[0] / 100
            
            # 跳过空值
            if current_val is None or prev_val is None or pd.isna(current_val) or pd.isna(prev_val):
                continue
            
            # 计算增长率
            if prev_val != 0:
                growth_rate = (current_val - prev_val) / abs(prev_val) * 100
            else:
                growth_rate = 0
            
            scatter_data.append({
                '公司': co,
                '公司类型': co_type,
                display_name: current_val,
                f'{display_name}增长率': growth_rate
            })
        
        df_result = pd.DataFrame(scatter_data)
        
        if not df_result.empty:
            df_result = df_result.sort_values(display_name, ascending=False).reset_index(drop=True)
            df_result['id'] = range(1, len(df_result) + 1)
    
        
        return df_result

    def calc_stack_distribution(df_raw, selected_types, latest_year, calc_func, x_bins_custom=None):
        """
        计算堆叠分布数据
        输入：原始数据、公司类型、目标年份、计算公式、自定义区间
        输出：分布矩阵、区间标签
        """
        df_year = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)].copy()
        
        # 1. 计算各公司比率
        company_ratios = {}
        company_type_map = {}
        
        for co in df_year['公司'].unique():
            df_co = df_year[df_year['公司'] == co]
            ratio = calc_func(df_co)
            if ratio is not None and not np.isnan(ratio):
                company_ratios[co] = ratio
                co_type = df_co['公司类型'].iloc[0] if not df_co.empty else None
                if co_type:
                    company_type_map[co] = co_type
        
        # 🌟 关键修复：如果没有有效数据，直接返回空 DataFrame 和空列表
        if not company_ratios:
            return pd.DataFrame(), []
        
        # 2. 确定分箱边界
        ratios_list = list(company_ratios.values())
        if x_bins_custom and len(x_bins_custom) > 1:
            bins = sorted(set(x_bins_custom))
            labels = [f"({int(bins[i])}, {int(bins[i+1])})" for i in range(len(bins) - 1)]
        else:
            min_r, max_r = min(ratios_list), max(ratios_list)
            n_bins = min(6, max(2, len(set(ratios_list))))
            if max_r == min_r:
                bins = [min_r - 1, max_r + 1]
            else:
                step = (max_r - min_r) / n_bins
                bins = [min_r + i * step for i in range(n_bins + 1)]
            labels = [f"({bins[i]:.1f}, {bins[i+1]:.1f})" for i in range(len(bins) - 1)]
        
        # 3. 按公司类型统计分布
        distribution = {ct: {lbl: 0 for lbl in labels} for ct in selected_types}
        
        def get_bin_idx(val, bins):
            if val <= bins[0]:
                return 0
            for i, (low, high) in enumerate(zip(bins[:-1], bins[1:])):
                if low < val <= high:
                    return i
            return len(bins) - 2
        
        for co, ratio in company_ratios.items():
            ct = company_type_map.get(co)
            if ct in selected_types:
                idx = get_bin_idx(ratio, bins)
                if 0 <= idx < len(labels):
                    distribution[ct][labels[idx]] += 1
        
        distribution_df = pd.DataFrame(distribution).T.reindex(selected_types).fillna(0)
        return distribution_df, labels

    def calc_industry_composition(df, company_type, year, numerator_fields, denominator_field):
        """计算单个公司类型的行业平均构成（返回百分比字典）
        
        优化：自动排除数据不完整的公司（已知项总和 < 10%）
        """
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()

        if df_filtered.empty:
            return None

        companies = df_filtered['公司'].unique()
        company_ratios = []

        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            denominator_row = df_co[df_co['字段名'] == denominator_field]
            if denominator_row.empty:
                continue

            denominator = denominator_row['(百万)人民币'].iloc[0]
            if pd.isna(denominator) or denominator == 0:
                continue

            ratios = {}
            for field in numerator_fields:
                numerator_row = df_co[df_co['字段名'] == field]
                if numerator_row.empty or pd.isna(numerator_row['(百万)人民币'].iloc[0]):
                    ratio = 0.0
                else:
                    ratio = numerator_row['(百万)人民币'].iloc[0] / denominator
                ratios[field] = ratio

            # 计算已知项总和
            total = sum(ratios.values())
            
            # 排除数据不完整的公司（总和小于10%）
            if total < 0.1:
                continue
            
            # 添加"其他"项补足到100%
            if '其他收入调整' in numerator_fields:
                other_field = '其他收入调整'
                ratios[other_field] = max(0, 1 - total)
            
            company_ratios.append(ratios)

        if not company_ratios:
            return None

        ratios_df = pd.DataFrame(company_ratios)
        result = {field: ratios_df[field].mean() * 100 for field in ratios_df.columns}
        
        return result

    def calc_industry_two_cat_composition(df, company_type, year, fields):
        """计算单个公司类型的行业平均构成（两个大类：PAA和非PAA）"""
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()

        if df_filtered.empty:
            return None

        companies = df_filtered['公司'].unique()
        company_ratios = []

        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            total = 0
            for f in fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                total += val
            if total == 0:
                continue
            
            ratios = {}
            for f in fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                ratios[f] = val / total * 100
            company_ratios.append(ratios)

        if not company_ratios:
            return None

        ratios_df = pd.DataFrame(company_ratios)
        return {field: ratios_df[field].mean() for field in fields}

    def calc_industry_expense_structure(df, company_type, year):
        """计算单个公司类型的行业平均费用结构（返回占比字典和平均总费用）"""
        expense_fields = ["获取费用", "维持费用", "非履约费用"]
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_ratios = []
        company_totals = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            total = 0
            for f in expense_fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                total += val
            
            if total == 0:
                continue
            
            ratios = {}
            for f in expense_fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                ratios[f] = val / total * 100
            company_ratios.append(ratios)
            company_totals.append(total)
        
        if not company_ratios:
            return None
        
        ratios_df = pd.DataFrame(company_ratios)
        return {
            'ratios': {field: ratios_df[field].mean() for field in expense_fields},
            'avg_total': np.mean(company_totals)
        }
    

    def calc_industry_profit_composition(df, company_type, year):
        """计算行业平均利润构成 - 不过滤异常公司，但记录剔除原因"""
        import pandas as pd
        import numpy as np
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None, []
        
        companies = df_filtered['公司'].unique()
        company_contributions = []
        excluded_companies = []  # 记录被剔除的公司及原因（仅用于展示，不参与计算）
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            # 获取保险利润
            insurance_profit_row = df_co[df_co['字段名'] == "保险利润"]
            if insurance_profit_row.empty:
                excluded_companies.append(f"{co}: 缺少保险利润字段")
                continue
            
            insurance_profit = insurance_profit_row['(百万)人民币'].iloc[0]
            
            # 🔥 关键修改：不再过滤！即使保险利润 <= 0 也参与计算
            # 但记录原因供展示
            if pd.isna(insurance_profit):
                excluded_companies.append(f"{co}: 保险利润为空")
                continue
            
            if insurance_profit <= 0:
                excluded_companies.append(f"{co}: 保险利润≤0")
                # 仍然继续计算，不过滤
            
            # 获取各组成部分
            csm_amort = df_co[df_co['字段名'] == "合同服务边际的摊销"]['(百万)人民币'].sum()
            ra_change = df_co[df_co['字段名'] == "非金融风险调整的变动"]['(百万)人民币'].sum()
            lc_recog = df_co[df_co['字段名'] == "亏损部分的确认及转回"]['(百万)人民币'].sum()
            paa_result = df_co[df_co['字段名'] == "采用保费分配法计量的保险合同保险业绩"]['(百万)人民币'].sum()
            reinsurance = df_co[df_co['字段名'] == "再保净损益"]['(百万)人民币'].sum()
            
            # 营运偏差 = 保险利润 - 各项和
            operating_var = insurance_profit - csm_amort - ra_change + lc_recog - paa_result - reinsurance
            
            # 🔥 关键修改：防止除零错误
            if insurance_profit == 0:
                # 如果保险利润为0，所有贡献设为0
                contribution = {
                    "合同服务边际的释放": 0.0,
                    "非金融风险调整的变动": 0.0,
                    "亏损部分的确认及转回": 0.0,
                    "营运偏差及其他": 0.0,
                    "保费分配法业务净损益": 0.0,
                    "再保净损益": 0.0
                }
            else:
                # 正常计算各项贡献百分比
                contribution = {
                    "合同服务边际的释放": csm_amort / insurance_profit * 100,
                    "非金融风险调整的变动": ra_change / insurance_profit * 100,
                    "亏损部分的确认及转回": -lc_recog / insurance_profit * 100,
                    "营运偏差及其他": operating_var / insurance_profit * 100,
                    "保费分配法业务净损益": paa_result / insurance_profit * 100,
                    "再保净损益": reinsurance / insurance_profit * 100
                }
            
            company_contributions.append(contribution)
        
        if not company_contributions:
            return None, excluded_companies
        
        # 使用平均值（包含异常值，让所有类型都能显示）
        contributions_df = pd.DataFrame(company_contributions)
        
        result = {
            'contributions': {col: contributions_df[col].mean() for col in contributions_df.columns},
            'csm_ratio': contributions_df["合同服务边际的释放"].mean(),
            'excluded': excluded_companies
        }
        
        return result, excluded_companies
    
    
    def calc_industry_profit_mix(df, company_type, year):
        """计算单个公司类型的行业平均利润构成（保险利润和投资利润的占比）"""
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_ratios = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            insurance_profit = df_co[df_co['字段名'] == "保险利润"]['(百万)人民币'].sum()
            investment_profit = df_co[df_co['字段名'] == "投资利润"]['(百万)人民币'].sum()
            total_profit = insurance_profit + investment_profit
            
            if total_profit == 0:
                continue
            
            ratios = {
                "保险利润": insurance_profit / total_profit * 100,
                "投资利润": investment_profit / total_profit * 100
            }
            company_ratios.append(ratios)
        
        if not company_ratios:
            return None
        
        ratios_df = pd.DataFrame(company_ratios)
        return {field: ratios_df[field].mean() for field in ["保险利润", "投资利润"]}

    def calc_industry_oci_composition(df, company_type, year):
        """计算单个公司类型的行业平均OCI构成"""
        oci_fields = ['可转损益OCI合计', '不可转损益OCI合计', '净利润', '综合收益总额']
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_data = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            data = {}
            for f in oci_fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                data[f] = val
            
            data['other'] = data['综合收益总额'] - data['净利润'] - data['可转损益OCI合计'] - data['不可转损益OCI合计']
            
            company_data.append(data)
        
        if not company_data:
            return None
        
        df_company = pd.DataFrame(company_data)
        avg_data = {
            "净利润": df_company['净利润'].mean(),
            "可转损益OCI合计": df_company['可转损益OCI合计'].mean(),
            "不可转损益OCI合计": df_company['不可转损益OCI合计'].mean(),
            "other": df_company['other'].mean(),
            "综合收益总额": df_company['综合收益总额'].mean()
        }
        
        return avg_data

    def calc_industry_oci_deep(df, company_type, year):
        """计算单个公司类型的行业平均OCI深度分析"""
        fields = ['可转损益的负债OCI', 'FVOCI债券公允价值']
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_data = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            data = {}
            for f in fields:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                data[f] = val
            
            company_data.append(data)
        
        if not company_data:
            return None
        
        df_company = pd.DataFrame(company_data)
        return {
            "可转损益的负债OCI": df_company['可转损益的负债OCI'].mean(),
            "FVOCI债券公允价值": df_company['FVOCI债券公允价值'].mean()
        }
    
    def calc_csm_bel_ratio(df_co):
        """CSM/BEL占比 = CSM期末余额 / BEL期末余额 × 100%"""
        csm = df_co[df_co['字段名'] == 'CSM期末余额']['(百万)人民币'].sum()
        bel = df_co[df_co['字段名'] == 'BEL期末余额']['(百万)人民币'].sum()
        return csm / bel * 100 if bel != 0 else np.nan
    
    def calc_csm_continuity_ratio(df_co):
        """CSM持续率 = 新业务CSM / CSM摊销 × 100%"""
        nb = df_co[df_co['字段名'] == '新业务CSM（集团口径）']['(百万)人民币'].sum()
        amort = df_co[df_co['字段名'] == 'CSM摊销']['(百万)人民币'].sum()
        return -nb / amort * 100 if amort != 0 else np.nan
    
    def calc_csm_amort_ratio(df_co):
        """CSM摊销比率 = CSM摊销 / (期末CSM - CSM摊销) × 100%"""
        amort = df_co[df_co['字段名'] == 'CSM摊销']['(百万)人民币'].sum()
        end = df_co[df_co['字段名'] == 'CSM期末余额']['(百万)人民币'].sum()
        denominator = end - amort
        return -amort / denominator * 100 if denominator != 0 else np.nan
    
    def calc_nb_ifrs_margin(df_co):
        """新业务IFRS利润率 = (新业务CSM - 新业务亏损合同) / 新业务未来现金流入现值 × 100%"""
        nb = df_co[df_co['字段名'] == '新业务CSM（集团口径）']['(百万)人民币'].sum()
        lc = df_co[df_co['字段名'] == '新业务亏损合同（LC）——非PAA']['(百万)人民币'].sum()
        pv = df_co[df_co['字段名'] == '新业务未来现金流入现值']['(百万)人民币'].sum()
        return (nb - lc) / pv * 100 if pv != 0 else np.nan        

    def calc_nb_csm_margin(df_co):
        """新业务CSM利润率 = 新业务CSM / 新业务未来现金流入现值（盈利） × 100%"""
        nb_csm = df_co[df_co['字段名'] == '新业务CSM（集团口径）']['(百万)人民币'].sum()
        pv_profit = df_co[df_co['字段名'] == '新业务未来现金流入现值（盈利）']['(百万)人民币'].sum()
        return nb_csm / pv_profit * 100 if pv_profit != 0 else np.nan

    def calc_industry_asset_composition(df, company_type, year):
        """计算单个公司类型的行业平均资产构成"""
        ASSET_FIELDS = ["AC", "FVOCI", "FVTPL", "指定FVOCI"]
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_ratios = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            total = sum(df_co[df_co['字段名'] == f]['(百万)人民币'].sum() for f in ASSET_FIELDS)
            if total == 0:
                continue
            ratios = {f: df_co[df_co['字段名'] == f]['(百万)人民币'].sum() / total * 100 for f in ASSET_FIELDS}
            company_ratios.append(ratios)
        
        if not company_ratios:
            return None
        return {f: pd.DataFrame(company_ratios)[f].mean() for f in ASSET_FIELDS}
    def calc_industry_csm_composition(df, company_type, year):
        """
        计算单个公司类型的行业平均摊销前 CSM 构成
        完全参照 Step 7 的逻辑：分母用绝对值之和
        """
        CSM_FIELDS = ["新业务CSM（集团口径）", "CSM计息", "CSM调整"]
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        
        # 收集所有公司的占比（与 Step 7 完全一致）
        all_ratios = {f: [] for f in CSM_FIELDS}
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            vals = {}
            for f in CSM_FIELDS:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                vals[f] = val if pd.notna(val) else 0
            
            # Step 7 核心：分母用绝对值之和
            total = abs(vals["新业务CSM（集团口径）"]) + abs(vals["CSM计息"]) + abs(vals["CSM调整"])
            
            if total == 0:
                continue
            
            # 计算每个字段的占比（保留原始符号）
            for f in CSM_FIELDS:
                ratio = vals[f] / total * 100
                all_ratios[f].append(ratio)
        
        if not all_ratios["新业务CSM（集团口径）"]:
            return None
        
        # 简单平均（与 Step 7 保持一致）
        result = {}
        for f in CSM_FIELDS:
            result[f] = np.mean(all_ratios[f])
        
        return result
    
    def calc_industry_expense_composition(df, company_type, year):
        """
        计算单个公司类型的行业平均保险服务费用构成
        完全参照 Step 7 的 exp_1 逻辑：分母用绝对值之和
        """
        EXPENSE_FIELDS = [
            "保险获取现金流的摊销（保险服务费用）",
            "亏损部分的确认及转回",
            "当期发生的赔款及其他相关费用",
            "已发生赔款负债相关的履约现金流量变动"
        ]
        
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        
        # 收集所有公司的占比
        all_ratios = {f: [] for f in EXPENSE_FIELDS}
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            vals = {}
            for f in EXPENSE_FIELDS:
                val = df_co[df_co['字段名'] == f]['(百万)人民币'].sum()
                vals[f] = val if pd.notna(val) else 0
            
            # 分母：绝对值之和
            total = sum(abs(vals[f]) for f in EXPENSE_FIELDS)
            
            if total == 0:
                continue
            
            # 占比 = 原始值 / 绝对值之和 × 100（保留符号）
            for f in EXPENSE_FIELDS:
                ratio = vals[f] / total * 100
                all_ratios[f].append(ratio)
        
        if not all_ratios[EXPENSE_FIELDS[0]]:
            return None
        
        # 简单平均
        result = {}
        for f in EXPENSE_FIELDS:
            result[f] = np.mean(all_ratios[f])
        
        return result

    # SCATTER_AXIS_META：散点图统一配置

    SCATTER_AXIS_META = {
        # 1. 保险服务收入
        "industry_inc_total": {
            "title": "保险服务收入",
            "x_field": "保险服务收入",
            "y_field": "保险服务收入增长率",
            "x_label": "保险服务收入区间（亿元）",
            "y_label": "保险服务收入增长率区间（%）",
            "default_x": "0,100,200,500,1000,2000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 2. 净投资回报
        "industry_inv_return": {
            "title": "净投资回报",
            "x_field": "净投资回报",
            "y_field": "净投资回报增长率",
            "x_label": "净投资回报区间（亿元）",
            "y_label": "净投资回报增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 3. 保险服务费用
        "industry_exp_total": {
            "title": "保险服务费用",
            "x_field": "保险服务费用",
            "y_field": "保险服务费用增长率",
            "x_label": "保险服务费用区间（亿元）",
            "y_label": "保险服务费用增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 4. 承保财务净损益（计算字段）
        "industry_uw_profit": {
            "title": "承保财务净损益",
            "x_field": "承保财务净损益",
            "y_field": "承保财务净损益增长率",
            "x_label": "承保财务净损益区间（亿元）",
            "y_label": "承保财务净损益增长率区间（%）",
            "default_x": "-1000,-500,-200,-100,0,100,200,500",
            "default_y": "-1000,-500,-200,-100,0,100,200,500",
        },
        # 5. 保险服务业绩（计算字段）
        "industry_perf_total": {
            "title": "保险服务业绩",
            "x_field": "保险服务业绩",
            "y_field": "保险服务业绩增长率",
            "x_label": "保险服务业绩区间（亿元）",
            "y_label": "保险服务业绩增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 6. 投资利润（计算字段）
        "industry_inv_profit": {
            "title": "投资利润",
            "x_field": "投资利润",
            "y_field": "投资利润增长率",
            "x_label": "投资利润区间（亿元）",
            "y_label": "投资利润增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 7. 净利润
        "industry_net_profit": {
            "title": "净利润",
            "x_field": "净利润",
            "y_field": "净利润增长率",
            "x_label": "净利润区间（亿元）",
            "y_label": "净利润增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 8. 税前利润
        "industry_tax_profit": {
            "title": "税前利润",
            "x_field": "税前利润",
            "y_field": "税前利润增长率",
            "x_label": "税前利润区间（亿元）",
            "y_label": "税前利润增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 9. 其他综合收益
        "industry_oci_profit": {
            "title": "其他综合收益",
            "x_field": "其他综合收益",
            "y_field": "其他综合收益增长率",
            "x_label": "其他综合收益区间（亿元）",
            "y_label": "其他综合收益增长率区间（%）",
            "default_x": "-500,-200,-100,0,100,200,500",
            "default_y": "-1000,-500,-200,0,200,500,1000",
        },
        # 10. 综合收益总额
        "industry_total_profit": {
            "title": "综合收益总额",
            "x_field": "综合收益总额",
            "y_field": "综合收益总额增长率",
            "x_label": "综合收益总额区间（亿元）",
            "y_label": "综合收益总额增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,0,100,200,500,1000",
        },
        # 11. CSM余额
        "industry_csm_bal": {
            "title": "CSM余额",
            "x_field": "CSM余额",
            "y_field": "CSM余额增长率",
            "x_label": "CSM余额区间（亿元）",
            "y_label": "CSM余额增长率区间（%）",
            "default_x": "0,100,500,1000,2000,5000",
            "default_y": "-50,-20,0,20,50,100,200",
        },
        # 12. 净资产
        "industry_equity_trend": {
            "title": "净资产",
            "x_field": "净资产",
            "y_field": "净资产增长率",
            "x_label": "净资产区间（亿元）",
            "y_label": "净资产增长率区间（%）",
            "default_x": "0,100,500,1000,2000,5000",
            "default_y": "-50,-20,0,20,50,100,200",
        },
        # 13. CSM/净资产占比（计算字段，变化是百分点）
        "industry_csm_equity": {
            "title": "CSM/净资产占比",
            "x_field": "CSM/净资产占比",
            "y_field": "CSM/净资产占比增长率",
            "x_label": "CSM/净资产占比（%）",
            "y_label": "占比变化（百分点）",
            "default_x": "0,20,40,60,80,100,120",
            "default_y": "-20,-10,-5,0,5,10,20",
        },
        # 14. 新业务CSM
        "industry_nb_csm": {
            "title": "新业务CSM",
            "x_field": "新业务CSM",
            "y_field": "新业务CSM增长率",
            "x_label": "新业务CSM区间（亿元）",
            "y_label": "新业务CSM增长率区间（%）",
            "default_x": "0,50,100,200,500,1000",
            "default_y": "-100,-50,-20,0,20,50,100",
        },
        # 15. 新业务LC亏损率（计算字段，变化是百分点）
        "industry_lc_loss_ratio": {
            "title": "新业务LC亏损率",
            "x_field": "新业务LC亏损率",
            "y_field": "新业务LC亏损率增长率",
            "x_label": "LC亏损率（%）",
            "y_label": "LC亏损率变化（百分点）",
            "default_x": "0,10,20,30,40,50,60,70,80,90,100",
            "default_y": "-50,-30,-20,-10,0,10,20,30,50",
        },
    }
    
    # STACK_DIST_META：堆叠分布图统一配置
    STACK_DIST_META = {
        "industry_csm_ratio": {
            "name": "CSM/BEL占比",
            "calc_func": calc_csm_bel_ratio,
            "default_x": "0,20,40,60,80,100,120",
            "y_label": "公司数量",
        },
        "industry_csm_continuity_ratio": {
            "name": "CSM持续率",
            "calc_func": calc_csm_continuity_ratio,
            "default_x": "0,50,100,150,200,300",
            "y_label": "公司数量",
        },
        "industry_csm_amort_ratio": {
            "name": "CSM摊销比率",
            "calc_func": calc_csm_amort_ratio,
            "default_x": "0,5,10,15,20,30",
            "y_label": "公司数量",
        },
        "industry_nb_ifrs_margin": {
            "name": "新业务IFRS利润率",
            "calc_func": calc_nb_ifrs_margin,
            "default_x": "0,10,20,30,40,50",
            "y_label": "公司数量",
        },
    }   
   
    
   
    def get_uploaded_bins_by_mid(m_id):
        custom_map = st.session_state.get("custom_bins_map", {})
        cfg = custom_map.get(m_id, {})
        return cfg.get("x_bins_custom") or [], cfg.get("y_bins_custom") or []
    def render_axis_bin_popover(m_id, title, default_x_bins, default_y_bins):
        uploaded_x, uploaded_y = get_uploaded_bins_by_mid(m_id)

        use_upload_key = f"use_uploaded_bins_{m_id}"
        x_text_key = f"x_bins_text_{m_id}"
        y_text_key = f"y_bins_text_{m_id}"

        if use_upload_key not in st.session_state:
            st.session_state[use_upload_key] = bool(uploaded_x or uploaded_y)
        if x_text_key not in st.session_state:
            st.session_state[x_text_key] = bins_to_str(uploaded_x or default_x_bins)
        if y_text_key not in st.session_state:
            st.session_state[y_text_key] = bins_to_str(uploaded_y or default_y_bins)

        with st.popover(f"⚙️ 当前图参数设置：{title}", use_container_width=True):
            st.checkbox("优先使用上传 Excel 的刻度", key=use_upload_key)

            if uploaded_x or uploaded_y:
                st.caption(f"已上传：X = {bins_to_str(uploaded_x) or '（空）'}；Y = {bins_to_str(uploaded_y) or '（空）'}")

            st.text_input("X 轴刻度（英文逗号分隔）", key=x_text_key, placeholder=bins_to_str(default_x_bins))
            st.text_input("Y 轴刻度（英文逗号分隔）", key=y_text_key, placeholder=bins_to_str(default_y_bins))

            if st.button("恢复默认刻度", key=f"reset_bins_{m_id}", use_container_width=True):
                st.session_state[use_upload_key] = False
                st.session_state[x_text_key] = bins_to_str(default_x_bins)
                st.session_state[y_text_key] = bins_to_str(default_y_bins)
                st.rerun()

        if st.session_state[use_upload_key] and (uploaded_x or uploaded_y):
            x_bins = uploaded_x or default_x_bins
            y_bins = uploaded_y or default_y_bins
        else:
            x_bins = parse_bins_input_safe(st.session_state[x_text_key], default_x_bins)
            y_bins = parse_bins_input_safe(st.session_state[y_text_key], default_y_bins)

        return x_bins, y_bins
    
    def bins_to_str(bins):
        """将 bins 列表转为逗号分隔的字符串，整数不显示小数点"""
        if not bins:
            return ""
        return ",".join(str(int(x)) if x == int(x) else str(x) for x in bins)

    def parse_bins_input_safe(text, fallback):
        """安全解析 bins 输入，失败时返回 fallback"""
        try:
            result = parse_bins_input(text)
            return result if result else fallback
        except Exception:
            return fallback
        
    def get_dist_bins(m_id, default_x):
        """获取堆叠分布图的 X 轴区间（优先级：上传Excel > 页面输入 > 默认）"""
        # 1. 优先使用上传 Excel 的 x_bins
        uploaded_x, _ = get_uploaded_bins_by_mid(m_id)
        if uploaded_x:
            return uploaded_x
        
        # 2. 其次使用页面手工输入
        use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
        if use_custom:
            # 改成 parse_bins_input
            x_bins = parse_bins_input(st.session_state.get(f"{m_id}_dist_x_bins", default_x))
            if x_bins:
                return x_bins
        
        # 3. 最后使用默认值
        return parse_bins_input(default_x)
    
    # ==========================================
    # 第3层：绘图函数（只负责画图，不计算数据）
    # ==========================================

    def apply_kpmg_stack_style(fig, year_list, selected_types, extra_layout=None):
        n_cols = len(selected_types) if isinstance(selected_types, list) else fig.layout.annotations.__len__()
        
        for i in range(1, n_cols + 1):
            fig.add_shape(type="rect", xref="x domain" if i == 1 else f"x{i} domain",
                          yref="y domain" if i == 1 else f"y{i} domain",
                          x0=-0.08, x1=1.08, y0=-0.12, y1=1.15,
                          fillcolor="rgba(200, 200, 200, 0.15)",
                          line=dict(color="rgba(150, 150, 150, 0.6)", width=1.5), layer="below", row=1, col=i)
        
        layout_update = dict(barmode='stack', bargap=0.25, bargroupgap=0.0,
                             legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5, font=dict(size=10)))
        if extra_layout:
            layout_update.update(extra_layout)
        fig.update_layout(**layout_update)
        return fig
    
    def create_scatter_chart(df_scatter, config, x_label, y_label):
        """Plotly 版散点热力图 - 紧凑单列合并表格版"""
        import plotly.graph_objects as go
        import numpy as np
        import pandas as pd
        from plotly.subplots import make_subplots
    
        if df_scatter.empty:
            st.warning("没有有效数据")
            return
    
        category_colors = {
            "头部": get_company_color("头部"),
            "银行系": get_company_color("银行系"),
            "外资系": get_company_color("外资系"),
            "养老健康": get_company_color("养老健康"),
            "小型": get_company_color("小型"),
        }
        category_order = ["头部", "银行系", "外资系", "养老健康", "小型"]
    
        x_col = config['x_col']
        y_col = config['y_col']
    
        # ── 字体尺寸独立控制 ──────────────────────────────────────────────────────
        CHART_FONT_SIZE   = 10    # 热力图散点编号、轴标签字体（独立）
        TABLE_FONT_SIZE   = 6    # 表格单元格字体（独立）
        TABLE_HEADER_SIZE = 7    # 表头字体
        TABLE_ROW_HEIGHT  = 13   # 每行高度（压缩以容纳更多行）
        TABLE_HEADER_H    = 14   # 表头高度
    
        # 数据排序
        existing_types = [ct for ct in category_order if ct in df_scatter['公司类型'].values]
        df_sorted = pd.DataFrame()
        for ct in existing_types:
            df_ct = df_scatter[df_scatter['公司类型'] == ct].sort_values(x_col, ascending=False)
            df_sorted = pd.concat([df_sorted, df_ct])
        other_types = [ct for ct in df_scatter['公司类型'].unique() if ct not in existing_types]
        for ct in other_types:
            df_ct = df_scatter[df_scatter['公司类型'] == ct].sort_values(x_col, ascending=False)
            df_sorted = pd.concat([df_sorted, df_ct])
    
        df_scatter = df_sorted.reset_index(drop=True)
        df_scatter['id'] = range(1, len(df_scatter) + 1)
    
        x_values = df_scatter[x_col].values
        y_values = df_scatter[y_col].values
        min_x, max_x = min(x_values), max(x_values)
        min_y, max_y = min(y_values), max(y_values)
    
        # X轴区间
        if config.get('x_bins_custom') and len(config['x_bins_custom']) > 1:
            bins_x = list(config['x_bins_custom'])
            x_labels = [f"({int(bins_x[i])}, {int(bins_x[i+1])}]" for i in range(len(bins_x) - 1)]
        else:
            step = max(10, int((max_x - min_x) / 4 / 10) * 10) if max_x > 10 else max(1, int((max_x - min_x) / 4))
            bins_x = list(np.arange(min_x - 5, max_x + step, step))
            if len(bins_x) < 2:
                bins_x = [min_x - 5, max_x + 5]
            bins_x = sorted(set(bins_x))
            x_labels = [f"({int(bins_x[i])}, {int(bins_x[i+1])}]" for i in range(len(bins_x) - 1)]
    
        # Y轴区间
        if config.get('y_bins_custom') and len(config['y_bins_custom']) > 1:
            bins_y = list(config['y_bins_custom'])
            y_labels = [f"({int(bins_y[i])}, {int(bins_y[i+1])}]" for i in range(len(bins_y) - 1)]
        else:
            y_range = max_y - min_y
            step = max(10, int(y_range / 5)) if y_range != 0 else 10
            bins_y = list(np.arange(min_y - 10, max_y + step, step))
            if len(bins_y) < 2:
                bins_y = [min_y - 10, max_y + 10]
            bins_y = sorted(set(bins_y))
            y_labels = [f"({int(bins_y[i])}%, {int(bins_y[i+1])}%]" for i in range(len(bins_y) - 1)]
    
        nx, ny = len(x_labels), len(y_labels)
    
        def get_bin_idx(val, bins):
            if val <= bins[0]: return 0
            for i, (low, high) in enumerate(zip(bins[:-1], bins[1:])):
                if low < val <= high: return i
            return len(bins) - 2
    
        df_scatter['x_bin'] = df_scatter[x_col].apply(lambda v: get_bin_idx(v, bins_x))
        df_scatter['y_bin'] = df_scatter[y_col].apply(lambda v: get_bin_idx(v, bins_y))
    
        density = np.zeros((ny, nx))
        for _, row in df_scatter.iterrows():
            if 0 <= row['x_bin'] < nx and 0 <= row['y_bin'] < ny:
                density[int(row['y_bin']), int(row['x_bin'])] += 1
    
        def get_position_in_bin(val, bins, bin_idx):
            low, high = bins[bin_idx], bins[bin_idx + 1]
            if high > low:
                frac = (val - low) / (high - low)
                frac = max(0.1, min(0.9, frac))
            else:
                frac = 0.5
            return bin_idx + frac
    
        df_scatter['x_pos'] = df_scatter.apply(
            lambda r: get_position_in_bin(r[x_col], bins_x, int(r['x_bin'])), axis=1)
        df_scatter['y_pos'] = df_scatter.apply(
            lambda r: get_position_in_bin(r[y_col], bins_y, int(r['y_bin'])), axis=1)
    
        np.random.seed(42)
        for y in range(ny):
            for x in range(nx):
                mask = (df_scatter['x_bin'] == x) & (df_scatter['y_bin'] == y)
                indices = df_scatter[mask].index.tolist()
                n = len(indices)
                if n > 1:
                    for i, idx in enumerate(indices):
                        angle = (i / n) * 2 * np.pi
                        radius = 0.13 * (1 + 0.15 * i)
                        df_scatter.loc[idx, 'x_pos'] += radius * np.cos(angle)
                        df_scatter.loc[idx, 'y_pos'] += radius * np.sin(angle)
                        df_scatter.loc[idx, 'x_pos'] = np.clip(df_scatter.loc[idx, 'x_pos'], x + 0.08, x + 0.92)
                        df_scatter.loc[idx, 'y_pos'] = np.clip(df_scatter.loc[idx, 'y_pos'], y + 0.08, y + 0.92)
    
        # ── 表格数据准备 ──────────────────────────────────────────────────────────
        df_table = df_scatter.copy()
    
        # ── 修复2：公司名截断放宽到8字，避免列被撑宽 ─────────────────────────────
        def shorten_name(name, max_len=8):
            if len(name) > max_len:
                return name[:max_len - 1] + "."
            return name
    
        # ── 修复1：总高度 = 行数×行高 + 表头 + margin，确保所有行可见 ────────────
        n_rows      = len(df_table)
        table_body_h = n_rows * TABLE_ROW_HEIGHT + TABLE_HEADER_H + 40   # 40px 缓冲
        chart_min_h  = 550                                                 # 热力图最小高度
        total_height = max(table_body_h, chart_min_h) + 70                # +70 留给 legend/margin
    
        # ── subplot：图65% / 表35% ─────────────────────────────────────────────
        fig = make_subplots(
            rows=1, cols=2,
            column_widths=[0.65, 0.35],
            shared_yaxes=False,
            horizontal_spacing=0.015,
            specs=[[{"type": "xy"}, {"type": "table"}]]
        )
    
        # ── 热力图 ────────────────────────────────────────────────────────────────
        max_count = max(1, int(density.max()))
        fig.add_trace(
            go.Heatmap(
                z=density,
                x=[i + 0.5 for i in range(nx)],
                y=[i + 0.5 for i in range(ny)],
                colorscale=[
                    [0,    "#FFFFFF"],
                    [0.15, "#FDEBEC"],
                    [0.4,  "#F9C7C8"],
                    [0.7,  "#F39A9C"],
                    [1.0,  "#ED2124"]
                ],
                showscale=False,
                hoverinfo='none',
                zmin=0,
                zmax=max_count,
                name=""
            ),
            row=1, col=1
        )
    
        # 格子内数量标注 —— 使用独立的 CHART_FONT_SIZE
        for i in range(ny):
            for j in range(nx):
                count = int(density[i, j])
                if count > 0:
                    fig.add_annotation(
                        x=j + 0.85, y=i + 0.85,
                        text=str(count),
                        showarrow=False,
                        font=dict(size=CHART_FONT_SIZE - 1, color="#6B7280"),  # ← 独立图字体
                        row=1, col=1
                    )
    
        point_count = len(df_scatter)
        marker_size = max(9, min(14, int(150 / max(1, point_count / 22))))
    
        for ct in category_order:
            df_ct = df_scatter[df_scatter['公司类型'] == ct]
            if df_ct.empty:
                continue
            color = category_colors.get(ct, "#1E49E2")
    
            fig.add_trace(
                go.Scatter(
                    x=df_ct['x_pos'], y=df_ct['y_pos'],
                    mode='markers', name=ct,
                    marker=dict(size=marker_size, color=color,
                                line=dict(width=0.8, color='white')),
                    hoverinfo='skip', showlegend=True
                ),
                row=1, col=1
            )
            fig.add_trace(
                go.Scatter(
                    x=df_ct['x_pos'], y=df_ct['y_pos'],
                    mode='text', name=ct + "_text",
                    text=df_ct['id'].astype(str),
                    textposition='middle center',
                    # ── 散点编号字体独立于表格 ──
                    textfont=dict(size=max(6, int(marker_size * 0.52)), color='white'),
                    showlegend=False, hoverinfo='skip'
                ),
                row=1, col=1
            )
    
        for i in range(nx + 1):
            fig.add_shape(type="line", x0=i, x1=i, y0=0, y1=ny,
                          line=dict(color="#D9E2EC", width=0.4), layer='above', row=1, col=1)
        for i in range(ny + 1):
            fig.add_shape(type="line", x0=0, x1=nx, y0=i, y1=i,
                          line=dict(color="#D9E2EC", width=0.4), layer='above', row=1, col=1)
    
        # ── 修复2：表格列宽重新分配，压窄公司列 ─────────────────────────────────
        # columnwidth 比例：编号12 / 公司42 / 类型26（合计80，Plotly按比例分配）
        fill_colors = ['#F8FAFC' if i % 2 == 0 else 'white' for i in range(len(df_table))]
        type_colors = [category_colors.get(t, "#1F2937") for t in df_table['公司类型'].tolist()]
    
        fig.add_trace(
            go.Table(
                columnwidth=[12, 25, 25],
                header=dict(
                    values=["编号", "公司", "类型"],
                    fill_color="#00338D",
                    font=dict(color='white', size=TABLE_HEADER_SIZE),
                    align=['center', 'center', 'center'],  # 表头全部居中
                    height=TABLE_HEADER_H
                ),
                cells=dict(
                    values=[
                        df_table['id'].astype(str).tolist(),
                        [shorten_name(str(n)) for n in df_table['公司'].tolist()],
                        df_table['公司类型'].tolist()
                    ],
                    fill_color=[fill_colors, fill_colors, fill_colors],
                    font=dict(
                        color=[
                            ['#0F172A'] * len(df_table),
                            ['#1F2937'] * len(df_table),
                            type_colors,
                        ],
                        size=[TABLE_FONT_SIZE, TABLE_FONT_SIZE, TABLE_FONT_SIZE]
                    ),
                    align=['center', 'center', 'center'],  # 🔴 修改：全部改为 center
                    height=TABLE_ROW_HEIGHT
                )
            ),
            row=1, col=2
        )
            
        # ── 布局 ─────────────────────────────────────────────────────────────────
        fig.update_layout(
            showlegend=True,
            height=total_height,               # ← 修复1：动态高度确保表格不截断
            width=1000,
            plot_bgcolor='white',
            paper_bgcolor='white',
            margin=dict(l=50, r=30, t=40, b=40),
            legend=dict(
                orientation="h",
                yanchor="bottom", y=1.02,
                xanchor="center", x=0.33,
                # ── 图例字体也独立 ──
                font=dict(size=CHART_FONT_SIZE - 1)
            )
        )
    
        # ── 坐标轴字体使用 CHART_FONT_SIZE，与表格完全解耦 ──────────────────────
        fig.update_xaxes(
            title=dict(text=x_label, font=dict(size=CHART_FONT_SIZE)),
            tickmode='array',
            tickvals=[i + 0.5 for i in range(nx)],
            ticktext=x_labels,
            tickangle=-20,
            tickfont=dict(size=CHART_FONT_SIZE - 1, color="#475569"),
            showgrid=False, showline=True,
            linecolor="#CBD5E1", linewidth=0.5,
            range=[-0.1, nx + 0.1],
            row=1, col=1
        )
        fig.update_yaxes(
            title=dict(text=y_label, font=dict(size=CHART_FONT_SIZE)),
            tickmode='array',
            tickvals=[i + 0.5 for i in range(ny)],
            ticktext=y_labels,
            tickfont=dict(size=CHART_FONT_SIZE - 1, color="#475569"),
            showgrid=False, showline=True,
            linecolor="#CBD5E1", linewidth=0.5,
            range=[-0.1, ny + 0.1],
            row=1, col=1
        )
    
        fig.update_xaxes(visible=False, row=1, col=2)
        fig.update_yaxes(visible=False, row=1, col=2)
    
        st.markdown('<div style="display: flex; justify-content: center; margin: 0 auto;">', unsafe_allow_html=True)
        st.plotly_chart(fig, use_container_width=False)
        st.markdown('</div>', unsafe_allow_html=True)
  
    def create_stack_chart_and_table(distribution_df, labels, metric_name, target_year, show_labels, label_size):
        """创建堆叠柱状图和明细表格"""
        ct_list = [ct for ct in distribution_df.index if distribution_df.loc[ct].sum() > 0]
        ct_list = sort_company_types(ct_list)
        
        if not ct_list:
            st.warning("没有有效数据")
            return
        
        # ========== 1. 绘制堆叠柱状图 ==========
        fig = go.Figure()
        for ct in ct_list:
            values = distribution_df.loc[ct].values
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            fig.add_trace(go.Bar(
                x=labels, y=values, name=ct, marker_color=color, width=0.45,
                text=[f"{int(v)}" if show_labels and v > 0 else "" for v in values],
                textposition='inside', textfont=dict(size=label_size, color="white"),textangle=0,  
                hovertemplate=f"{ct}: %{{y:.0f}}家<br>%{{x}}<extra></extra>"
            ))
        
        # Y轴设置
        total_per_bin = distribution_df.sum(axis=0).values
        max_total = int(max(total_per_bin)) if len(total_per_bin) > 0 else 1
        y_max = max_total + 1
        
        if y_max <= 10:
            tick_step = 1
        elif y_max <= 20:
            tick_step = 2
        elif y_max <= 50:
            tick_step = 5
        else:
            tick_step = 10
        
        tick_vals = list(range(0, y_max + 1, tick_step))
        if tick_vals[-1] != y_max:
            tick_vals.append(y_max)
        tick_text = [f"{i}家" for i in tick_vals]
        tick_angle = -15 if len(labels) > 5 else 0
        
        fig.update_layout(
            title=dict(text=f"{metric_name}分布 - {target_year}年", x=0.5, xanchor='center', font=dict(size=14, color="#00338D")),
            barmode='stack', bargap=0.02, bargroupgap=0, height=380,
            width=900,
            autosize=False,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(
                tickfont=dict(size=9), tickangle=tick_angle, showgrid=True, gridcolor="#E8ECF1",
                categoryorder='array', categoryarray=labels,
                title=""
            ),
            yaxis=dict(
                title="公司数量", range=[0, y_max], tickvals=tick_vals, ticktext=tick_text,
                showgrid=True, gridcolor="#E8ECF1", zeroline=True, zerolinecolor="#ccc",
                title_font=dict(size=10)
            ),
            legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5, font=dict(size=9)),
            margin=dict(t=40, b=10, l=40, r=20)
        )
        
        col1, col2, col3 = st.columns([1, 2.5, 1])
        with col2:
            st.plotly_chart(fig, use_container_width=False, config={'displayModeBar': False})
        
        # ========== 2. 明细表格 ==========
        html_rows = []
        for i, ct in enumerate(ct_list):
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            values = distribution_df.loc[ct].values
            total = int(sum(values))
            
            row_bg = "#F8FAFC" if i % 2 == 0 else "white"
            
            # 🌟 公司列左对齐
            row_cells = f'<td style="padding: 4px 6px; border: 1px solid #EAEAEA; background-color: {color}; color: white; font-weight: bold; font-size: 9px; text-align: left;">{ct}</td>'
            for v in values:
                row_cells += f'<td style="padding: 4px 6px; text-align: center; border: 1px solid #EAEAEA; background-color: {row_bg}; font-size: 9px;">{int(v)}家</td>'
            row_cells += f'<td style="padding: 4px 6px; text-align: center; border: 1px solid #EAEAEA; background-color: {row_bg}; font-weight: bold; font-size: 9px;">{total}家</td>'
            
            html_rows.append(f"<tr>{row_cells}</tr>")
            html_table = f"""
                    <div style="margin-top: 5px; max-height: 350px; overflow-y: auto; display: flex; justify-content: center;">
                        <div>
                            <p style="font-size: 11px; font-weight: bold; margin-bottom: 5px; margin-top: 5px; text-align: left;">分布数据明细</p>
                            <table style="width: 1000px; border-collapse: collapse; font-family: sans-serif;">
                                <thead>
                                    <tr style="background-color: #00338D; color: white;">
                                        <th style="padding: 5px 6px; border: 1px solid white; text-align: left; font-size: 9px;">公司类型</th>
                                        {"".join([f'<th style="padding: 5px 6px; border: 1px solid white; text-align: center; font-size: 9px;">{label}</th>' for label in labels])}
                                        <th style="padding: 5px 6px; border: 1px solid white; text-align: center; font-size: 9px;">合计</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {''.join(html_rows)}
                                </tbody>
                            </table>
                        </div>
                    </div>
                    """        

        st.markdown(html_table, unsafe_allow_html=True)
        
    def create_composition_chart_v1(results_by_year, year_list, selected_types, config):
        from plotly.subplots import make_subplots
        
        all_cts = [ct for ct in selected_types if ct in results_by_year.get(year_list[0], {})]
        if not all_cts:
            return go.Figure()
        
        fields = ["采用保费分配法计量的保险合同保险服务收入", "未采用保费分配法计量的保险合同保险服务收入"]
        short_names = {"采用保费分配法计量的保险合同保险服务收入": "采用保费分配法", "未采用保费分配法计量的保险合同保险服务收入": "未采用保费分配法"}
        color_map = {"采用保费分配法计量的保险合同保险服务收入": "#510DBC", "未采用保费分配法计量的保险合同保险服务收入": "#C7A0F7"}
        
        # 🌟 和 V2 完全一致
        fig = make_subplots(
            rows=1, cols=len(all_cts), shared_yaxes=True,
            horizontal_spacing=0.03,  # V2 用的是 0.03
            column_titles=[f"<b>{ct}</b>" for ct in all_cts]
        )
        
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_vals = [f"{year}年" for year in year_list]
            
            for f in fields:
                y_vals = [results_by_year.get(year, {}).get(ct, {}).get(f, 0) for year in year_list]
                is_dark = f == "采用保费分配法计量的保险合同保险服务收入"
                fig.add_trace(
                    go.Bar(
                        x=x_vals, y=y_vals, name=short_names[f],
                        marker_color=color_map[f],
                        text=[f"{v:.1f}%" if config['show_labels'] and v > 0 else "" for v in y_vals],
                        textposition='inside', insidetextanchor='middle',
                        textfont=dict(size=config['label_size'], color="white" if is_dark else "black"),
                        width=config['bar_width'], showlegend=(i == 0), legendgroup=f
                    ),
                    row=1, col=col_idx
                )
            
            # 🌟 和 V2 完全一致的灰色框
            fig.add_shape(
                type="rect",
                xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                yref="y domain" if col_idx == 1 else f"y{col_idx} domain",
                x0=-0.06, x1=1.06,
                y0=-0.1, y1=1.15,
                fillcolor="rgba(0,0,0,0)",
                line=dict(color="#E0E0E0", width=1),
                layer="below",
                row=1, col=col_idx
            )
        
        # 🌟 和 V2 完全一致的布局
        fig.update_layout(
            barmode='stack',
            bargap=0.2,
            height=500,
            margin=dict(t=50, b=80, l=20, r=20),
            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=10)),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        
        # 🌟 和 V2 完全一致的轴设置
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", ticklen=0, row=1, col=i)
            fig.update_yaxes(showgrid=False, range=[0, 100], tickvals=[0, 25, 50, 75, 100],
                            ticktext=["0%", "25%", "50%", "75%", "100%"], zeroline=True, zerolinecolor="#E0E0E0", row=1, col=i)
        
        for ann in fig.layout.annotations:
            if "<b>" in str(ann.text):
                ann.update(y=1.08, font=dict(size=config['co_font_size'], color="#00338D"))
        
        return fig
    
    def create_composition_chart_v2(results_by_year, year_list, selected_types, config):
        from plotly.subplots import make_subplots
        import pandas as pd
        import plotly.graph_objects as go
    
        all_cts = [ct for ct in selected_types if ct in results_by_year.get(year_list[0], {})]
        if not all_cts:
            return go.Figure(), pd.DataFrame()
    
        field_map = {
            "合同服务边际的摊销": "合同服务边际的释放",
            "非金融风险调整的变动": "非金融风险调整的变动",
            "预计当期发生的保险服务费用": "预期当期发生的保险服务费用",
            "保险获取现金流的摊销（保险服务收入）": "保险获取现金流的摊销",
            "与当期服务或过去服务相关得保费经验调整": "与当期服务或过去服务相关的保费经验调整",
            "其他收入调整": "其他"
        }
        color_map = {
            "合同服务边际的摊销": "rgb(30, 73, 226)",
            "非金融风险调整的变动": "rgb(254, 174, 215)",
            "预计当期发生的保险服务费用": "rgb(0, 163, 161)",
            "保险获取现金流的摊销（保险服务收入）": "rgb(1, 184, 245)",
            "与当期服务或过去服务相关得保费经验调整": "rgb(0, 219, 214)",
            "其他收入调整": "rgb(114, 19, 234)"
        }
    
        fields = list(field_map.keys())
        dark_colors = {"rgb(30, 73, 226)", "rgb(114, 19, 234)", "rgb(0, 163, 161)"}
    
        fig = make_subplots(
            rows=1,
            cols=len(all_cts),
            shared_yaxes=True,
            horizontal_spacing=0.03,
            column_titles=[f"<b>{ct}</b>" for ct in all_cts]
        )
    
        df_avg = pd.DataFrame(index=[f"{year}年" for year in year_list], columns=field_map.values()).fillna(0)
    
        shown_legend = set()  # 记录已经显示过图例的系列
    
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_labels = [f"{year}年" for year in year_list]
            cumulative = [0.0] * len(year_list)
    
            for f in fields:
                display_name = field_map[f]
                y_vals = [results_by_year.get(year, {}).get(ct, {}).get(display_name, 0) for year in year_list]
    
                show_legend = display_name not in shown_legend
                if show_legend:
                    shown_legend.add(display_name)
    
                fig.add_trace(
                    go.Bar(
                        x=x_labels,
                        y=y_vals,
                        name=display_name,
                        marker_color=color_map[f],
                        width=config['bar_width'],
                        showlegend=show_legend,
                        legendgroup=display_name,
                        hovertemplate="%{x}<br>%{fullData.name}: %{y:.1f}%<extra></extra>"
                    ),
                    row=1,
                    col=col_idx
                )
    
                if config.get('show_labels', True):
                    txt_c = "white" if color_map[f] in dark_colors else "black"
                    for j, v in enumerate(y_vals):
                        if abs(v) >= 1:
                            y_mid = cumulative[j] + v / 2
                            fig.add_annotation(
                                x=x_labels[j],
                                y=y_mid,
                                text=f"{v:.1f}%",
                                showarrow=False,
                                xref=f"x{col_idx}" if col_idx > 1 else "x",
                                yref=f"y{col_idx}" if col_idx > 1 else "y",
                                font=dict(size=config['label_size'], color=txt_c),
                                xanchor="center",
                                yanchor="middle"
                            )
    
                cumulative = [cumulative[j] + y_vals[j] for j in range(len(year_list))]
    
            fig.add_shape(
                type="rect",
                xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                yref="y domain" if col_idx == 1 else f"y{col_idx} domain",
                x0=-0.06, x1=1.06, y0=-0.1, y1=1.15,
                fillcolor="rgba(0,0,0,0)",
                line=dict(color="#E0E0E0", width=1),
                layer="below",
                row=1,
                col=col_idx
            )
    
            for year_idx, year in enumerate(year_list):
                year_str = f"{year}年"
                for f in fields:
                    df_avg.loc[year_str, field_map[f]] += results_by_year.get(year, {}).get(ct, {}).get(field_map[f], 0)
    
        df_avg = df_avg / len(all_cts)
    
        fig.update_layout(
            barmode='stack',
            bargap=0.2,
            height=500,
            margin=dict(t=50, b=80, l=20, r=20),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.2,
                xanchor="center",
                x=0.5,
                font=dict(size=10)
            ),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)'
        )
    
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", ticklen=0, row=1, col=i)
            fig.update_yaxes(
                showgrid=False,
                range=[0, 100],
                tickvals=[0, 25, 50, 75, 100],
                ticktext=["0%", "25%", "50%", "75%", "100%"],
                zeroline=True,
                zerolinecolor="#E0E0E0",
                row=1,
                col=i
            )
    
        for ann in fig.layout.annotations:
            if "<b>" in str(ann.text):
                ann.update(y=1.08, font=dict(size=config['co_font_size'], color="#00338D"))
    
        return fig, df_avg


    def create_profit_composition_chart(results_by_ct, config, excluded_info=None):
        """绘制行业平均利润构成图 - 强制显示所有公司类型"""
        import plotly.graph_objects as go
        
        if not results_by_ct:
            fig = go.Figure()
            fig.add_annotation(text="无有效数据", x=0.5, y=0.5, showarrow=False)
            return fig, None
        
        # ========== 🔥 关键修改：不再过滤异常值，强制显示所有类型 ==========
        filtered_results = results_by_ct  # 直接使用全部数据，不过滤
        filtered_excluded = excluded_info if excluded_info else {}
        
        # ========== 绘图配置 ==========
        display_mapping = [
            ("亏损部分的确认及转回", "亏损部分的确认及转回", "rgb(190, 190, 190)"),
            ("合同服务边际的释放", "合同服务边际的释放", "rgb(30, 73, 226)"),
            ("非金融风险调整的变动", "非金融风险调整的变动", "rgb(118, 210, 255)"),
            ("营运偏差及其他", "营运偏差及其他", "rgb(114, 19, 214)"),
            ("保费分配法业务净损益", "保费分配法业务净损益", "rgb(253, 52, 156)"),
            ("再保净损益", "再保净损益", "rgb(9, 142, 126)")
        ]
        
        fig = go.Figure()
        ct_list = list(filtered_results.keys())
        x_indices = list(range(len(ct_list)))
        
        dark_colors = {"rgb(30, 73, 226)", "rgb(114, 19, 214)", "rgb(9, 142, 126)"}
        show_labels = config.get('show_labels', True)
        label_size = config.get('label_size', 11)
        bar_width = config.get('bar_width', 0.4)
        
        # 收集数据
        all_data = {}
        for ct in ct_list:
            all_data[ct] = [filtered_results[ct]['contributions'].get(col_name, 0) for col_name, _, _ in display_mapping]
        
        # 🔥 扩大 Y 轴范围，容纳极端值（-500% 到 600%）
        all_positive_sums = [sum(v for v in all_data[ct] if v > 0) for ct in ct_list]
        all_negative_sums = [sum(v for v in all_data[ct] if v < 0) for ct in ct_list]
        y_max = max(max(all_positive_sums) + 50, 200) if all_positive_sums else 200
        y_min = min(min(all_negative_sums) - 50, -300) if all_negative_sums else -300
        
        # 画柱子
        for idx, (col_name, legend_name, color) in enumerate(display_mapping):
            y_vals = [all_data[ct][idx] for ct in ct_list]
            fig.add_trace(go.Bar(
                name=legend_name, x=x_indices, y=y_vals, width=bar_width,
                marker_color=color, text=None,
                hovertemplate="%{fullData.name}<br>%{y:.1f}%<extra></extra>"
            ))
        
        # 手动添加标签（强制显示，即使数值再小）
        if show_labels:
            for i, ct in enumerate(ct_list):
                values = all_data[ct]
                pos_cursor, neg_cursor = 0.0, 0.0
                
                for j, (col_name, legend_name, color) in enumerate(display_mapping):
                    v = values[j]
                    # 🔥 不再跳过小数值，全都显示
                    txt_color = "white" if color in dark_colors else "black"
                    
                    if v >= 0:
                        center_y = pos_cursor + v / 2
                        pos_cursor += v
                    else:
                        center_y = neg_cursor + v / 2
                        neg_cursor += v
                    
                    # 只显示绝对值 >= 1% 的标签，避免太拥挤
                    if abs(v) >= 1:
                        fig.add_annotation(
                            x=i, y=center_y, text=f"{v:.1f}%", showarrow=False,
                            font=dict(size=label_size, color=txt_color, weight='bold'),
                            xanchor="center", yanchor="middle", xref="x", yref="y"
                        )
        
        # 灰色背景框
        for i in range(len(ct_list)):
            fig.add_shape(type="rect", xref="x", yref="paper", x0=i - 0.46, x1=i + 0.46,
                          y0=-0.08, y1=1.08, fillcolor="rgba(200, 200, 200, 0.15)",
                          line=dict(color="rgba(150, 150, 150, 0.6)", width=1.5), layer="below")
        
        # Layout - 扩大 Y 轴范围
        fig.update_layout(
            barmode='relative', height=600, bargap=0.25,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=50, b=100, l=220, r=80),
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=-0.05,
                        traceorder="reversed", font=dict(size=11, color="#00338D")),
            yaxis=dict(side='right', showgrid=False, range=[y_min, y_max],
                       tickmode='array', 
                       tickvals=[-300, -200, -100, 0, 100, 200, 300, 400, 500, 600],
                       ticktext=["-300%", "-200%", "-100%", "0%", "100%", "200%", "300%", "400%", "500%", "600%"],
                       zeroline=True, zerolinecolor="#F7860C", zerolinewidth=2)
        )
        
        x_labels = [f"<span style='font-size:{config.get('co_font_size', 14)}px;color:#00338D;'><b>{ct}</b></span>" for ct in ct_list]
        fig.update_xaxes(showgrid=False, zeroline=False, tickvals=x_indices, ticktext=x_labels, side="top")
        
        return fig, None

    def create_expense_chart(results_by_year, year_list, selected_types, config):
        """绘制行业平均费用结构图"""
        from plotly.subplots import make_subplots
        import plotly.graph_objects as go
    
        first_year = list(results_by_year.keys())[0]
        available_types = list(results_by_year[first_year].keys())
        all_cts = [ct for ct in selected_types if ct in available_types]
    
        if not all_cts:
            all_cts = available_types
    
        expense_fields = ["获取费用", "维持费用", "非履约费用"]
        color_map = {
            "获取费用": "rgb(30, 73, 226)",
            "维持费用": "rgb(118, 210, 255)",
            "非履约费用": "rgb(114, 19, 234)"
        }
    
        fig = make_subplots(
            rows=1,
            cols=len(all_cts),
            shared_yaxes=True,
            subplot_titles=[f"<b>{ct}</b>" for ct in all_cts],
            horizontal_spacing=0.02
        )
    
        # 不再额外手工添加 scatter 图例，避免重复
        # 只让第一个子图的 bar 显示图例
    
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
    
            years_list = []
            for year in year_list:
                ratios = results_by_year.get(year, {}).get(ct, {}).get('ratios', {})
                years_list.append([ratios.get(f, 0) for f in expense_fields])
    
            x_vals = list(range(len(year_list)))
    
            for j, field in enumerate(expense_fields):
                y_vals = [data[j] for data in years_list]
                is_dark = field in ["获取费用", "非履约费用"]
    
                fig.add_trace(
                    go.Bar(
                        x=x_vals,
                        y=y_vals,
                        name=field,
                        marker_color=color_map[field],
                        text=[f"{v:.0f}%" if config['show_labels'] and v > 0 else "" for v in y_vals],
                        textposition='inside',
                        insidetextanchor='middle',
                        textfont=dict(
                            size=config['label_size'],
                            color="white" if is_dark else "#1a1a2e"
                        ),
                        width=config['bar_width'],
                        showlegend=(i == 0),   # 只在第一个子图显示图例
                        legendgroup=field,
                        hovertemplate="%{x}<br>%{fullData.name}: %{y:.1f}%<extra></extra>"
                    ),
                    row=1, col=col_idx
                )
    
            # 添加顶部总费用标注
            for year_idx, year in enumerate(year_list):
                total_val = results_by_year.get(year, {}).get(ct, {}).get('avg_total', 0) / 100
    
                fig.add_annotation(
                    x=year_idx,
                    y=105,
                    text=f"{total_val:.0f}",   # 不用<b>，避免被标题样式循环误伤
                    showarrow=False,
                    font=dict(size=config['label_size'], color="#222"),
                    bgcolor="white",
                    bordercolor="#BBB",
                    borderwidth=1,
                    xanchor="center",
                    yanchor="bottom",
                    xref=f"x{col_idx}" if col_idx > 1 else "x",
                    yref=f"y{col_idx}" if col_idx > 1 else "y"
                )
    
            # 灰色背景框
            fig.add_shape(
                type="rect",
                xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                yref="y domain" if col_idx == 1 else f"y{col_idx} domain",
                x0=-0.08, x1=1.08, y0=-0.1, y1=1.15,
                fillcolor="rgba(200, 200, 200, 0.15)",
                line=dict(color="rgba(150, 150, 150, 0.6)", width=1.5),
                layer="below",
                row=1, col=col_idx
            )
    
            fig.update_xaxes(
                tickvals=list(range(len(year_list))),
                ticktext=[f"{year}年" for year in year_list],
                showgrid=False,
                showline=False,
                zeroline=False,
                ticks="",
                row=1, col=col_idx
            )
    
        fig.update_layout(
            barmode='stack',
            bargap=0.25,
            height=550,
            margin=dict(t=50, b=80, l=10, r=10),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.15,
                xanchor="center",
                x=0.5,
                font=dict(size=12)
            )
        )
    
        # 单独设置每个子图的y轴
        for i in range(1, len(all_cts) + 1):
            fig.update_yaxes(
                range=[0, 115],
                showgrid=False,
                zeroline=False,
                tickformat=".0f",
                ticksuffix="%",
                row=1, col=i
            )
    
        # 只改 subplot title，不改总数 annotation
        title_count = len(all_cts)
        for idx in range(title_count):
            fig.layout.annotations[idx].update(
                y=1.05,
                font=dict(size=14, color="#00338D")
            )
    
        return fig
        
   
    def create_profit_mix_chart(results_by_year, year_list, selected_types, config):
        """绘制行业平均利润构成图（保险利润和投资利润的占比）"""
        from plotly.subplots import make_subplots
        
        all_cts = list(results_by_year[list(results_by_year.keys())[0]].keys())
        c = {'PI': '#FFD6EB', 'PS': '#00B8F5', 'CI': '#FD349C', 'CS': '#00338D'}
        
        fig = make_subplots(
            rows=2, cols=len(all_cts),
            row_heights=[0.18, 0.82],
            subplot_titles=[f"<span style='color:#00338D;'><b>{ct}</b></span>" for ct in all_cts],
            shared_yaxes=True,
            vertical_spacing=0.15,
            horizontal_spacing=0.03,
            specs=[[{"type": "domain"} for _ in all_cts], [{"type": "xy"} for _ in all_cts]]
        )
        
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            
            pi_prev = results_by_year.get(year_list[0], {}).get(ct, {}).get("投资利润", 0)
            ps_prev = results_by_year.get(year_list[0], {}).get(ct, {}).get("保险利润", 0)
            pi_curr = results_by_year.get(year_list[1], {}).get(ct, {}).get("投资利润", 0)
            ps_curr = results_by_year.get(year_list[1], {}).get(ct, {}).get("保险利润", 0)
            
            fig.add_trace(
                go.Pie(
                    labels=['投资', '保险'],
                    values=[pi_curr, ps_curr],
                    marker_colors=[c['CI'], c['CS']],
                    hole=0.75,
                    textinfo='none',
                    showlegend=False,
                    sort=False
                ),
                row=1, col=col_idx
            )
            
            if i == 0:
                for nm, cl in [(f"{year_list[0]}YE 投资利润", c['PI']),
                               (f"{year_list[0]}YE 保险利润", c['PS']),
                               (f"{year_list[1]}YE 投资利润", c['CI']),
                               (f"{year_list[1]}YE 保险利润", c['CS'])]:
                    fig.add_trace(
                        go.Scatter(
                            x=[None], y=[None], mode='markers',
                            marker=dict(color=cl, symbol='square', size=15),
                            name=nm, showlegend=True
                        ),
                        row=2, col=col_idx
                    )
            
            fig.add_trace(
                go.Bar(
                    x=[f"{year_list[0]}YE", f"{year_list[1]}YE"],
                    y=[pi_prev, pi_curr],
                    marker_color=[c['PI'], c['CI']],
                    text=[f"{pi_prev:.0f}", f"{pi_curr:.0f}"] if config['show_labels'] else None,
                    textposition='outside',
                    textfont=dict(size=12),
                    cliponaxis=False,
                    showlegend=False
                ),
                row=2, col=col_idx
            )
            
            fig.add_trace(
                go.Bar(
                    x=[f"{year_list[0]}YE", f"{year_list[1]}YE"],
                    y=[ps_prev, ps_curr],
                    marker_color=[c['PS'], c['CS']],
                    text=[f"{ps_prev:.0f}", f"{ps_curr:.0f}"] if config['show_labels'] else None,
                    textposition='outside',
                    textfont=dict(size=12),
                    cliponaxis=False,
                    showlegend=False
                ),
                row=2, col=col_idx
            )
        
        fig.update_layout(
            barmode='group',
            bargap=config['gap'],
            bargroupgap=0.0,
            height=500,
            margin=dict(t=50, b=110, l=40, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.15)
        )
        
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(type='category', showgrid=False, tickangle=0, tickfont=dict(color='gray'), row=2, col=i)
            fig.update_yaxes(showgrid=False, zeroline=True, zerolinecolor='lightgray', showticklabels=False, row=2, col=i)
        
        for ann in fig.layout.annotations:
            if "<b>" in str(ann.text):
                ann.update(y=1.08, font=dict(size=13, weight="bold"))
        
        return fig

    def create_oci_chart_industry(results_by_ct, year_list, selected_types, config, is_latest=True):
        """绘制行业平均OCI变动分析图"""
        from plotly.subplots import make_subplots
        
        all_cts = list(results_by_ct.keys())
        
        mc = {
            "净利润": {"c": "rgb(0,176,240)", "n": "净利润"},
            "可转损益OCI合计": {"c": "rgb(253,52,156)", "n": "可转损益OCI变动+FVOCI债券公允价值变动"},
            "不可转损益OCI合计": {"c": "rgb(114,19,234)", "n": "不可转损益负债OCI变动+FVOCI股权公允价值变动"},
            "other": {"c": "rgb(127,127,127)", "n": "其他"},
            "综合收益总额": {"c": "rgb(172,234,255)", "n": "综合收益总额"}
        }
        
        all_vals = []
        for ct in all_cts:
            data = results_by_ct[ct]
            for key in ["净利润", "可转损益OCI合计", "不可转损益OCI合计", "other", "综合收益总额"]:
                all_vals.append(data.get(key, 0))
        
        max_val = max(all_vals) if all_vals else 0
        min_val = min(all_vals) if all_vals else 0
        buffer = (max_val - min_val) * 0.8 if max_val != min_val else 100
        y_range = [min_val - buffer if min_val < 0 else min_val - abs(min_val) * 0.3, max_val + buffer]
        
        fig = make_subplots(
            rows=1, cols=len(all_cts),
            shared_yaxes=True,
            horizontal_spacing=0.015,
            subplot_titles=[f"<b><span style='color:#00338D;'>{ct}</span></b>" for ct in all_cts]
        )
        
        for col_idx, ct in enumerate(all_cts):
            data = results_by_ct[ct]
            
            for m_key, m_info in mc.items():
                val = data.get(m_key, 0)
                fig.add_trace(
                    go.Bar(
                        name=m_info["n"],
                        x=[m_key],
                        y=[val],
                        text=[f"{val:.0f}" if (config['show_labels'] and val != 0) else ""],
                        textposition='outside',
                        textfont=dict(size=11, color='#00338D'),
                        marker_color=m_info["c"],
                        width=0.8,
                        legendgroup=m_key,
                        showlegend=(col_idx == 0),
                        cliponaxis=False,
                        constraintext='none'
                    ),
                    row=1, col=col_idx + 1
                )
            
            fig.update_xaxes(showticklabels=False, showline=False, zeroline=False, showgrid=False, ticks="", ticklen=0, row=1, col=col_idx + 1)
        
        fig.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            barmode='group',
            bargap=config['bar_gap'],
            height=420,
            margin=dict(t=50, b=40, l=40, r=30),
            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=11))
        )
        fig.update_yaxes(range=y_range, showline=False, zeroline=False, showgrid=False, gridcolor='rgba(0,0,0,0)', gridwidth=0)
        
        for ann in fig.layout.annotations:
            ann.update(y=1.18, font_size=config['co_font_size'])
        
        return fig
    
    def create_oci_deep_chart(results_by_year, year_list, selected_types, config):
        """绘制行业平均OCI深度分析图（负债OCI和FVOCI债券公允价值）"""
        from plotly.subplots import make_subplots
        
        all_cts = list(results_by_year[list(results_by_year.keys())[0]].keys())
        colors = {"可转损益的负债OCI": "rgb(0, 184, 245)", "FVOCI债券公允价值": "rgb(253, 52, 156)"}
        
        all_vals = []
        for year in year_list:
            for ct in all_cts:
                data = results_by_year.get(year, {}).get(ct, {})
                all_vals.append(data.get("可转损益的负债OCI", 0))
                all_vals.append(data.get("FVOCI债券公允价值", 0))
        
        y_min = min(all_vals) if all_vals else 0
        y_max = max(all_vals) if all_vals else 0
        y_range = [y_min - (y_max - y_min) * 0.3 if y_min < 0 else y_min - abs(y_min) * 0.3,
                   y_max + (y_max - y_min) * 0.3] if all_vals else None
        
        fig = make_subplots(
            rows=1, cols=len(all_cts),
            shared_yaxes=True,
            horizontal_spacing=0.01,
            subplot_titles=[f"<b><span style='color:#00338D;'>{ct}</span></b>" for ct in all_cts]
        )
        
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_vals = [f"{year}YE" for year in year_list]
            
            for fn, color in colors.items():
                y_vals = []
                for year in year_list:
                    val = results_by_year.get(year, {}).get(ct, {}).get(fn, 0)
                    y_vals.append(val)
                
                fig.add_trace(
                    go.Bar(
                        name=fn, x=x_vals, y=y_vals,
                        text=[f"{x:.0f}" if config['show_labels'] and x != 0 else "" for x in y_vals],
                        textposition='outside',
                        textfont=dict(size=11, color='#00338D'),
                        marker_color=color,
                        legendgroup=fn,
                        showlegend=(i == 0),
                        width=0.4,
                        cliponaxis=False,
                        constraintext='none'
                    ),
                    row=1, col=col_idx
                )
            
            fig.update_xaxes(showline=False, zeroline=False, showgrid=False, ticks="", ticklen=0, row=1, col=col_idx)
        
        fig.update_layout(
            barmode='group',
            bargap=config['bar_gap'],
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            height=420,
            margin=dict(t=50, b=80, l=20, r=20),
            legend=dict(orientation="h", yanchor="top", y=-0.28, x=0.5, xanchor="center")
        )
        
        if y_range:
            fig.update_yaxes(range=y_range, showline=False, zeroline=True, zerolinecolor="#E0E0E0", zerolinewidth=1.02, showgrid=False)
        
        for ann in fig.layout.annotations:
            ann.update(y=1.18, font_size=config['co_font_size'])
        
        return fig
    
    def create_industry_asset_composition_chart(results_by_year, year_list, selected_types, config):
        """绘制行业平均资产构成图 - 完全内置配置"""
        from plotly.subplots import make_subplots
        import plotly.graph_objects as go
        
        # 内置配置
        COMPANY_TYPE_ORDER = ["头部", "银行系", "外资系", "养老健康", "小型"]
        ASSET_FIELD_MAP = {"AC": "债权投资", "FVOCI": "其他债权投资", "FVTPL": "交易性金融资产", "指定FVOCI": "其他权益工具投资"}
        ASSET_COLOR_MAP = {"AC": "rgb(0, 184, 245)", "FVOCI": "rgb(114, 19, 234)", "FVTPL": "rgb(253, 52, 156)", "指定FVOCI": "rgb(181, 2, 95)"}
        DARK_BG_FIELDS = {"FVOCI", "指定FVOCI", "AC"}
        
        first_year = list(results_by_year.keys())[0]
        all_cts = [ct for ct in COMPANY_TYPE_ORDER if ct in results_by_year[first_year].keys()]
        if not all_cts:
            return go.Figure()
        
        fig = make_subplots(rows=1, cols=len(all_cts), shared_yaxes=True,
                            column_titles=[f"<span style='color:#00338D;'><b>{ct}</b></span>" for ct in all_cts],
                            horizontal_spacing=0.015)
        
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_labels = [f"{year}YE" for year in year_list]
            year_data = {year: results_by_year.get(year, {}).get(ct, {}) for year in year_list}
            
            for fn, dn in ASSET_FIELD_MAP.items():
                y_vals = [year_data[year].get(fn, 0) for year in year_list]
                fig.add_trace(go.Bar(x=x_labels, y=y_vals, name=dn if i == 0 else None,
                                     marker_color=ASSET_COLOR_MAP[fn],
                                     text=[f"{v:.0f}%" if config['show_labels'] and v > 0 else "" for v in y_vals],
                                     textposition='inside', insidetextanchor='middle',
                                     textfont=dict(size=config['label_size'], color="white" if fn in DARK_BG_FIELDS else "black"),
                                     width=config['bar_width'], showlegend=(i == 0), legendgroup=fn, hoverinfo="skip"),
                              row=1, col=col_idx)
            
            fig.add_shape(type="rect", xref=f"x{col_idx} domain" if col_idx > 1 else "x domain",
                          yref="y domain", x0=-0.06, x1=1.06, y0=-0.12, y1=1.12,
                          fillcolor="rgba(0,0,0,0)", line=dict(color="rgba(0,0,0,0)", width=0),
                          layer="above", row=1, col=col_idx)
        
        fig.update_layout(barmode='stack', height=550, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                          margin=dict(t=50, b=120, l=40, r=40),
                          legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=11)))
        
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", row=1, col=i)
            fig.update_yaxes(showgrid=False, range=[0, 101], tickvals=[0, 25, 50, 75, 100],
                             ticktext=["0%", "25%", "50%", "75%", "100%"], showticklabels=(i == 1), row=1, col=i)
        
        for ann in fig.layout.annotations:
            ann.update(y=ann.y + 0.05, font_size=config['co_font_size'])
        
        return fig

    def create_industry_expense_composition_chart(results_by_year, year_list, selected_types, config):
        from plotly.subplots import make_subplots
        import plotly.graph_objects as go
    
        # 数据定义
        EXPENSE_FIELDS = [
            "保险获取现金流的摊销（保险服务费用）",
            "亏损部分的确认及转回",
            "当期发生的赔款及其他相关费用",
            "已发生赔款负债相关的履约现金流量变动"
        ]
    
        FIELD_DISPLAY = {
            "保险获取现金流的摊销（保险服务费用）": "保险获取现金流的摊销",
            "亏损部分的确认及转回": "亏损部分的确认及转回",
            "当期发生的赔款及其他相关费用": "当期发生的赔款及费用",
            "已发生赔款负债相关的履约现金流量变动": "已发生赔款负债变动"
        }
    
        COLOR_MAP = {
            "保险获取现金流的摊销（保险服务费用）": "rgb(30, 73, 226)",
            "亏损部分的确认及转回": "rgb(254, 174, 215)",
            "当期发生的赔款及其他相关费用": "rgb(0, 163, 161)",
            "已发生赔款负债相关的履约现金流量变动": "rgb(1, 184, 245)"
        }
    
        COMPANY_TYPE_ORDER = ["头部", "银行系", "外资系", "养老健康", "小型"]
    
        # 绘图区分：按顺序和选择类型
        all_cts = [ct for ct in COMPANY_TYPE_ORDER if ct in selected_types]
        other_cts = [ct for ct in selected_types if ct not in COMPANY_TYPE_ORDER]
        all_cts.extend(other_cts)
    
        if not all_cts:
            fig = go.Figure()
            fig.add_annotation(text="无有效数据", x=0.5, y=0.5, showarrow=False)
            return fig
    
        fig = make_subplots(
            rows=1, cols=len(all_cts),
            shared_yaxes=True,
            subplot_titles=[f"<b>{ct}</b>" for ct in all_cts],
            horizontal_spacing=0.02
        )
    
        # 堆叠柱 + 文本标签分离
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
    
            # 按公司类型提取数据
            y_vals_dict = {}
            for f in EXPENSE_FIELDS:
                y_vals_dict[f] = []
                for year in year_list:
                    val = results_by_year.get(year, {}).get(ct, {}).get(f, 0)
                    y_vals_dict[f].append(val)
    
            x_vals = list(range(len(year_list)))  # x 位置
            x_labels = [f"{year}年" for year in year_list]
    
            # 追踪堆叠位置
            pos_stack = [0] * len(year_list)
            neg_stack = [0] * len(year_list)
    
            for f_name in EXPENSE_FIELDS:
                y_vals = y_vals_dict[f_name]
    
                # 绘制柱子
                fig.add_trace(
                    go.Bar(
                        x=x_vals,
                        y=y_vals,
                        name=FIELD_DISPLAY[f_name],
                        marker_color=COLOR_MAP[f_name],
                        width=config['bar_width'],
                        showlegend=(i == 0),  # 图例只在第一列显示
                        legendgroup=f_name
                    ),
                    row=1,
                    col=col_idx
                )
    
                # 单独绘制文字标签
                text_positions = []
                for j, v in enumerate(y_vals):
                    if v >= 0:
                        text_positions.append(pos_stack[j] + v / 2)
                        pos_stack[j] += v
                    else:
                        text_positions.append(neg_stack[j] + v / 2)
                        neg_stack[j] += v
    
                fig.add_trace(
                    go.Scatter(
                        x=x_vals,
                        y=text_positions,
                        mode="text",
                        text=[f"{v:.0f}%" if config['show_labels'] else "" for v in y_vals],
                        textfont=dict(
                            size=config['label_size'],
                            color="white" if f_name in ["保险获取现金流的摊销（保险服务费用）", "亏损部分的确认及转回"] else "#1a1a2e",
                        ),
                        showlegend=False,  # 不显示图例
                    ),
                    row=1,
                    col=col_idx
                )
    
            # 添加灰色背景框
            fig.add_shape(
                type="rect",
                xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                yref="y domain",
                x0=-0.08,
                x1=1.08,
                y0=-0.1,
                y1=1.15,
                fillcolor="rgba(200, 200, 200, 0.15)",
                line=dict(color="rgba(150, 150, 150, 0.6)", width=1.5),
                layer="below",
                row=1,
                col=col_idx
            )
    
            # 添加X轴标签
            fig.update_xaxes(
                tickvals=x_vals,
                ticktext=x_labels,
                row=1,
                col=col_idx
            )
    
        # 计算Y轴范围（按堆叠柱子总高度）- 保持你的逻辑不变
        max_stack_height = []
        for year in year_list:
            for ct in all_cts:
                vals = [
                    results_by_year.get(year, {}).get(ct, {}).get(f, 0) for f in EXPENSE_FIELDS
                ]
                max_stack_height.append(sum(v for v in vals if v > 0))
        y_max = max(max_stack_height) * 1.15 if max_stack_height else 100
        fig.update_yaxes(
            range=[-10, y_max],
            showgrid=False,
            zeroline=True,
            zerolinecolor="#E0E0E0",
            ticksuffix="%",
            tickformat=".0f"
        )
    
        # ========== 修改1：添加橙色基准线（0%线），与 Step 7 一致 ==========
        fig.add_hline(y=0, line_color="orange", line_width=1.5, layer="below")
    
        # ========== 修改2：图例往下移，与 Step 7 一致 ==========
        fig.update_layout(
            barmode="relative",
            bargap=0.3,
            height=500,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=40, b=80, l=40, r=40),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.15,  # 从 -0.1 改成 -0.15
                xanchor="center",
                x=0.5,
                font=dict(size=10)
            )
        )
    
        return fig
   
    def create_industry_csm_composition_chart(results_by_ct, target_year, config):
        """
        绘制行业平均摊销前 CSM 构成图
        完全参照 Step 7 的 create_csm_composition_chart 样式
        """
        import plotly.graph_objects as go
        
        # 字段配置（与 Step 7 完全一致）
        CSM_FIELDS = ["新业务CSM（集团口径）", "CSM计息", "CSM调整"]
        FIELD_DISPLAY = {"新业务CSM（集团口径）": "新业务 CSM", "CSM计息": "CSM 计息", "CSM调整": "CSM 调整"}
        COLOR_MAP = {
            "新业务CSM（集团口径）": "rgb(0, 51, 140)",   # 深蓝色
            "CSM计息": "rgb(147, 157, 253)",            # 浅蓝色
            "CSM调整": "rgb(253, 52, 156)"              # 粉红色
        }
        COMPANY_TYPE_ORDER = ["头部", "银行系", "外资系", "养老健康", "小型"]
        
        # 获取公司类型列表（按固定顺序）
        all_cts = [ct for ct in COMPANY_TYPE_ORDER if ct in results_by_ct.keys()]
        other_cts = [ct for ct in results_by_ct.keys() if ct not in COMPANY_TYPE_ORDER]
        all_cts.extend(other_cts)
        
        if not all_cts:
            fig = go.Figure()
            fig.add_annotation(text="无有效数据", x=0.5, y=0.5, showarrow=False)
            return fig
        
        x_idx = list(range(len(all_cts)))
        tick_labels = [f"<span style='font-size:{config['co_font_size']}px; color:#00338D;'><b>{ct}</b></span>" for ct in all_cts]
        
        # 识别未披露的公司类型（所有字段占比都为0）
        no_data_cts = []
        for ct in all_cts:
            total = abs(results_by_ct[ct].get("新业务CSM（集团口径）", 0)) + \
                    abs(results_by_ct[ct].get("CSM计息", 0)) + \
                    abs(results_by_ct[ct].get("CSM调整", 0))
            if total == 0:
                no_data_cts.append(ct)
        
        no_data_x = [i for i, ct in enumerate(all_cts) if ct in no_data_cts]
        
        fig = go.Figure()
        
        # Step 7 风格：先画灰色占位柱（未披露的公司类型）
        if no_data_x:
            fig.add_trace(go.Bar(
                x=no_data_x, y=[100] * len(no_data_x),
                width=config['bar_width'],
                marker_color="#D9D9D9",
                showlegend=False,
                text=["未披露"] * len(no_data_x),
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(size=config['label_size'], color="white"),
                hoverinfo='skip'
            ))
        
        # 画每个字段的柱子（按顺序堆叠）
        for f_name in CSM_FIELDS:
            vals_pct = [results_by_ct[ct].get(f_name, 0) for ct in all_cts]
            # 未披露的公司强制设为0，不覆盖灰色占位柱
            vals_pct = [0 if ct in no_data_cts else v for ct, v in zip(all_cts, vals_pct)]
            
            # 文字颜色：深色背景用白色，浅色背景用黑色
            text_color = "white" if f_name == "新业务CSM（集团口径）" else "black"
            
            fig.add_trace(go.Bar(
                name=FIELD_DISPLAY[f_name],
                x=x_idx,
                y=vals_pct,
                width=config['bar_width'],
                marker_color=COLOR_MAP[f_name],
                text=[f"{v:.0f}%" if config['show_labels'] and abs(v) >= 1 else "" for v in vals_pct],
                textposition='inside',
                insidetextanchor='middle',
                textfont=dict(size=config['label_size'], color=text_color, weight='bold'),
                constraintext='none'
            ))
        
        # Step 7 布局：Y轴范围固定 -25% 到 105%
        fig.update_layout(
            barmode='relative',
            height=500,
            margin=dict(t=80, b=100, l=60, r=40),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5)
        )
        
        fig.update_xaxes(
            showgrid=False, zeroline=False,
            ticktext=tick_labels, tickvals=x_idx,
            ticks="", ticklen=0
        )
        
        # Step 7 Y轴刻度
        fig.update_yaxes(
            showgrid=False,
            range=[-25, 105],
            tickvals=[-20, 0, 20, 40, 60, 80, 100],
            ticktext=["-20%", "0%", "20%", "40%", "60%", "80%", "100%"],
            zeroline=True,
            zerolinecolor='#FFB07C',
            zerolinewidth=1.5
        )
        
        return fig

    # ==========================================
    # 第4层：辅助函数
    # ==========================================

    def show_chart(fig, p_mode):
        """统一图表输出"""
        if fig:
            if p_mode:
                fig.update_layout(width=850, autosize=False)
                st.plotly_chart(fig, use_container_width=False)
            else:
                st.plotly_chart(fig, use_container_width=True)

    def render_pure_chart_entity(m_id, print_mode):
        """纯图表实体渲染（计算 + 调用绘图函数）"""
        years = sorted([int(y) for y in df_raw['报告年份'].dropna().astype(str).str.replace(".0", "", regex=False).unique() if y.isdigit()])
        latest_year, prev_year = years[-1], years[-2] if len(years) > 1 else years[-1] - 1
        selected_types = st.session_state.get('step8_selected_types', [])
        
        # ==========================================
        # 统一处理所有散点图（使用 SCATTER_AXIS_META）
        # ==========================================
        if m_id in SCATTER_AXIS_META:
            meta = SCATTER_AXIS_META[m_id]
    
            df_scatter = calc_scatter_data(
                df_raw, selected_types, latest_year, prev_year,
                meta["x_field"], meta["title"]
            )
    
            if df_scatter.empty:
                st.warning(f"缺少 {meta['x_field']} 字段")
                return
    
            if not print_mode:
                # 非打印模式：渲染弹窗
                _, btn_col = st.columns([6, 1])
                with btn_col:
                    x_bins, y_bins = render_axis_bin_popover(
                        m_id=m_id,
                        title=meta["title"],
                        default_x_bins=parse_bins_input(meta["default_x"]),
                        default_y_bins=parse_bins_input(meta["default_y"])
                    )
            else:
                # 打印模式：直接用上传刻度或默认值
                uploaded_x, uploaded_y = get_uploaded_bins_by_mid(m_id)
                x_bins = uploaded_x or parse_bins_input(meta["default_x"])
                y_bins = uploaded_y or parse_bins_input(meta["default_y"])
    
            config = {
                "x_col": meta["x_field"],
                "y_col": meta["y_field"],
                "x_bins_custom": x_bins,
                "y_bins_custom": y_bins,
            }
    
            create_scatter_chart(df_scatter, config, meta["x_label"], meta["y_label"])
            return  # 散点图处理完直接返回
        # ==========================================
        # 统一处理所有堆叠分布图
        # ==========================================
        elif m_id in STACK_DIST_META:
            meta = STACK_DIST_META[m_id]
            
            # 🌟 1. 获取上传的配置
            uploaded_x, _ = get_uploaded_bins_by_mid(m_id)
            
            # 🌟 2. 设置 session_state 键名（和热力图一致）
            use_upload_key = f"dist_use_upload_{m_id}"
            x_text_key = f"dist_x_bins_text_{m_id}"
            
            # 🌟 3. 初始化 session_state（和热力图一致）
            if use_upload_key not in st.session_state:
                st.session_state[use_upload_key] = bool(uploaded_x)  # 有上传配置时默认勾选
            if x_text_key not in st.session_state:
                st.session_state[x_text_key] = bins_to_str(uploaded_x) if uploaded_x else meta["default_x"]
            
            if not print_mode:
                # 🌟 4. 参数按钮放在右上角（和热力图一样的位置）
                _, btn_col = st.columns([6, 1])
                with btn_col:
                    with st.popover(f"⚙️ {meta['name']} 参数设置", use_container_width=True):
                        show_labels = st.toggle("显示数值标签", value=True, key=f"lab_{m_id}")
                        label_size = st.slider("标签大小", 8, 16, 11, key=f"sz_{m_id}")
                        
                        st.markdown("---")
                        
                        # 🌟 5. 上传配置优先复选框（和热力图完全一致）
                        if uploaded_x:
                            st.checkbox("优先使用上传 Excel 的刻度", key=use_upload_key)
                            st.caption(f"已上传：X = {bins_to_str(uploaded_x)}")
                        else:
                            st.caption("未检测到上传的刻度配置")
                        
                        # 🌟 6. 手动输入区间（和热力图完全一致）
                        st.text_input(
                            "X 轴区间（英文逗号分隔）", 
                            key=x_text_key, 
                            placeholder=meta["default_x"]
                        )
                        
                        # 🌟 7. 恢复默认按钮（和热力图完全一致）
                        if st.button("恢复默认刻度", key=f"reset_dist_{m_id}", use_container_width=True):
                            st.session_state[use_upload_key] = bool(uploaded_x)
                            st.session_state[x_text_key] = meta["default_x"]
                            st.rerun()
                
                # 🌟 8. 根据复选框决定使用哪个配置（和热力图完全一致）
                if st.session_state[use_upload_key] and uploaded_x:
                    x_bins = uploaded_x
                else:
                    x_bins = parse_bins_input_safe(st.session_state[x_text_key], parse_bins_input(meta["default_x"]))
            else:
                # 打印模式：优先使用上传配置
                if uploaded_x:
                    x_bins = uploaded_x
                else:
                    x_bins = parse_bins_input(meta["default_x"])
                show_labels, label_size = True, 11
            
            # 调用计算函数
            distribution_df, labels = calc_stack_distribution(
                df_raw, selected_types, latest_year, 
                meta["calc_func"], 
                x_bins_custom=x_bins
            )
            
            if distribution_df.empty or not labels:
                st.warning(f"无法计算 {meta['name']} 分布，数据不足")
                return
            
            create_stack_chart_and_table(
                distribution_df, labels, meta['name'], 
                latest_year, show_labels, label_size
            )
                
        # 图表1：非PAA和PAA合同组占比分析
        elif m_id == "industry_comp_1":
            fields = ["采用保费分配法计量的保险合同保险服务收入", "未采用保费分配法计量的保险合同保险服务收入"]
            
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_two_cat_composition(df_raw, ct, year, fields)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签大小", 5, 20, 12, key=f"lsz_{m_id}")
                with c3: bar_width = st.slider("柱子宽度", 0.1, 1.0, 0.6, 0.05, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司名称大小", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = True, 12, 0.6, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig = create_composition_chart_v1(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表2：非PAA合同组收入构成分析
        elif m_id == "industry_comp_2":
            field_map = {
                "合同服务边际的摊销": "合同服务边际的释放",
                "非金融风险调整的变动": "非金融风险调整的变动",
                "预计当期发生的保险服务费用": "预期当期发生的保险服务费用",
                "保险获取现金流的摊销（保险服务收入）": "保险获取现金流的摊销",
                "与当期服务或过去服务相关得保费经验调整": "与当期服务或过去服务相关的保费经验调整",
                "其他收入调整": "其他"
            }
            color_map = {
                "合同服务边际的摊销": "rgb(30, 73, 226)",
                "非金融风险调整的变动": "rgb(254, 174, 215)",
                "预计当期发生的保险服务费用": "rgb(0, 163, 161)",
                "保险获取现金流的摊销（保险服务收入）": "rgb(1, 184, 245)",
                "与当期服务或过去服务相关得保费经验调整": "rgb(0, 219, 214)",
                "其他收入调整": "rgb(114, 19, 234)"
            }
            
            denominator_field = "未采用保费分配法计量的保险合同保险服务收入"
            numerator_fields = list(field_map.keys())
            
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_composition(df_raw, ct, year, numerator_fields, denominator_field)
                    if result:
                        display_result = {field_map[k]: v for k, v in result.items()}
                        year_results[ct] = display_result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签大小", 5, 20, 12, key=f"lsz_{m_id}")
                with c3: bar_width = st.slider("柱子宽度", 0.1, 1.0, 0.6, 0.05, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司名称大小", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = True, 12, 0.6, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig, df_avg = create_composition_chart_v2(results_by_year, year_list, selected_types, config)
            
            # 输出平均值表格
            if fig and not df_avg.empty:
                st.markdown("<div style='font-size: 13px; font-weight: bold; margin-bottom: 8px; color:#333;'>各公司平均占比情况 (样本均值)</div>", unsafe_allow_html=True)
                
                html = "<table style='width:100%; border-collapse: collapse; font-family: sans-serif; font-size: 11px; margin-bottom: 10px;'>"
                html += "<tr style='background-color: #00338D; color: white; text-align: center; font-weight: bold;'>"
                html += "<th style='padding: 4px 6px; border: 1px solid white;'>报告年份</th>"
                for orig_k, display_name in field_map.items():
                    html += f"<th style='padding: 4px 6px; background-color: {color_map[orig_k]}; color: white; border: 1px solid white;'>{display_name}</th>"
                html += "</tr>"
                
                for year in year_list:
                    year_str = f"{year}年"
                    html += f"<tr><td style='padding: 4px 6px; font-weight: bold; background-color: #F8F9FA; border: 1px solid #EAEAEA;'>{year_str}</td>"
                    for comp in field_map.values():
                        val = df_avg.loc[year_str, comp]
                        html += f"<td style='padding: 4px 6px; text-align: center; background-color: white; border: 1px solid #EAEAEA;'>{val:.1f}%</td>"
                    html += "</tr>"
                
                html += "</table>"
                st.markdown(html, unsafe_allow_html=True)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
        #保险服务费用构成
        elif m_id == "industry_exp_1":
            # 收集数据
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_expense_composition(df_raw, ct, year)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均费用构成")
                return
            
            # UI 配置（与 Step 7 exp_1 风格一致）
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1:
                    show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2:
                    bar_width = st.slider("柱子宽度", 0.2, 0.8, 0.6, key=f"wid_{m_id}")
                with c3:
                    label_size = st.slider("标签大小", 8, 16, 11, key=f"sz_{m_id}")
            else:
                show_labels, bar_width, label_size = True, 0.6, 11
            
            config = {
                'show_labels': show_labels,
                'bar_width': bar_width,
                'label_size': label_size,
                'co_font_size': 12
            }
            
            fig = create_industry_expense_composition_chart(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
        # 业务及管理费（新准则分类）
        elif m_id == "industry_exp_struct":
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_expense_structure(df_raw, ct, year)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均费用结构")
                return
            
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2: bar_width = st.slider("柱子宽度", 0.2, 0.8, 0.6, key=f"wid_{m_id}")
                with c3: label_size = st.slider("标签大小", 8, 20, 10, key=f"sz_{m_id}")
            else:
                show_labels, bar_width, label_size = True, 0.6, 10
            
            config = {'show_labels': show_labels, 'bar_width': bar_width, 'label_size': label_size}
            
            fig = create_expense_chart(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)，顶部数字为总费用（亿元）</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表4：行业平均最新年保险服务业绩构成
        elif m_id == "industry_prof_2025":
            target_year = latest_year
            
            all_excluded = {}  # 收集所有类型的剔除记录
            results_by_ct = {}
                 
            
            for ct in selected_types:
                result, excluded = calc_industry_profit_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
                    if excluded:
                        all_excluded[ct] = excluded
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均利润构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示利润标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签字号", 8, 16, 11, key=f"psz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.4, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司字号", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = True, 11, 0.4, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 
                      'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig, excluded_html = create_profit_composition_chart(results_by_ct, config, all_excluded)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
            
            # 显示剔除说明
            if excluded_html and not print_mode:
                st.markdown(excluded_html, unsafe_allow_html=True)
        
        # 图表5：行业平均上一年保险服务业绩构成（同样的修改）
        elif m_id == "industry_prof_2024":
            target_year = prev_year
            
            all_excluded = {}
            results_by_ct = {}
            
            for ct in selected_types:
                result, excluded = calc_industry_profit_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
                    if excluded:
                        all_excluded[ct] = excluded
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均利润构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示利润标签", value=True, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签字号", 8, 16, 11, key=f"psz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.4, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司字号", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = True, 11, 0.4, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 
                      'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig, excluded_html = create_profit_composition_chart(results_by_ct, config, all_excluded)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
            
            if excluded_html and not print_mode:
                st.markdown(excluded_html, unsafe_allow_html=True)
       
        # 图表6：行业平均利润构成（保险利润 vs 投资利润）
        elif m_id == "industry_prof_mix":
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_profit_mix(df_raw, ct, year)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均利润构成")
                return
            
            if not print_mode:
                c1, c2 = st.columns([1, 2])
                with c1: show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2: gap = st.slider("柱子间距", 0.2, 0.7, 0.4, key=f"gap_{m_id}")
            else:
                show_labels, gap = True, 0.4
            
            config = {'show_labels': show_labels, 'gap': gap}
            
            fig = create_profit_mix_chart(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百万元人民币</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表7：行业平均OCI深度分析
        elif m_id == "industry_oci_deep":
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_oci_deep(df_raw, ct, year)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均OCI深度分析")
                return
            
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: bar_gap = st.slider("间距", 0.05, 0.5, 0.15, key=f"gap_{m_id}")
                with c3: co_font_size = st.slider("公司字号", 8, 20, 12, key=f"sz_{m_id}")
            else:
                show_labels, bar_gap, co_font_size = True, 0.15, 12
            
            config = {'show_labels': show_labels, 'bar_gap': bar_gap, 'co_font_size': co_font_size}
            
            fig = create_oci_deep_chart(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百万元人民币</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表8：行业平均最新年综合收益变动情况（OCI）
        elif m_id == "industry_oci_year_lat":
            target_year = latest_year
            
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_oci_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均OCI构成")
                return
            
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: bar_gap = st.slider("间距", 0.05, 0.5, 0.15, key=f"gap_{m_id}")
                with c3: co_font_size = st.slider("公司字号", 8, 20, 12, key=f"sz_{m_id}")
            else:
                show_labels, bar_gap, co_font_size = True, 0.15, 12
            
            config = {'show_labels': show_labels, 'bar_gap': bar_gap, 'co_font_size': co_font_size}
            
            fig = create_oci_chart_industry(results_by_ct, year_list, selected_types, config, is_latest=True)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百万元人民币</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表9：行业平均上一年综合收益变动情况（OCI）
        elif m_id == "industry_oci_year_pre":
            target_year = prev_year
            
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_oci_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均OCI构成")
                return
            
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1: show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2: bar_gap = st.slider("间距", 0.05, 0.5, 0.15, key=f"gap_{m_id}")
                with c3: co_font_size = st.slider("公司字号", 8, 20, 12, key=f"sz_{m_id}")
            else:
                show_labels, bar_gap, co_font_size = True, 0.15, 12
            
            config = {'show_labels': show_labels, 'bar_gap': bar_gap, 'co_font_size': co_font_size}
            
            fig = create_oci_chart_industry(results_by_ct, year_list, selected_types, config, is_latest=False)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百万元人民币</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
        
        #IFRS9 资产端分类        
        elif m_id == "industry_asset_struct":
            # 收集数据
            results_by_year = {}
            for year in year_list:
                year_results = {}
                for ct in selected_types:
                    result = calc_industry_asset_composition(df_raw, ct, year)
                    if result:
                        year_results[ct] = result
                if year_results:
                    results_by_year[year] = year_results
            
            if not results_by_year:
                st.warning("无法计算行业平均资产构成")
                return
            
            # ========== 添加以下代码 ==========
            # UI 配置
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1:
                    show_labels = st.toggle("显示数据标签", value=True, key=f"lab_{m_id}")
                with c2:
                    bar_width = st.slider("柱子宽度", 0.1, 1.0, 0.6, key=f"wid_{m_id}")
                with c3:
                    label_size = st.slider("标签大小", 8, 16, 11, key=f"sz_{m_id}")
            else:
                show_labels, bar_width, label_size = True, 0.6, 11
            
            config = {
                'show_labels': show_labels,
                'bar_width': bar_width,
                'label_size': label_size,
                'co_font_size': 12
            }
            
            fig = create_industry_asset_composition_chart(results_by_year, year_list, selected_types, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
       
        #最新年摊销前 CSM 变动项占比  
        elif m_id == "industry_csm_comp_lat":
            target_year = latest_year
            
            # 收集各公司类型的平均构成数据
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_csm_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均摊销前 CSM 构成")
                return
            
            # UI 配置
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1:
                    show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2:
                    bar_width = st.slider("柱宽", 0.1, 1.0, 0.5, key=f"wid_{m_id}")
                with c3:
                    label_size = st.slider("字号", 8, 20, 11, key=f"sz_{m_id}")
            else:
                show_labels, bar_width, label_size = True, 0.5, 11
            
            config = {
                'show_labels': show_labels,
                'bar_width': bar_width,
                'label_size': label_size,
                'co_font_size': 14
            }
            
            fig = create_industry_csm_composition_chart(results_by_ct, target_year, config)
            
            st.markdown(f"<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%) | 数据年份：{target_year}年</p>", 
                        unsafe_allow_html=True)
            show_chart(fig, print_mode)
        # 上一年摊销前 CSM 变动项占比
        elif m_id == "industry_csm_comp_pre":
            target_year = prev_year  # 只用改这一行：latest_year → prev_year
            
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_csm_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均摊销前 CSM 构成")
                return
            
            if not print_mode:
                c1, c2, c3 = st.columns(3)
                with c1:
                    show_labels = st.toggle("显示标签", value=True, key=f"lab_{m_id}")
                with c2:
                    bar_width = st.slider("柱宽", 0.1, 1.0, 0.5, key=f"wid_{m_id}")
                with c3:
                    label_size = st.slider("字号", 8, 20, 11, key=f"sz_{m_id}")
            else:
                show_labels, bar_width, label_size = True, 0.5, 11
            
            config = {
                'show_labels': show_labels,
                'bar_width': bar_width,
                'label_size': label_size,
                'co_font_size': 14
            }
            
            fig = create_industry_csm_composition_chart(results_by_ct, target_year, config)
            
            st.markdown(f"<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%) | 数据年份：{target_year}年</p>", 
                        unsafe_allow_html=True)
            show_chart(fig, print_mode)
        else:
            st.info(f"⏳ 模块 [{m_id}] 尚未配置底层绘图代码")
            
    # ==========================================
    # 第5层：报告包装器
    # ==========================================

    def render_report_module(m_id, print_mode, is_first=False):
        """完整报告模块渲染（标题 + 分析框 + 图表 + 注释）- 纯净版，无动态单位，无AI"""
        mod_data = notes_dict_8.get(m_id, {})
        
        full_title = mod_data.get('title', m_id)
        if 'df_notes' in st.session_state and isinstance(st.session_state['df_notes'], pd.DataFrame):
            df_n = st.session_state['df_notes']
            if '模块ID' in df_n.columns:
                match = df_n[df_n['模块ID'] == m_id]
                if not match.empty:
                    r = match.iloc[0]
                    title_parts = []
                    for field in ['一级分类', '二级分类', '对应图表名称']:
                        val = str(r.get(field, '')).strip()
                        if val and val.lower() != 'nan' and val != '全部':
                            title_parts.append(val)
                    if title_parts:
                        full_title = " - ".join(title_parts)

        # ====== 打印容器 & 网页分割线 ======
        if print_mode:
            st.markdown("<div class='page-break-container' style='margin:0;padding:0;'>", unsafe_allow_html=True)
        if not print_mode:
            st.markdown(
                "<div class='no-print' style='height:2px; background:linear-gradient(to right, #00338D, #005EBB, #FFFFFF); margin-bottom:15px;'></div>",
                unsafe_allow_html=True
            )   
            
        # ====== 第 1 步：大标题 ======
        title_cls = "page-break-title" if (print_mode and not is_first) else ""
        mt = "0px" if (print_mode and is_first) else "20px"
        font_size = "35px" if print_mode else "30px"
        st.markdown(
            f"<h3 class='{title_cls}' style='"
            f"text-align:left; color:#00338D; font-size:{font_size}; font-weight:900; "
            f"font-family:Microsoft YaHei, 微软雅黑, sans-serif; "
            f"margin-top:{mt}; margin-bottom:20px; border:none; padding-bottom:0px;'>"
            f"{full_title}</h3>",
            unsafe_allow_html=True
        )

        # ==========================================
        # ====== 第 2 步：分析内容框 (🌟 重点修复区域) ======
        # ==========================================
        
        # 1. 定义超级清洗器：把 pd.NA, None, 'nan', 'null', 空白 全部杀干净
        def clean_note(val):
            if pd.isna(val): return ""
            val_str = str(val).strip()
            if val_str.lower() in ['nan', 'none', 'null', '']: return ""
            return val_str
            
        analysis_default = clean_note(mod_data.get('analysis_default', ''))
        analysis_custom = clean_note(mod_data.get('analysis_custom', ''))
        
        # 2. 只有当两者至少有一个有真实文字时，才渲染这个带底色的框！
        if analysis_default or analysis_custom:
            # 🌟 完美复刻图1：浅灰蓝底色 + KPMG深蓝左侧粗边框
            html = '<div style="background-color:#F4F7FC; border-left:4px solid #00338D; padding:5px 5px; margin-bottom:5px; text-align:left; border-radius:3px;">'
            
            if analysis_default:
                # 默认内容：深蓝色 (#0A1F5C)
                html += f'<p style="margin:0; color:#0A1F5C; font-size:12px; line-height:1.4;">{analysis_default}</p>'
                
            if analysis_custom:
                # 自定义内容：亮蓝色 (#1E49E2)，加粗
                mt_space = "4px" if analysis_default else "0px" # 如果上面有默认文字，中间空出8像素距
                html += f'<p style="margin:{mt_space} 0 0 0; color:#002678; font-size:13px; line-height:1.4;">{analysis_custom}</p>'
                
            html += '</div>'
            st.markdown(html, unsafe_allow_html=True)


        # ====== 第 3 步：图表与顶部单位提示 ======
        is_pct_chart = any(x in m_id for x in ["comp", "ratio", "margin", "csm_trans", "struct"]) or m_id == "asset_struct"
        
        if print_mode:
            unit_text = "百分比 (%)" if is_pct_chart else "百万人民币"
            st.markdown(f"<p style='text-align:right; font-size:11px; margin-bottom:2px; color:#666;'>单位：{unit_text}</p>", unsafe_allow_html=True)
            render_pure_chart_entity(m_id, print_mode)
        else:
            chart_col_left, chart_col_center, chart_col_right = st.columns([1, 10, 1])
            with chart_col_center:
                unit_text = "百分比 (%)" if is_pct_chart else "百万人民币"
                st.markdown(f"<p style='text-align:right; font-size:12px; margin-bottom:2px; color:#666;'>单位：{unit_text}</p>", unsafe_allow_html=True)
                render_pure_chart_entity(m_id, print_mode)


        # ====== 第 4 步：底部注释 ======
        # 顺手把底部的注释也用清洗器过一遍，防止底部跑出 nan
        note_text = clean_note(mod_data.get('note', '')) 
        if note_text:
            st.markdown(
                f'<div style="margin-top:10px; margin-bottom:20px; text-align:left;">'
                f'<p style="margin:0; color:#888; font-size:12px; font-style:italic; line-height:1.4;">注：{note_text}</p>'
                f'</div>',
                unsafe_allow_html=True
            )

        if print_mode:
            st.markdown("</div>", unsafe_allow_html=True)

    # ==========================================
    # 🌐 网页模式 / 🖨️ 打印模式 的最终执行器 (带封面与封底)
    # ==========================================
    if not print_mode:
        st.markdown(
            "<hr class='no-print' style='border:none;border-top:1px solid #EAEAEA;margin:10px 0;'>",
            unsafe_allow_html=True
        )
    if print_mode:
        if 'ordered_modules' not in locals() or not ordered_modules:
            st.warning("⚠️ 报告顺序由【模块ID】的先后顺序决定，请先在上方传入有模块ID的注释表文件。")
        else:
            import datetime
            today = datetime.date.today()
            date_str = f"{today.year}年{today.month}月"
            # 🌟 防止 selected_type 报错，用 session_state 安全获取
            type_str = st.session_state.get('selected_type', '')
            if type_str == "全部": 
                type_str = ""
            # 🌟 防止 latest_year 报错
            ly_str = st.session_state.get('latest_year', '2025')
            
            cover_url = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/%E6%A0%87%E9%A2%98%E9%A1%B5.png"
            back_url  = "https://raw.githubusercontent.com/z-xylym/my-actuary-tool/main/%E5%B0%81%E5%BA%95%E9%A1%B5.png"
            
            # ✅ 封面
            st.markdown(f"""
            <div style="position:relative; width:338.67mm;  height:175.5mm;    
                page-break-after:always; overflow:hidden; margin:0; padding:0;
                -webkit-print-color-adjust:exact; print-color-adjust:exact; forced-color-adjust:none;">
                <img src="{cover_url}" style="width:100%; height:100%; object-fit:cover; display:block;"/>
                <div style="position:absolute; top:0; left:0; width:100%; height:100%;
                    display:flex; flex-direction:column; justify-content:center; align-items:flex-start;
                    padding:0 8%; box-sizing:border-box; margin-top:-30px; z-index:10;
                    forced-color-adjust:none; -webkit-print-color-adjust:exact; print-color-adjust:exact;">
                    <div style="font-size:52px; font-weight:900; line-height:1.4; margin-bottom:16px;
                        font-family:Microsoft YaHei,微软雅黑,sans-serif;
                        color:white; -webkit-text-fill-color:white;
                        text-shadow:2px 2px 4px rgba(0,0,0,0.5), 0 0 20px rgba(0,0,0,0.3);
                        forced-color-adjust:none; -webkit-print-color-adjust:exact;">
                        {ly_str}年新会计准则行业表现和洞察<br>{type_str}保险公司<br>
                    </div>
                    <div style="font-size:22px; font-weight:500; margin:0;
                        font-family:Microsoft YaHei,微软雅黑,sans-serif;
                        color:white; -webkit-text-fill-color:white;
                        text-shadow:1px 1px 3px rgba(0,0,0,0.5);
                        forced-color-adjust:none; -webkit-print-color-adjust:exact;">{date_str}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # --------渲染图表和页码-------
            for i, mod in enumerate(ordered_modules):
                render_report_module(mod, print_mode=True, is_first=(i == 0))
                
            # ✅ 封底
            st.markdown(f"""
            <div style="position:relative; width:338.67mm;  height:175.5mm;
                page-break-before:always; overflow:hidden; margin:0; padding:0;
                -webkit-print-color-adjust:exact; print-color-adjust:exact; forced-color-adjust:none;">
                <img src="{back_url}" style="width:100%; height:100%; object-fit:cover; display:block;"/>
            </div>
            """, unsafe_allow_html=True)
    else:
        if active_m_id:
            render_report_module(active_m_id, print_mode=False, is_first=True)
        else:
            st.info("💡 请从左侧导航栏选择要查看的行业分析模块")

