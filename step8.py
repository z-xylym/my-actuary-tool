# -*- coding: utf-8 -*-
"""
Created on Tue Jun  2 15:41:16 2026

@author: 34872
"""

def show_step_8_content():
    """行业统计分析 - 完全参照 Step 7 的三层架构"""

    # ==========================================
    # 第1层：样式注入与数据准备
    # ==========================================
    
    # 1.1 样式注入
    st.markdown("""
        <style>
        .industry-analysis-box { background: #F0F4FA; border-left: 4px solid #00338D; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        @media print {
            .nav-floating-sign, [data-testid="collapsedControl"], header, footer,
            [data-testid="stSidebar"], button, .stExpander, hr { display: none !important; }
            h2, h3 { color: #00338D !important; }
            .plotly-graph-div { page-break-inside: avoid !important; }
        }
        </style>
    """, unsafe_allow_html=True)
    
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
    # 2. 图表配置定义
    # ==========================================
    
    # 散点热力图配置：模块ID -> 配置
    scatter_config_map = {
        "industry_inc_total": {
            "prefix": "revenue_scatter",
            "name": "保险服务收入",
            "x_default": "0,100,200,500,1000,2000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间边界（保险服务收入，亿元）",
            "y_help": "Y轴区间边界（保险服务收入增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,100,200,500,1000,2000，对应区间为 (0,100]、(100,200] ..."
        },
        "industry_inv_return": {
            "prefix": "inv_return_scatter",
            "name": "净投资回报",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间边界（净投资回报，亿元）",
            "y_help": "Y轴区间边界（净投资回报增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000，对应区间为 (0,50]、(50,100] ..."
        },
        "industry_exp_total": {
            "prefix": "expense_total_scatter",
            "name": "保险服务费用",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间边界（保险服务费用，亿元）",
            "y_help": "Y轴区间边界（保险服务费用增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000，对应区间为 (0,50]、(50,100] ..."
        },
        "industry_uw_profit": {
            "prefix": "uw_profit_scatter",
            "name": "承保财务净损益",
            "x_default": "-1000,-500,-200,-100,0,100,200,500",
            "y_default": "-1000,-500,-200,-100,0,100,200,500",
            "x_help": "X轴区间边界（承保财务净损益，亿元）",
            "y_help": "Y轴区间边界（承保财务净损益增长率，%）",
            "caption": "请注意该指标可能为负数，请输入包含负值的区间边界。例如：-1000,-500,-200,-100,0,100,200,500"
        },
        "industry_perf_total": {
            "prefix": "perf_total_scatter",
            "name": "保险服务业绩",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间（保险服务业绩，亿元）",
            "y_help": "Y轴区间（保险服务业绩增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_inv_profit": {
            "prefix": "inv_profit_scatter",
            "name": "投资利润",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间（投资利润，亿元）",
            "y_help": "Y轴区间（投资利润增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_net_profit": {
            "prefix": "net_profit_scatter",
            "name": "净利润",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间边界（净利润，亿元）",
            "y_help": "Y轴区间边界（净利润增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_tax_profit": {
        "prefix": "tax_profit_scatter",
        "name": "税前利润",
        "x_default": "0,50,100,200,500,1000",
        "y_default": "-100,0,100,200,500,1000",
        "x_help": "X轴区间边界（税前利润，亿元）",
        "y_help": "Y轴区间边界（税前利润增长率，%）",
        "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_oci_profit": {
            "prefix": "oci_profit_scatter",
            "name": "其他综合收益",
            "x_default": "-500,-200,-100,0,100,200,500",
            "y_default": "-1000,-500,-200,0,200,500,1000",
            "x_help": "X轴区间边界（其他综合收益，亿元）",
            "y_help": "Y轴区间边界（其他综合收益增长率，%）",
            "caption": "请注意：其他综合收益可能为负数，请输入包含负值的区间边界。例如：-500,-200,-100,0,100,200,500"
        },
        "industry_total_profit": {
            "prefix": "total_profit_scatter",
            "name": "综合收益总额",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,0,100,200,500,1000",
            "x_help": "X轴区间边界（综合收益总额，亿元）",
            "y_help": "Y轴区间边界（综合收益总额增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_csm_bal": {
            "prefix": "csm_bal_scatter",
            "name": "CSM余额",
            "x_default": "0,100,500,1000,2000,5000",
            "y_default": "-50,-20,0,20,50,100,200",
            "x_help": "X轴区间边界（CSM余额，亿元）",
            "y_help": "Y轴区间边界（CSM余额增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,100,500,1000,2000,5000"
        },
        "industry_equity_trend": {
            "prefix": "equity_trend_scatter",
            "name": "净资产",
            "x_default": "0,100,500,1000,2000,5000",
            "y_default": "-50,-20,0,20,50,100,200",
            "x_help": "X轴区间边界（净资产，亿元）",
            "y_help": "Y轴区间边界（净资产增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,100,500,1000,2000,5000"
        },
        "industry_csm_equity": {
            "prefix": "csm_equity_scatter",
            "name": "CSM/净资产占比",
            "x_default": "0,20,40,60,80,100,120",
            "y_default": "-20,-10,-5,0,5,10,20",
            "x_help": "X轴区间边界（CSM/净资产占比，%）",
            "y_help": "Y轴区间边界（占比变化，百分点）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,20,40,60,80,100,120"
        },
        "industry_nb_csm": {
            "prefix": "nb_csm_scatter",
            "name": "新业务CSM",
            "x_default": "0,50,100,200,500,1000",
            "y_default": "-100,-50,-20,0,20,50,100",
            "x_help": "X轴区间边界（新业务CSM，亿元）",
            "y_help": "Y轴区间边界（新业务CSM增长率，%）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,50,100,200,500,1000"
        },
        "industry_lc_loss_ratio": {
            "prefix": "lc_loss_ratio_scatter",
            "name": "新业务LC亏损率",
            "x_default": "0,10,20,30,40,50,60,70,80,90,100",
            "y_default": "-50,-30,-20,-10,0,10,20,30,50",
            "x_help": "X轴区间边界（LC亏损率，%）",
            "y_help": "Y轴区间边界（LC亏损率变化，百分点）",
            "caption": "请输入区间边界值，用英文逗号分隔。例如：0,10,20,30,40,50,60,70,80,90,100"
        },
    }
    
    # 比率堆叠分布图配置：模块ID -> 配置
    stack_dist_config_map = {
        "industry_csm_ratio": {
            "name": "CSM/BEL占比",
            "default": "0,20,40,60,80,100,120",
            "caption": "X轴为比率区间，Y轴为公司数量"
        },
        "industry_csm_continuity_ratio": {
            "name": "CSM持续率",
            "default": "0,50,100,150,200,300",
            "caption": "X轴为比率区间，Y轴为公司数量"
        },
        "industry_csm_amort_ratio": {
            "name": "CSM摊销比率",
            "default": "0,5,10,15,20,30",
            "caption": "X轴为比率区间，Y轴为公司数量"
        },
        "industry_nb_margin_trend": {
            "name": "新业务IFRS利润率",
            "default": "0,10,20,30,40,50",
            "caption": "X轴为比率区间，Y轴为公司数量"
        },
    }
    
    # ==========================================
    # 3. 公共函数
    # ==========================================
    
    def init_scatter_state(cfg):
        prefix = cfg["prefix"]
        if f"{prefix}_use_custom_bins" not in st.session_state:
            st.session_state[f"{prefix}_use_custom_bins"] = False
        if f"{prefix}_x_bins" not in st.session_state:
            st.session_state[f"{prefix}_x_bins"] = cfg["x_default"]
        if f"{prefix}_y_bins" not in st.session_state:
            st.session_state[f"{prefix}_y_bins"] = cfg["y_default"]
    
    def init_stack_dist_state(m_id, cfg):
        if f"{m_id}_dist_use_custom_bins" not in st.session_state:
            st.session_state[f"{m_id}_dist_use_custom_bins"] = False
        if f"{m_id}_dist_x_bins" not in st.session_state:
            st.session_state[f"{m_id}_dist_x_bins"] = cfg["default"]
    
    def render_scatter_config(cfg, widget_ns="main"):
        """
        widget_ns 用于区分不同渲染区域，避免重复 key：
        - current：当前图参数设置
        - advanced：全部图区间设置（高级）
        """
        prefix = cfg["prefix"]
        init_scatter_state(cfg)
    
        # 主状态 key（真正配置值）
        state_use_key = f"{prefix}_use_custom_bins"
        state_x_key = f"{prefix}_x_bins"
        state_y_key = f"{prefix}_y_bins"
    
        # 当前区域 widget key
        widget_use_key = f"{widget_ns}_{prefix}_use_custom_bins"
        widget_x_key = f"{widget_ns}_{prefix}_x_bins"
        widget_y_key = f"{widget_ns}_{prefix}_y_bins"
        widget_reset_key = f"{widget_ns}_reset_{prefix}"
    
        # 渲染前，把主状态同步到当前 widget
        st.session_state[widget_use_key] = st.session_state[state_use_key]
        st.session_state[widget_x_key] = st.session_state[state_x_key]
        st.session_state[widget_y_key] = st.session_state[state_y_key]
    
        def sync_use():
            st.session_state[state_use_key] = st.session_state[widget_use_key]
    
        def sync_x():
            st.session_state[state_x_key] = st.session_state[widget_x_key]
    
        def sync_y():
            st.session_state[state_y_key] = st.session_state[widget_y_key]
    
        def reset():
            st.session_state[state_use_key] = False
            st.session_state[state_x_key] = cfg["x_default"]
            st.session_state[state_y_key] = cfg["y_default"]
    
            st.session_state[widget_use_key] = False
            st.session_state[widget_x_key] = cfg["x_default"]
            st.session_state[widget_y_key] = cfg["y_default"]
    
        st.markdown(f"#### {cfg['name']} × {cfg['name']}增长率散点图 - 坐标轴区间设置")
        st.caption(cfg["caption"])
    
        st.checkbox(
            "启用自定义区间（勾选后生效）",
            key=widget_use_key,
            on_change=sync_use
        )
    
        c1, c2 = st.columns(2)
        with c1:
            st.text_input(
                cfg["x_help"],
                key=widget_x_key,
                help=f"示例：{cfg['x_default']}",
                on_change=sync_x
            )
        with c2:
            st.text_input(
                cfg["y_help"],
                key=widget_y_key,
                help=f"示例：{cfg['y_default']}",
                on_change=sync_y
            )
    
        c3, _ = st.columns([1, 3])
        with c3:
            st.button(
                "恢复默认区间",
                key=widget_reset_key,
                on_click=reset
            )
    
    def render_stack_dist_config(m_id, cfg, widget_ns="main"):
        """
        widget_ns 用于区分不同渲染区域，避免重复 key：
        - current：当前图参数设置
        - advanced：全部图区间设置（高级）
        """
        init_stack_dist_state(m_id, cfg)
    
        # 主状态 key
        state_use_key = f"{m_id}_dist_use_custom_bins"
        state_x_key = f"{m_id}_dist_x_bins"
    
        # 当前区域 widget key
        widget_use_key = f"{widget_ns}_{m_id}_dist_use_custom_bins"
        widget_x_key = f"{widget_ns}_{m_id}_dist_x_bins"
        widget_reset_key = f"{widget_ns}_reset_{m_id}_dist"
    
        # 渲染前同步
        st.session_state[widget_use_key] = st.session_state[state_use_key]
        st.session_state[widget_x_key] = st.session_state[state_x_key]
    
        def sync_use():
            st.session_state[state_use_key] = st.session_state[widget_use_key]
    
        def sync_x():
            st.session_state[state_x_key] = st.session_state[widget_x_key]
    
        def reset():
            st.session_state[state_use_key] = False
            st.session_state[state_x_key] = cfg["default"]
    
            st.session_state[widget_use_key] = False
            st.session_state[widget_x_key] = cfg["default"]
    
        st.markdown(f"#### {cfg['name']}堆叠分布图 - X轴区间设置")
        st.caption(cfg["caption"])
    
        st.checkbox(
            "启用自定义区间",
            key=widget_use_key,
            on_change=sync_use
        )
    
        c1, c2 = st.columns([2, 1])
        with c1:
            st.text_input(
                "X轴区间边界（%）",
                key=widget_x_key,
                help=f"示例：{cfg['default']}",
                on_change=sync_x
            )
        with c2:
            st.button(
                "恢复默认",
                key=widget_reset_key,
                on_click=reset
            )
    
    def render_all_advanced_configs():
        """
        注意：这里会被放在外层 expander 内部，因此不能再嵌套 expander
        改为 tabs 展示
        """
        st.markdown("### 散点热力图区间设置")
        scatter_items = list(scatter_config_map.items())
    
        if scatter_items:
            scatter_tabs = st.tabs([cfg["name"] for _, cfg in scatter_items])
            for tab, (m_id, cfg) in zip(scatter_tabs, scatter_items):
                with tab:
                    render_scatter_config(cfg, widget_ns="advanced")
        else:
            st.caption("暂无散点热力图配置。")
    
        st.markdown("---")
        st.markdown("### 比率堆叠分布图区间设置")
        stack_items = list(stack_dist_config_map.items())
    
        if stack_items:
            stack_tabs = st.tabs([cfg["name"] for _, cfg in stack_items])
            for tab, (m_id, cfg) in zip(stack_tabs, stack_items):
                with tab:
                    render_stack_dist_config(m_id, cfg, widget_ns="advanced")
        else:
            st.caption("暂无比率堆叠分布图配置。")
    
    # ==========================================
    # 4. 注释表加载区（只保留注释表相关）
    # ==========================================
    
    notes_dict_8 = {}
    ordered_modules = []
    df_notes = None
    first_levels = []
    
    with st.expander("📥 行业分析注释输入", expanded=False):
        use_default = st.toggle("使用默认注释表", value=False, key="step8_use_default")
    
        if use_default:
            try:
                df_notes = pd.read_excel(
                    "https://github.com/z-xylym/my-actuary-tool/raw/refs/heads/main/%E5%9B%BE%E7%89%87%E5%86%85%E5%AE%B9%E5%88%86%E6%9E%90%E5%92%8C%E6%B3%A8%E9%87%8A520.xlsx"
                )
                st.success("✅ 内置默认注释表加载成功")
            except Exception as e:
                st.error(f"❌ 加载失败：{e}")
        else:
            notes_file = st.file_uploader("上传 Excel 分析注释表", type=['xlsx', 'xls'], key="step8_notes")
            if notes_file:
                df_notes = pd.read_excel(notes_file)
    
        if df_notes is not None:
            for col in ['一级分类', '二级分类', '对应图表名称', '模块ID', '分析内容', '注释内容']:
                if col in df_notes.columns:
                    df_notes[col] = df_notes[col].astype(str).str.strip().replace(['nan', 'NaN', 'NAN', 'None'], '')
                else:
                    df_notes[col] = ''
    
            for _, r in df_notes.iterrows():
                m_id = str(r.get('模块ID', '')).strip()
                if m_id:
                    notes_dict_8[m_id] = {
                        'title': str(r.get('对应图表名称', '')).strip(),
                        'analysis': str(r.get('分析内容', '')).strip(),
                        'note': str(r.get('注释内容', '')).strip(),
                        '一级分类': str(r.get('一级分类', '')).strip(),
                        '二级分类': str(r.get('二级分类', '')).strip()
                    }
                    ordered_modules.append(m_id)
    
            first_levels = [x for x in df_notes['一级分类'].unique() if x and x != '']
            st.session_state['step8_notes_dict'] = notes_dict_8
            st.session_state['step8_ordered_modules'] = ordered_modules
            st.session_state['step8_df_notes'] = df_notes.copy()
    
    # 如果本次未在 expander 中加载到，则尝试从 session_state 取
    if df_notes is None and 'step8_df_notes' in st.session_state:
        df_notes = st.session_state['step8_df_notes'].copy()
        notes_dict_8 = st.session_state.get('step8_notes_dict', {})
        ordered_modules = st.session_state.get('step8_ordered_modules', [])
        first_levels = [x for x in df_notes['一级分类'].unique() if x and x != ''] if df_notes is not None else []
    
    # ==========================================
    # 5. 侧边栏导航
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
    # 6. 当前图参数设置
    # ==========================================
    
    if not print_mode and active_m_id:
        with st.expander(f"⚙️ 当前图参数设置：{active_chart_name or active_m_id}", expanded=True):
            st.info(f"当前正在配置：{active_chart_name or active_m_id}")
    
            if active_m_id in scatter_config_map:
                render_scatter_config(scatter_config_map[active_m_id], widget_ns="current")
            elif active_m_id in stack_dist_config_map:
                render_stack_dist_config(active_m_id, stack_dist_config_map[active_m_id], widget_ns="current")
            else:
                st.caption("当前图暂无可配置的区间参数。")
    
    # ==========================================
    # 7. 全部图区间设置（高级）
    # ==========================================
    
    if not print_mode:
        with st.expander("🗂️ 全部图区间设置（高级）", expanded=False):
            st.caption("如需批量维护所有图的默认区间或自定义区间，可在此统一设置。")
            
            # 散点热力图区间设置
            st.markdown("### 散点热力图区间设置")
            scatter_items = list(scatter_config_map.items())
            
            if scatter_items:
                scatter_options = [cfg["name"] for _, cfg in scatter_items]
                selected_scatter = st.selectbox(
                    "选择要配置的散点图",
                    options=scatter_options,
                    key="advanced_scatter_select"
                )
                
                for m_id, cfg in scatter_items:
                    if cfg["name"] == selected_scatter:
                        st.markdown("---")
                        render_scatter_config(cfg, widget_ns="advanced")
                        break
            else:
                st.caption("暂无散点热力图配置。")
            
            st.markdown("---")
            st.markdown("### 比率堆叠分布图区间设置")
            stack_items = list(stack_dist_config_map.items())
            
            if stack_items:
                stack_options = [cfg["name"] for _, cfg in stack_items]
                selected_stack = st.selectbox(
                    "选择要配置的堆叠分布图",
                    options=stack_options,
                    key="advanced_stack_select"
                )
                
                for m_id, cfg in stack_items:
                    if cfg["name"] == selected_stack:
                        st.markdown("---")
                        render_stack_dist_config(m_id, cfg, widget_ns="advanced")
                        break
            else:
                st.caption("暂无比率堆叠分布图配置。")
            
            
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
        """计算散点图数据
        
        Args:
            df: 原始数据
            selected_types: 选中的公司类型列表
            latest_year: 最新年份
            prev_year: 前一年份
            field_name: 字段名（如 '净投资回报', '保险服务收入合计', '净利润'）
            display_name: 显示名称（用于列名）
        
        Returns:
            DataFrame 包含 公司, 公司类型, display_name, 增长率, id
        """
        df_current = df[df['报告年份'].astype(str) == str(latest_year)].copy()
        df_prev = df[df['报告年份'].astype(str) == str(prev_year)].copy()
        
        scatter_data = []
        
        for co in df_current['公司'].unique():
            current_rows = df_current[(df_current['公司'] == co) & (df_current['字段名'] == field_name)]
            if current_rows.empty:
                continue
            current_val = current_rows['(百万)人民币'].iloc[0] / 100
        
            prev_rows = df_prev[(df_prev['公司'] == co) & (df_prev['字段名'] == field_name)]
            if prev_rows.empty:
                continue
            prev_val = prev_rows['(百万)人民币'].iloc[0] / 100
            
            if prev_val != 0:
                growth_rate = (current_val - prev_val) / abs(prev_val) * 100
            else:
                growth_rate = 0
            
            co_type_rows = df_current[df_current['公司'] == co]['公司类型']
            if co_type_rows.empty:
                continue
            co_type = co_type_rows.iloc[0]
            
            if co_type in selected_types:
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
        
        if not company_ratios:
            return pd.DataFrame(), []
        
        # 2. 确定分箱边界
        ratios_list = list(company_ratios.values())
        if x_bins_custom and len(x_bins_custom) > 1:
            bins = sorted(set(x_bins_custom))
        else:
            min_r, max_r = min(ratios_list), max(ratios_list)
            n_bins = min(6, max(2, len(set(ratios_list))))
            if max_r == min_r:
                bins = [min_r - 1, max_r + 1]
            else:
                step = (max_r - min_r) / n_bins
                bins = [min_r + i * step for i in range(n_bins + 1)]
        
        labels = [f"({bins[i]:.1f}%, {bins[i+1]:.1f}%]" for i in range(len(bins) - 1)]
        
        # 3. 按公司类型统计分布
        distribution = {ct: {lbl: 0 for lbl in labels} for ct in selected_types}
        
        def get_bin_idx(val, bins):
            if val <= bins[0]: return 0
            for i, (low, high) in enumerate(zip(bins[:-1], bins[1:])):
                if low < val <= high: return i
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
        """计算单个公司类型的行业平均利润构成（返回贡献百分比字典和CSM释放贡献）"""
        mask = (df['公司类型'] == company_type) & (df['报告年份'].astype(str) == str(year))
        df_filtered = df[mask].copy()
        
        if df_filtered.empty:
            return None
        
        companies = df_filtered['公司'].unique()
        company_contributions = []
        company_csm_ratios = []
        
        for co in companies:
            df_co = df_filtered[df_filtered['公司'] == co]
            
            insurance_profit_row = df_co[df_co['字段名'] == "保险利润"]
            if insurance_profit_row.empty:
                continue
            insurance_profit = insurance_profit_row['(百万)人民币'].iloc[0]
            if pd.isna(insurance_profit) or insurance_profit == 0:
                continue
            
            csm_amort = df_co[df_co['字段名'] == "合同服务边际的摊销"]['(百万)人民币'].sum()
            ra_change = df_co[df_co['字段名'] == "非金融风险调整的变动"]['(百万)人民币'].sum()
            lc_recog = df_co[df_co['字段名'] == "亏损部分的确认及转回"]['(百万)人民币'].sum()
            paa_result = df_co[df_co['字段名'] == "采用保费分配法计量的保险合同保险业绩"]['(百万)人民币'].sum()
            reinsurance = df_co[df_co['字段名'] == "再保净损益"]['(百万)人民币'].sum()
            
            operating_var = insurance_profit - csm_amort - ra_change + lc_recog - paa_result - reinsurance
            
            contribution = {
                "合同服务边际的释放": csm_amort / insurance_profit * 100,
                "非金融风险调整的变动": ra_change / insurance_profit * 100,
                "亏损部分的确认及转回": -lc_recog / insurance_profit * 100,
                "营运偏差及其他": operating_var / insurance_profit * 100,
                "保费分配法业务净损益": paa_result / insurance_profit * 100,
                "再保净损益": reinsurance / insurance_profit * 100
            }
            company_contributions.append(contribution)
            company_csm_ratios.append(csm_amort / insurance_profit * 100)
        
        if not company_contributions:
            return None
        
        contributions_df = pd.DataFrame(company_contributions)
        return {
            'contributions': {col: contributions_df[col].mean() for col in contributions_df.columns},
            'csm_ratio': np.mean(company_csm_ratios)
        }

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
        """创建散点图
        
        Args:
            df_scatter: 散点图数据（必须包含 '公司', '公司类型', id, 以及 x 和 y 对应的列）
            config: 配置字典，包含:
                - x_col: X轴列名
                - y_col: Y轴列名
                - x_bins_custom: 自定义X轴区间（可选）
                - y_bins_custom: 自定义Y轴区间（可选）
            x_label: X轴标签
            y_label: Y轴标签
        """
        import matplotlib.pyplot as plt
        import matplotlib as mpl
        import numpy as np
        from matplotlib.patches import Rectangle
        from matplotlib.colors import to_rgb, LinearSegmentedColormap
        
        mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
        mpl.rcParams['axes.unicode_minus'] = False
        mpl.rcParams['figure.dpi'] = 150          # 增加全局 DPI
        mpl.rcParams['savefig.dpi'] = 300         # 保存时的 DPI
        mpl.rcParams['font.size'] = 10  
        
        mpl.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
        mpl.rcParams['axes.unicode_minus'] = False
        
        if df_scatter.empty:
            st.warning("没有有效数据")
            return
        
        # 使用统一色卡
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
        
        # 排序
        existing_types = [ct for ct in category_order if ct in df_scatter['公司类型'].values]
        
        df_sorted = pd.DataFrame()
        for ct in existing_types:
            df_ct = df_scatter[df_scatter['公司类型'] == ct].sort_values(x_col, ascending=False)
            df_sorted = pd.concat([df_sorted, df_ct])
        
        other_types = [ct for ct in df_scatter['公司类型'].unique() if ct not in existing_types]
        for ct in other_types:
            df_ct = df_scatter[df_scatter['公司类型'] == ct].sort_values(x_col, ascending=False)
            df_sorted = pd.concat([df_sorted, df_ct])
        
        df_sorted = df_sorted.reset_index(drop=True)
        df_sorted['id'] = range(1, len(df_sorted) + 1)
        df_scatter = df_sorted.copy()
        
        # 获取数据范围
        x_values = df_scatter[x_col].values
        y_values = df_scatter[y_col].values
        min_x = min(x_values)
        max_x = max(x_values)
        min_y = min(y_values)
        max_y = max(y_values)
        
        # X轴区间设置
        if config.get('x_bins_custom') and len(config['x_bins_custom']) > 1:
            bins_x = list(config['x_bins_custom'])
            x_labels = [f"({bins_x[i]:.0f}, {bins_x[i+1]:.0f}]" for i in range(len(bins_x) - 1)]
        else:
            step = max(10, int((max_x - min_x) / 4 / 10) * 10) if max_x > 10 else max(1, int((max_x - min_x) / 4))
            bins_x = list(np.arange(min_x - 5, max_x + step, step))
            if len(bins_x) < 2:
                bins_x = [min_x - 5, max_x + 5]
            bins_x = sorted(set(bins_x))
            x_labels = [f"({bins_x[i]:.0f}, {bins_x[i+1]:.0f}]" for i in range(len(bins_x) - 1)]
        
        # Y轴区间设置
        if config.get('y_bins_custom') and len(config['y_bins_custom']) > 1:
            bins_y = list(config['y_bins_custom'])
            y_labels = [f"({bins_y[i]:.0f}%, {bins_y[i+1]:.0f}%]" for i in range(len(bins_y) - 1)]
        else:
            y_range = max_y - min_y
            step = max(10, int(y_range / 5)) if y_range != 0 else 10
            bins_y = list(np.arange(min_y - 10, max_y + step, step))
            if len(bins_y) < 2:
                bins_y = [min_y - 10, max_y + 10]
            bins_y = sorted(set(bins_y))
            y_labels = [f"({bins_y[i]:.0f}%, {bins_y[i+1]:.0f}%]" for i in range(len(bins_y) - 1)]
        
        def assign_x_bin(val):
            if val <= bins_x[0]:
                return 0
            for idx, (low, high) in enumerate(zip(bins_x[:-1], bins_x[1:])):
                if low < val <= high:
                    return idx
            return len(bins_x) - 2
        
        def assign_y_bin(val):
            if val <= bins_y[0]:
                return 0
            for idx, (low, high) in enumerate(zip(bins_y[:-1], bins_y[1:])):
                if low < val <= high:
                    return idx
            return len(bins_y) - 2
        
        df_scatter['x_bin'] = df_scatter[x_col].apply(assign_x_bin)
        df_scatter['y_bin'] = df_scatter[y_col].apply(assign_y_bin)
        
        nx, ny = len(x_labels), len(y_labels)
        
        # 热力图背景矩阵
        density = np.zeros((ny, nx))
        for _, row in df_scatter.iterrows():
            if 0 <= row['x_bin'] < nx and 0 <= row['y_bin'] < ny:
                density[int(row['y_bin']), int(row['x_bin'])] += 1
        
        max_count = max(1, int(density.max()))
        kpmg_heat_cmap = LinearSegmentedColormap.from_list(
            "kpmg_heat_red",
            ["#FFFFFF", "#FDEBEC", "#F9C7C8", "#F39A9C", KPMG_CATEGORIES["Traffic Light"]["Red"]]
        )
        
        fig = plt.figure(figsize=(16.5, 9), facecolor="white", dpi=150)
        ax_scatter = fig.add_axes([0.05, 0.11, 0.56, 0.80], facecolor="white")
        
        # 画背景格子
        for y in range(ny):
            for x in range(nx):
                count = int(density[y, x])
                face = "#FFFFFF" if count == 0 else kpmg_heat_cmap(0.10 + 0.70 * count / max_count)
                rect = Rectangle((x, y), 1, 1, facecolor=face, edgecolor="#D9E2EC", linewidth=1.0)
                ax_scatter.add_patch(rect)
                if count > 0:
                    ax_scatter.text(x + 0.93, y + 0.93, str(count), ha="right", va="top", fontsize=8, color="#4B5563", alpha=0.75)
        
        for x in range(nx + 1):
            ax_scatter.axvline(x, color="#D9E2EC", linewidth=1)
        for y in range(ny + 1):
            ax_scatter.axhline(y, color="#D9E2EC", linewidth=1)
        
        # 插值计算点位置
        def add_interpolated_positions(df):
            x_positions = pd.Series(index=df.index, dtype=float)
            y_positions = pd.Series(index=df.index, dtype=float)
            
            for idx, row in df.iterrows():
                x_bin_idx = int(row['x_bin'])
                x_min, x_max = bins_x[x_bin_idx], bins_x[x_bin_idx + 1]
                x_val = row[x_col]
                if x_max > x_min:
                    x_frac = (x_val - x_min) / (x_max - x_min)
                    x_frac = max(0.05, min(0.95, x_frac))
                else:
                    x_frac = 0.5
                x_positions.loc[idx] = x_bin_idx + x_frac
                
                y_bin_idx = int(row['y_bin'])
                y_min, y_max = bins_y[y_bin_idx], bins_y[y_bin_idx + 1]
                y_val = row[y_col]
                if y_max > y_min:
                    y_frac = (y_val - y_min) / (y_max - y_min)
                    y_frac = max(0.05, min(0.95, y_frac))
                else:
                    y_frac = 0.5
                y_positions.loc[idx] = y_bin_idx + y_frac
                
            df["x_pos"] = x_positions
            df["y_pos"] = y_positions
        
        add_interpolated_positions(df_scatter)
        
        # 画散点
        point_count = len(df_scatter)
        bubble_size = max(50, int(240 * (50 / point_count))) if point_count > 0 else 100
        
        for ct in category_order:
            if ct not in df_scatter['公司类型'].values:
                continue
            df_ct = df_scatter[df_scatter['公司类型'] == ct]
            color = category_colors.get(ct, KPMG_CATEGORIES["Primary Colors"]["Cobalt Blue"])
            ax_scatter.scatter(df_ct["x_pos"], df_ct["y_pos"], s=bubble_size, c=color, alpha=0.85,
                              edgecolors="white", linewidths=1.8, label=ct, zorder=3)
        
        # 画编号
        for _, row in df_scatter.iterrows():
            color = category_colors.get(row["公司类型"], KPMG_CATEGORIES["Primary Colors"]["Cobalt Blue"])
            rgb = to_rgb(color)
            luminance = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
            text_color = "white" if luminance < 0.55 else "#0C233C"
            ax_scatter.text(row["x_pos"], row["y_pos"], str(int(row["id"])), ha="center", va="center",
                           fontsize=10, fontweight="bold", color=text_color, zorder=4)
        
        # 坐标轴
        ax_scatter.set_xlim(0, nx)
        ax_scatter.set_ylim(0, ny)
        ax_scatter.set_xticks([i + 0.5 for i in range(nx)])
        ax_scatter.set_xticklabels(x_labels, fontsize=9.5, rotation=15, color="#475569")
        ax_scatter.set_yticks([i + 0.5 for i in range(ny)])
        ax_scatter.set_yticklabels(y_labels, fontsize=9.5, color="#475569")
        ax_scatter.tick_params(length=0, pad=8)
        ax_scatter.set_xlabel(x_label, fontsize=11, color="#334155", labelpad=12)
        ax_scatter.set_ylabel(y_label, fontsize=11, color="#334155", labelpad=12)
        
        for spine in ax_scatter.spines.values():
            spine.set_color("#CBD5E1")
            spine.set_linewidth(1.0)
        
        # 图例
        legend = ax_scatter.legend(loc="upper left", bbox_to_anchor=(1.02, 1.00), ncol=1, frameon=False,
                                   fontsize=9.5, borderaxespad=0.0, labelspacing=1.1, handletextpad=0.6)
        if legend:
            for txt in legend.get_texts():
                txt.set_color("#334155")
        
        # 右侧表格
        ax_table = fig.add_axes([0.71, 0.08, 0.25, 0.80], facecolor="white")
        ax_table.axis("off")
        ax_table.set_xlim(0, 1)
        ax_table.set_ylim(0, 1)
        top, row_h = 0.98, 0.015
        ax_table.add_patch(Rectangle((0.00, top - row_h), 0.98, row_h, 
                                     facecolor=KPMG_CATEGORIES["Primary Colors"]["KPMG Blue"], edgecolor="none"))
        ax_table.text(0.05, top - row_h / 2, "编号", va="center", ha="left", fontsize=8, color="white", fontweight="bold")
        ax_table.text(0.19, top - row_h / 2, "公司", va="center", ha="left", fontsize=8, color="white", fontweight="bold")
        ax_table.text(0.72, top - row_h / 2, "类型", va="center", ha="left", fontsize=8, color="white", fontweight="bold")
        
        max_rows = int((top - 0.02) / row_h) - 1
        df_show = df_scatter.head(max_rows).copy()
        
        for i, (_, row) in enumerate(df_show.iterrows(), start=1):
            y = top - row_h * (i + 1)
            fill = "#F8FAFC" if i % 2 == 0 else "white"
            ax_table.add_patch(Rectangle((0.00, y), 0.98, row_h, facecolor=fill, edgecolor="#E2E8F0", linewidth=0.6))
            ax_table.text(0.06, y + row_h / 2, str(int(row["id"])), ha="center", va="center", fontsize=7.5, fontweight="bold", color="#0F172A")
            company_name = str(row["公司"])
            if len(company_name) > 10:
                company_name = company_name[:10] + "…"
            ax_table.text(0.19, y + row_h / 2, company_name, ha="left", va="center", fontsize=5.5, color="#1F2937")
            type_color = category_colors.get(row["公司类型"], KPMG_CATEGORIES["Primary Colors"]["Cobalt Blue"])
            ax_table.text(0.72, y + row_h / 2, str(row["公司类型"]), ha="left", va="center", fontsize=7.2, color=type_color, fontweight="bold")
        
        st.pyplot(fig, clear_figure=True)
        plt.close(fig)
    
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
                x=labels, y=values, name=ct, marker_color=color, width=0.36,
                text=[f"{int(v)}" if show_labels and v > 0 else "" for v in values],
                textposition='inside', textfont=dict(size=label_size, color="white"),
                hovertemplate=f"{ct}: %{{y:.0f}}家<br>%{{x}}<extra></extra>"
            ))
        
        # Y轴设置
        total_per_bin = distribution_df.sum(axis=0).values
        max_total = int(max(total_per_bin)) if len(total_per_bin) > 0 else 1
        y_max = max_total + 2
        
        if y_max <= 10: tick_step = 1
        elif y_max <= 20: tick_step = 2
        elif y_max <= 50: tick_step = 5
        else: tick_step = 10
        
        tick_vals = list(range(0, y_max + 1, tick_step))
        if tick_vals[-1] != y_max:
            tick_vals.append(y_max)
        tick_text = [f"{i}家" for i in tick_vals]
        tick_angle = -15 if len(labels) > 5 else 0
        
        fig.update_layout(
            title=dict(text=f"{metric_name}分布 - {target_year}年", x=0.5, font=dict(size=16, color="#00338D")),
            barmode='stack', bargap=0.02, bargroupgap=0, height=450,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(tickfont=dict(size=11), tickangle=tick_angle, showgrid=True, gridcolor="#E8ECF1", 
                       categoryorder='array', categoryarray=labels),
            yaxis=dict(title="公司数量", range=[0, y_max], tickvals=tick_vals, ticktext=tick_text,
                       showgrid=True, gridcolor="#E8ECF1", zeroline=True, zerolinecolor="#ccc"),
            legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="right", x=1, font=dict(size=11)),
            margin=dict(t=60, b=40, l=40, r=40)
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # ========== 2. 明细表格（使用与原代码完全相同的结构） ==========
        st.markdown("---")
        st.markdown("##### 分布数据明细")
        
        # 颜色转rgba辅助函数
        def _hex_to_rgba(hex_color, alpha=0.12):
            hex_color = str(hex_color).strip()
            if hex_color.startswith("#"):
                hex_color = hex_color[1:]
            if len(hex_color) != 6:
                return f"rgba(148, 163, 184, {alpha})"
            try:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return f"rgba({r}, {g}, {b}, {alpha})"
            except Exception:
                return f"rgba(148, 163, 184, {alpha})"
        
        # 生成表格
        html_table = '''
        <table style="
            width:100%;
            border-collapse: collapse;
            font-family: sans-serif;
            font-size: 13px;
            table-layout: fixed;
            margin-top: 10px;
        ">
        '''
        
        html_table += '<tr style="background-color: #00338D; color: white;">'
        html_table += '<th style="padding: 10px 12px; text-align: left; border: 1px solid #ddd;">公司类型</th>'
        for label in labels:
            html_table += f'<th style="padding: 10px 12px; text-align: center; border: 1px solid #ddd;">{label}</th>'
        html_table += '<th style="padding: 10px 12px; text-align: center; border: 1px solid #ddd;">合计</th>'
        html_table += '</tr>'
        
        for ct in ct_list:
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            values = distribution_df.loc[ct].values
            total = int(sum(values))
            
            html_table += '<tr style="background-color: white;">'
            html_table += (
                f'<td style="padding: 10px 12px; border: 1px solid #ddd; '
                f'font-weight: bold; background-color: {color}; color: white;">{ct}</td>'
            )
            
            for v in values:
                html_table += (
                    f'<td style="padding: 10px 12px; text-align: center; '
                    f'border: 1px solid #ddd; background-color: white; '
                    f'color: #111827; font-weight: 500;">{int(v)}家</td>'
                )
            
            html_table += (
                f'<td style="padding: 10px 12px; text-align: center; '
                f'border: 1px solid #ddd; background-color: white; '
                f'color: #111827; font-weight: bold;">{total}家</td>'
            )
            html_table += '</tr>'
        
        html_table += '</table>'
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
    
        fig = make_subplots(rows=1, cols=len(all_cts), shared_yaxes=True,
                            horizontal_spacing=0.03, column_titles=[f"<b>{ct}</b>" for ct in all_cts])
    
        df_avg = pd.DataFrame(index=[f"{year}年" for year in year_list], columns=field_map.values()).fillna(0)
    
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            x_labels = [f"{year}年" for year in year_list]
            cumulative = [0.0] * len(year_list)
    
            for f in fields:
                display_name = field_map[f]
                y_vals = [results_by_year.get(year, {}).get(ct, {}).get(display_name, 0) for year in year_list]
    
                fig.add_trace(go.Bar(x=x_labels, y=y_vals, name=display_name if i == 0 else None,
                                     marker_color=color_map[f], width=config['bar_width'],
                                     showlegend=(i == 0), legendgroup=f,
                                     hovertemplate="%{x}<br>%{fullData.name}: %{y:.1f}%<extra></extra>"), row=1, col=col_idx)
    
                if config.get('show_labels', True):
                    txt_c = "white" if color_map[f] in dark_colors else "black"
                    for j, v in enumerate(y_vals):
                        if abs(v) >= 1:
                            y_mid = cumulative[j] + v / 2
                            fig.add_annotation(x=x_labels[j], y=y_mid, text=f"{v:.1f}%", showarrow=False,
                                               xref=f"x{col_idx}" if col_idx > 1 else "x",
                                               yref=f"y{col_idx}" if col_idx > 1 else "y",
                                               font=dict(size=config['label_size'], color=txt_c),
                                               xanchor="center", yanchor="middle")
                cumulative = [cumulative[j] + y_vals[j] for j in range(len(year_list))]
    
            # Step 7 风格灰色框
            fig.add_shape(type="rect", xref="x domain" if col_idx == 1 else f"x{col_idx} domain",
                          yref="y domain" if col_idx == 1 else f"y{col_idx} domain",
                          x0=-0.06, x1=1.06, y0=-0.1, y1=1.15,
                          fillcolor="rgba(0,0,0,0)", line=dict(color="#E0E0E0", width=1), layer="below", row=1, col=col_idx)
    
            for year_idx, year in enumerate(year_list):
                year_str = f"{year}年"
                for f in fields:
                    df_avg.loc[year_str, field_map[f]] += results_by_year.get(year, {}).get(ct, {}).get(field_map[f], 0)
    
        df_avg = df_avg / len(all_cts)
    
        fig.update_layout(barmode='stack', bargap=0.2, height=500,
                          margin=dict(t=50, b=80, l=20, r=20),
                          legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(size=10)),
                          plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", ticklen=0, row=1, col=i)
            fig.update_yaxes(showgrid=False, range=[0, 100], tickvals=[0, 25, 50, 75, 100],
                             ticktext=["0%", "25%", "50%", "75%", "100%"], zeroline=True, zerolinecolor="#E0E0E0", row=1, col=i)
    
        for ann in fig.layout.annotations:
            if "<b>" in str(ann.text):
                ann.update(y=1.08, font=dict(size=config['co_font_size'], color="#00338D"))
    
        return fig, df_avg
        
    def create_profit_composition_chart(results_by_ct, config):
        """绘制行业平均利润构成图（也需要加灰色框）"""
        display_mapping = [
            ("亏损部分的确认及转回", "亏损部分的确认及转回", "rgb(190, 190, 190)"),
            ("合同服务边际的释放", "合同服务边际的释放", "rgb(30, 73, 226)"),
            ("非金融风险调整的变动", "非金融风险调整的变动", "rgb(118, 210, 255)"),
            ("营运偏差及其他", "营运偏差及其他", "rgb(114, 19, 214)"),
            ("保费分配法业务净损益", "保费分配法业务净损益", "rgb(253, 52, 156)"),
            ("再保净损益", "再保净损益", "rgb(9, 142, 126)")
        ]
        
        fig = go.Figure()
        ct_list = list(results_by_ct.keys())
        x_indices = list(range(len(ct_list)))
        
        for col_name, legend_name, color in display_mapping:
            y_vals = [results_by_ct[ct]['contributions'].get(col_name, 0) for ct in ct_list]
            txt_color = "white" if color in ["rgb(30, 73, 226)", "rgb(114, 19, 214)", "rgb(9, 142, 126)"] else "black"
            fig.add_trace(go.Bar(
                name=legend_name, x=x_indices, y=y_vals, width=config['bar_width'],
                marker_color=color,
                text=[f"{v:.0f}%" if config['show_labels'] and abs(v) >= 2 else "" for v in y_vals],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=config['label_size'], color=txt_color)
            ))
        
        # 🌟 为每个公司添加灰色背景框（这个图没有 subplot，是单图多柱子）
        for i, ct in enumerate(ct_list):
            fig.add_shape(
                type="rect", xref="x", yref="paper",
                x0=i - 0.46, x1=i + 0.46, y0=0, y1=1,
                fillcolor="rgba(200, 200, 200, 0.15)",
                line=dict(color="rgba(150, 150, 150, 0.6)", width=1.5),
                layer="below"
            )
        
        fig.update_layout(
            barmode='relative', height=550,
            bargap=0.25,  # 🌟 紧凑间距
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, b=80, l=220, r=80),
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=-0.05,
                       traceorder="reversed", font=dict(size=11, color="#00338D")),
            yaxis=dict(side='right', showgrid=False, range=[-45, 105],
                      tickmode='array', showticklabels=True,
                      tickvals=[-40, -20, 0, 20, 40, 60, 80, 100],
                      ticktext=["-40%", "-20%", "0%", "20%", "40%", "60%", "80%", "100%"],
                      zeroline=True, zerolinecolor="#F7860C")
        )
        
        x_labels = [f"<span style='font-size:{config['co_font_size']}px;color:#00338D;'><b>{ct}</b></span>" for ct in ct_list]
        fig.update_xaxes(showgrid=False, zeroline=False, tickvals=x_indices, ticktext=x_labels, side="top")
        
        return fig
    def create_expense_chart(results_by_year, year_list, selected_types, config):
        """绘制行业平均费用结构图 - 参照 Step 7 的 create_expense_breakdown_chart 样式"""
        from plotly.subplots import make_subplots
        
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
        
        # 创建子图：一行多列
        fig = make_subplots(
            rows=1, cols=len(all_cts),
            shared_yaxes=True,
            subplot_titles=[f"<b>{ct}</b>" for ct in all_cts],
            horizontal_spacing=0.02
        )
        
        # 添加图例（只在第一个子图显示）
        for leg_name, leg_color in [("获取费用", "rgb(30, 73, 226)"), 
                                      ("维持费用", "rgb(118, 210, 255)"), 
                                      ("非履约费用", "rgb(114, 19, 234)")]:
            fig.add_trace(
                go.Scatter(x=[None], y=[None], mode="markers", 
                          marker=dict(symbol="square", size=12, color=leg_color), 
                          name=leg_name, showlegend=True),
                row=1, col=1
            )
        
        for i, ct in enumerate(all_cts):
            col_idx = i + 1
            
            # 获取两年数据
            years_list = []
            for year in year_list:
                ratios = results_by_year.get(year, {}).get(ct, {}).get('ratios', {})
                years_list.append([ratios.get(f, 0) for f in expense_fields])
            
            x_vals = list(range(len(year_list)))
            
            # 绘制三个费用类别（堆叠）
            for j, field in enumerate(expense_fields):
                y_vals = [data[j] for data in years_list]
                is_dark = field in ["获取费用", "非履约费用"]
                
                fig.add_trace(
                    go.Bar(
                        x=x_vals, y=y_vals, name=field,
                        marker_color=color_map[field],
                        text=[f"{v:.0f}%" if config['show_labels'] and v > 0 else "" for v in y_vals],
                        textposition='inside', insidetextanchor='middle',
                        textfont=dict(size=config['label_size'], color="white" if is_dark else "#1a1a2e"),
                        width=config['bar_width'], showlegend=(i == 0), legendgroup=field
                    ),
                    row=1, col=col_idx
                )
            
            # 添加顶部总费用标注（单位：亿元）
            for year_idx, year in enumerate(year_list):
                total_val = results_by_year.get(year, {}).get(ct, {}).get('avg_total', 0) / 100
                fig.add_annotation(
                    x=year_idx, y=105,
                    text=f"<b>{total_val:.0f}</b>",
                    showarrow=False, font=dict(size=config['label_size'], color="#222"),
                    bgcolor="white", bordercolor="#BBB", borderwidth=1,
                    xanchor="center", yanchor="bottom",
                    row=1, col=col_idx
                )
            
            # 🌟 关键：添加灰色背景框（参照 Step 7）
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
            
            # 设置X轴标签
            fig.update_xaxes(
                tickvals=[0, 1],
                ticktext=[f"{year}年" for year in year_list],
                row=1, col=col_idx
            )
        
        # 全局布局
        fig.update_layout(
            barmode='stack',
            bargap=0.25,           # 不同年份柱子间距
            height=550,
            margin=dict(t=50, b=80, l=10, r=10),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, font=dict(size=12))
        )
        
        # Y轴设置：百分比，上限110%留出顶部标注空间
        fig.update_yaxes(range=[0, 110], showgrid=False, zeroline=False, tickformat=".0f", ticksuffix="%")
        
        # 隐藏X轴网格线
        for i in range(1, len(all_cts) + 1):
            fig.update_xaxes(showgrid=False, showline=False, zeroline=False, ticks="", row=1, col=i)
        
        # 子图标题样式
        for ann in fig.layout.annotations:
            ann.update(y=1.08, font=dict(size=14, color="#00338D"))
        
        return fig
    
    def create_profit_composition_chart(results_by_ct, config):
        """绘制行业平均利润构成图"""
        display_mapping = [
            ("亏损部分的确认及转回", "亏损部分的确认及转回", "rgb(190, 190, 190)"),
            ("合同服务边际的释放", "合同服务边际的释放", "rgb(30, 73, 226)"),
            ("非金融风险调整的变动", "非金融风险调整的变动", "rgb(118, 210, 255)"),
            ("营运偏差及其他", "营运偏差及其他", "rgb(114, 19, 214)"),
            ("保费分配法业务净损益", "保费分配法业务净损益", "rgb(253, 52, 156)"),
            ("再保净损益", "再保净损益", "rgb(9, 142, 126)")
        ]
        
        fig = go.Figure()
        ct_list = list(results_by_ct.keys())
        x_indices = list(range(len(ct_list)))
        
        for col_name, legend_name, color in display_mapping:
            y_vals = [results_by_ct[ct]['contributions'].get(col_name, 0) for ct in ct_list]
            txt_color = "white" if color in ["rgb(30, 73, 226)", "rgb(114, 19, 214)", "rgb(9, 142, 126)"] else "black"
            
            fig.add_trace(go.Bar(
                name=legend_name, x=x_indices, y=y_vals, width=config['bar_width'],
                marker_color=color,
                text=[f"{v:.0f}%" if config['show_labels'] and abs(v) >= 2 else "" for v in y_vals],
                textposition='inside', insidetextanchor='middle',
                textfont=dict(size=config['label_size'], color=txt_color)
            ))
        
        for i, ct in enumerate(ct_list):
            fig.add_shape(
                type="rect", xref="x", yref="paper",
                x0=i - 0.46, x1=i + 0.46, y0=0, y1=1,
                fillcolor="rgba(0, 51, 141, 0.05)",
                line=dict(color="rgba(0, 51, 141, 0.9)", width=2.5),
                layer="above"
            )
        
        fig.update_layout(
            barmode='relative', height=550,
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            margin=dict(t=20, b=80, l=220, r=80),
            legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="right", x=-0.05,
                       traceorder="reversed", font=dict(size=11, color="#00338D")),
            yaxis=dict(side='right', showgrid=False, range=[-45, 105],
                      tickmode='array', showticklabels=True,
                      tickvals=[-40, -20, 0, 20, 40, 60, 80, 100],
                      ticktext=["-40%", "-20%", "0%", "20%", "40%", "60%", "80%", "100%"],
                      zeroline=True, zerolinecolor="#F7860C")
        )
        
        x_labels = [f"<span style='font-size:{config['co_font_size']}px;color:#00338D;'><b>{ct}</b></span>" for ct in ct_list]
        fig.update_xaxes(showgrid=False, zeroline=False, tickvals=x_indices, ticktext=x_labels, side="top")
        
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
                    showlegend=False
                ),
                row=2, col=col_idx
            )
        
        fig.update_layout(
            barmode='group',
            bargap=config['gap'],
            bargroupgap=0.0,
            height=500,
            margin=dict(t=50, b=80, l=40, r=20),
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

    def parse_custom_bins_input(text):
        """解析自定义分箱输入，如: 0,5,10,15,20"""
        if text is None or str(text).strip() == "":
            return None
        
        try:
            bins = [float(x.strip()) for x in str(text).split(",") if x.strip() != ""]
            bins = sorted(set(bins))
            if len(bins) < 2:
                return None
            return bins
        except:
            return None

    def create_csm_ratio_distribution_chart(distribution_df, labels, config):
        """绘制 CSM 摊销比率分布的堆叠柱状图 - 使用统一色卡"""
        
        import plotly.graph_objects as go
        
        ct_list = [ct for ct in distribution_df.index.tolist() if distribution_df.loc[ct].sum() > 0]
        ct_list = sort_company_types(ct_list)
        
        if not ct_list:
            st.warning("没有有效数据")
            return go.Figure()
        
        fig = go.Figure()
        
        for ct in ct_list:
            values = distribution_df.loc[ct].values
            color = COMPANY_TYPE_COLORS.get(ct, "#94A3B8")
            
            fig.add_trace(go.Bar(
                x=labels,
                y=values,
                name=ct,
                marker_color=color,
                width=0.36,
                text=[f"{int(v)}" if config['show_labels'] and v > 0 else "" for v in values],
                textposition='inside',
                textfont=dict(size=config['label_size'], color="white"),
                hovertemplate=f"{ct}: %{{y:.0f}}家<br>%{{x}}<extra></extra>"
            ))
        
        total_per_bin = distribution_df.sum(axis=0).values
        max_total = int(max(total_per_bin)) if len(total_per_bin) > 0 else 1
        y_max = max_total + 2
        
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
        title=dict(text=f"CSM摊销比率分布 - {config['target_year']}年", x=0.5, xanchor='center', font=dict(size=16, color="#00338D")),
        barmode='stack', bargap=0.02, bargroupgap=0, height=450,
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(title="", tickfont=dict(size=11), tickangle=tick_angle, showgrid=True, gridcolor="#E8ECF1", categoryorder='array', categoryarray=labels),
        yaxis=dict(title="公司数量", range=[0, y_max], tickvals=tick_vals, ticktext=tick_text, showgrid=True, gridcolor="#E8ECF1", zeroline=True, zerolinecolor="#ccc", zerolinewidth=1),
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="right", x=1, font=dict(size=11)),
        margin=dict(t=60, b=40, l=40, r=40)
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
        # 图表：保险服务收入 vs 保险服务收入增长率散点图
        if m_id == "industry_inc_total":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "保险服务收入合计", "保险服务收入")
            
            if df_scatter.empty:
                st.warning("无法计算保险服务收入散点图数据，请检查数据是否包含'保险服务收入合计'字段")
                return

            use_custom = st.session_state.get("revenue_scatter_use_custom_bins", False)
            x_bins_input = st.session_state.get("revenue_scatter_x_bins", "0,100,200,500,1000,2000")
            y_bins_input = st.session_state.get("revenue_scatter_y_bins", "-100,0,100,200,500,1000")
            
            x_bins_custom = parse_bins_input(x_bins_input) if use_custom else None
            y_bins_custom = parse_bins_input(y_bins_input) if use_custom else None
            
            config = {
                'x_col': '保险服务收入',
                'y_col': '保险服务收入增长率',
                'x_bins_custom': x_bins_custom,
                'y_bins_custom': y_bins_custom
            }
            
            create_scatter_chart(df_scatter, config, "保险服务收入区间（亿元）", "保险服务收入增长率区间（%）")
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
        # 净投资回报 vs 净投资回报增长率散点图
        elif m_id == "industry_inv_return":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "净投资回报", "净投资回报")
            
            if df_scatter.empty:
                st.warning("无法计算净投资回报散点图数据，请检查数据是否包含'净投资回报'字段")
                return

            use_custom = st.session_state.get("inv_return_scatter_use_custom_bins", False)
            x_bins_input = st.session_state.get("inv_return_scatter_x_bins", "0,50,100,200,500,1000")
            y_bins_input = st.session_state.get("inv_return_scatter_y_bins", "-100,0,100,200,500,1000")
            
            x_bins_custom = parse_bins_input(x_bins_input) if use_custom else None
            y_bins_custom = parse_bins_input(y_bins_input) if use_custom else None
            
            config = {
                'x_col': '净投资回报',
                'y_col': '净投资回报增长率',
                'x_bins_custom': x_bins_custom,
                'y_bins_custom': y_bins_custom
            }
            
            create_scatter_chart(df_scatter, config, "净投资回报区间（亿元）", "净投资回报增长率区间（%）")
        # 保险服务费用合计热力图
        elif m_id == "industry_exp_total":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "保险服务费用合计", "保险服务费用")
            
            if df_scatter.empty:
                st.warning("无法计算保险服务费用合计数据，请检查数据是否包含'保险服务费用合计'字段")
                return
        
            use_custom = st.session_state.get("expense_total_scatter_use_custom_bins", False)
            x_bins_input = st.session_state.get("expense_total_scatter_x_bins", "0,50,100,200,500,1000")
            y_bins_input = st.session_state.get("expense_total_scatter_y_bins", "-100,0,100,200,500,1000")
            
            x_bins_custom = parse_bins_input(x_bins_input) if use_custom else None
            y_bins_custom = parse_bins_input(y_bins_input) if use_custom else None
            
            config = {
                'x_col': '保险服务费用',
                'y_col': '保险服务费用增长率',
                'x_bins_custom': x_bins_custom,
                'y_bins_custom': y_bins_custom
            }
            
            create_scatter_chart(df_scatter, config, "保险服务费用区间（亿元）", "保险服务费用增长率区间（%）")
            
        # 图表：承保财务净损益 vs 承保财务净损益增长率热力图
        elif m_id == "industry_uw_profit":
            # 🌟 核心：先计算承保财务净损益 = 承保财务损益 - 分出再保险财务损益
            df_calc = df_raw.copy()
            
            # 提取最新年和前一年的数据
            df_current = df_calc[df_calc['报告年份'].astype(str) == str(latest_year)].copy()
            df_prev = df_calc[df_calc['报告年份'].astype(str) == str(prev_year)].copy()
            
            scatter_data = []
            
            for co in selected_types:  # 注意：这里遍历的是公司类型，不是公司名
                # 需要先获取该类型下的所有公司
                companies_of_type = df_calc[df_calc['公司类型'] == co]['公司'].unique()
                
                for company in companies_of_type:
                    # 计算最新年：承保财务损益 - 分出再保险财务损益
                    current_rows = df_current[(df_current['公司'] == company)]
                    inv_uw = current_rows[current_rows['字段名'] == '承保财务损益']['(百万)人民币'].sum()
                    inv_re = current_rows[current_rows['字段名'] == '分出再保险财务损益']['(百万)人民币'].sum()
                    current_val = (inv_uw - inv_re) / 100  # 转换为亿元
                    
                    # 计算前一年
                    prev_rows = df_prev[(df_prev['公司'] == company)]
                    inv_uw_prev = prev_rows[prev_rows['字段名'] == '承保财务损益']['(百万)人民币'].sum()
                    inv_re_prev = prev_rows[prev_rows['字段名'] == '分出再保险财务损益']['(百万)人民币'].sum()
                    prev_val = (inv_uw_prev - inv_re_prev) / 100
                    
                    if prev_val != 0:
                        growth_rate = (current_val - prev_val) / abs(prev_val) * 100
                    else:
                        growth_rate = 0
                    
                    scatter_data.append({
                        '公司': company,
                        '公司类型': co,
                        '承保财务净损益': current_val,
                        '承保财务净损益增长率': growth_rate
                    })
            
            df_scatter = pd.DataFrame(scatter_data)
            
            if df_scatter.empty:
                st.warning("无法计算承保财务净损益散点图数据，请检查数据是否包含'承保财务损益'和'分出再保险财务损益'字段")
                return
            
            # 获取用户自定义区间配置
            use_custom = st.session_state.get("uw_profit_scatter_use_custom_bins", False)
            x_bins_input = st.session_state.get("uw_profit_scatter_x_bins", "-1000,-500,-200,-100,0,100,200,500")
            y_bins_input = st.session_state.get("uw_profit_scatter_y_bins", "-1000,-500,-200,-100,0,100,200,500")
            
            x_bins_custom = parse_bins_input(x_bins_input) if use_custom else None
            y_bins_custom = parse_bins_input(y_bins_input) if use_custom else None
            
            config = {
                'x_col': '承保财务净损益',
                'y_col': '承保财务净损益增长率',
                'x_bins_custom': x_bins_custom,
                'y_bins_custom': y_bins_custom
            }
            
            create_scatter_chart(df_scatter, config, "承保财务净损益区间（亿元）", "承保财务净损益增长率区间（%）")
        # 图表3：行业平均费用结构
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
        # 图表：保险服务业绩 vs 保险服务业绩增长率散点图
        elif m_id == "industry_perf_total":
            df_curr, df_prev = df_raw[df_raw['报告年份'].astype(str)==str(latest_year)], df_raw[df_raw['报告年份'].astype(str)==str(prev_year)]
            scatter_data = []
            for ct in selected_types:
                for co in df_raw[df_raw['公司类型']==ct]['公司'].unique():
                    def get_val(df, co, field): return df[(df['公司']==co) & (df['字段名']==field)]['(百万)人民币'].sum()
                    curr = (get_val(df_curr, co, '保险服务收入合计') - get_val(df_curr, co, '保险服务费用合计')) / 100
                    prev = (get_val(df_prev, co, '保险服务收入合计') - get_val(df_prev, co, '保险服务费用合计')) / 100
                    growth = (curr - prev) / abs(prev) * 100 if prev != 0 else 0
                    scatter_data.append({'公司': co, '公司类型': ct, '保险服务业绩': curr, '保险服务业绩增长率': growth})
            df_scatter = pd.DataFrame(scatter_data)
            if df_scatter.empty: st.warning("缺少保险服务收入合计或保险服务费用合计字段"); return
            use_custom = st.session_state.get("perf_total_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("perf_total_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("perf_total_scatter_y_bins", "-100,0,100,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'保险服务业绩','y_col':'保险服务业绩增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "保险服务业绩区间（亿元）", "保险服务业绩增长率区间（%）")
        # 图表4：行业平均最新年保险服务业绩构成
        elif m_id == "industry_prof_2025":
            target_year = latest_year
            
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_profit_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均利润构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示利润标签", value=False, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签字号", 8, 16, 11, key=f"psz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.4, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司字号", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = False, 11, 0.4, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig = create_profit_composition_chart(results_by_ct, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)

        # 图表5：行业平均上一年保险服务业绩构成
        elif m_id == "industry_prof_2024":
            target_year = prev_year
            
            results_by_ct = {}
            for ct in selected_types:
                result = calc_industry_profit_composition(df_raw, ct, target_year)
                if result:
                    results_by_ct[ct] = result
            
            if not results_by_ct:
                st.warning(f"无法计算 {target_year} 年行业平均利润构成")
                return
            
            if not print_mode:
                c1, c2, c3, c4 = st.columns(4)
                with c1: show_labels = st.toggle("显示利润标签", value=False, key=f"lab_{m_id}")
                with c2: label_size = st.slider("标签字号", 8, 16, 11, key=f"psz_{m_id}")
                with c3: bar_width = st.slider("柱宽", 0.2, 0.8, 0.4, key=f"wid_{m_id}")
                with c4: co_font_size = st.slider("公司字号", 10, 20, 14, key=f"cfs_{m_id}")
            else:
                show_labels, label_size, bar_width, co_font_size = False, 11, 0.4, 14
            
            config = {'show_labels': show_labels, 'label_size': label_size, 'bar_width': bar_width, 'co_font_size': co_font_size}
            
            fig = create_profit_composition_chart(results_by_ct, config)
            
            st.markdown("<p style='text-align:right; font-size:12px; color:#666;'>单位：百分比 (%)</p>", unsafe_allow_html=True)
            show_chart(fig, print_mode)
            
        # 图表：投资利润 vs 投资利润增长率散点图
        elif m_id == "industry_inv_profit":
            df_curr, df_prev = df_raw[df_raw['报告年份'].astype(str)==str(latest_year)], df_raw[df_raw['报告年份'].astype(str)==str(prev_year)]
            scatter_data = []
            for ct in selected_types:
                for co in df_raw[df_raw['公司类型']==ct]['公司'].unique():
                    def get_val(df, co, field): return df[(df['公司']==co) & (df['字段名']==field)]['(百万)人民币'].sum()
                    # 先计算承保财务净损益 = 承保财务损益 - 分出再保险财务损益
                    uw_curr = get_val(df_curr, co, '承保财务损益') - get_val(df_curr, co, '分出再保险财务损益')
                    uw_prev = get_val(df_prev, co, '承保财务损益') - get_val(df_prev, co, '分出再保险财务损益')
                    # 投资利润 = 净投资回报 - 承保财务净损益
                    curr = (get_val(df_curr, co, '净投资回报') - uw_curr) / 100
                    prev = (get_val(df_prev, co, '净投资回报') - uw_prev) / 100
                    growth = (curr - prev) / abs(prev) * 100 if prev != 0 else 0
                    scatter_data.append({'公司': co, '公司类型': ct, '投资利润': curr, '投资利润增长率': growth})
            df_scatter = pd.DataFrame(scatter_data)
            if df_scatter.empty: st.warning("缺少净投资回报、承保财务损益或分出再保险财务损益字段"); return
            use_custom = st.session_state.get("inv_profit_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("inv_profit_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("inv_profit_scatter_y_bins", "-100,0,100,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'投资利润','y_col':'投资利润增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "投资利润区间（亿元）", "投资利润增长率区间（%）")
        
        # 图表：税前利润 vs 税前利润增长率热力图
        elif m_id == "industry_tax_profit":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "税前利润总额", "税前利润")
            if df_scatter.empty: st.warning("缺少税前利润总额字段"); return
            use_custom = st.session_state.get("tax_profit_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("tax_profit_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("tax_profit_scatter_y_bins", "-100,0,100,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'税前利润','y_col':'税前利润增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "税前利润区间（亿元）", "税前利润增长率区间（%）")
               
        # 图表：净利润 vs 净利润增长率热力图
        elif m_id == "industry_net_profit":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "净利润", "净利润")
            if df_scatter.empty: st.warning("缺少净利润字段"); return
            use_custom = st.session_state.get("net_profit_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("net_profit_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("net_profit_scatter_y_bins", "-100,0,100,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'净利润','y_col':'净利润增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "净利润区间（亿元）", "净利润增长率区间（%）")
       
        # 图表：其它综合收益 vs 其它综合收益增长率热力图
        elif m_id == "industry_oci_profit":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "其他综合收益", "其他综合收益")
            if df_scatter.empty: st.warning("缺少其他综合收益字段"); return
            use_custom = st.session_state.get("oci_profit_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("oci_profit_scatter_x_bins", "-500,-200,-100,0,100,200,500")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("oci_profit_scatter_y_bins", "-1000,-500,-200,0,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'其他综合收益','y_col':'其他综合收益增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "其他综合收益区间（亿元）", "其他综合收益增长率区间（%）")
       
        # 图表：综合收益 vs 综合收益增长率热力图    
        elif m_id == "industry_total_profit":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "综合收益总额", "综合收益总额")
            if df_scatter.empty: st.warning("缺少综合收益总额字段"); return
            use_custom = st.session_state.get("total_profit_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("total_profit_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("total_profit_scatter_y_bins", "-100,0,100,200,500,1000")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'综合收益总额','y_col':'综合收益总额增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "综合收益总额区间（亿元）", "综合收益总额增长率区间（%）")
       
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
            
        # CSM余额散点图
        elif m_id == "industry_csm_bal":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "CSM期末余额", "CSM余额")
            if df_scatter.empty: st.warning("缺少CSM期末余额字段"); return
            use_custom = st.session_state.get("csm_bal_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("csm_bal_scatter_x_bins", "0,100,500,1000,2000,5000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("csm_bal_scatter_y_bins", "-50,-20,0,20,50,100,200")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'CSM余额','y_col':'CSM余额增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "CSM余额区间（亿元）", "CSM余额增长率区间（%）")
            
        # CSM/BEL占比堆叠分布图
        elif m_id == "industry_csm_ratio":
            if not print_mode:
                c1,c2=st.columns(2)
                with c1: show_labels=st.toggle("显示数值标签",value=True,key=f"lab_{m_id}")
                with c2: label_size=st.slider("标签大小",8,16,11,key=f"sz_{m_id}")
            else: show_labels,label_size=True,11
            use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
            x_bins = parse_custom_bins_input(st.session_state.get(f"{m_id}_dist_x_bins","0,20,40,60,80,100,120")) if use_custom else None
            dist_df,labels=calc_stack_distribution(df_raw,selected_types,latest_year,calc_csm_bel_ratio,x_bins_custom=x_bins)
            if dist_df.empty: st.warning("无法计算CSM/BEL占比分布"); return
            create_stack_chart_and_table(dist_df,labels,"CSM/BEL占比",latest_year,show_labels,label_size)
        
        # 2. CSM持续率堆叠分布图
        elif m_id == "industry_csm_continuity_ratio":
            if not print_mode:
                c1,c2=st.columns(2)
                with c1: show_labels=st.toggle("显示数值标签",value=True,key=f"lab_{m_id}")
                with c2: label_size=st.slider("标签大小",8,16,11,key=f"sz_{m_id}")
            else: show_labels,label_size=True,11
            use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
            x_bins = parse_custom_bins_input(st.session_state.get(f"{m_id}_dist_x_bins","0,50,100,150,200,300")) if use_custom else None
            dist_df,labels=calc_stack_distribution(df_raw,selected_types,latest_year,calc_csm_continuity_ratio,x_bins_custom=x_bins)
            if dist_df.empty: st.warning("无法计算CSM持续率分布"); return
            create_stack_chart_and_table(dist_df,labels,"CSM持续率",latest_year,show_labels,label_size)
                
        # 3. CSM摊销比率堆叠分布图
        elif m_id == "industry_csm_amort_ratio":
            if not print_mode:
                c1,c2=st.columns(2)
                with c1: show_labels=st.toggle("显示数值标签",value=True,key=f"lab_{m_id}")
                with c2: label_size=st.slider("标签大小",8,16,11,key=f"sz_{m_id}")
            else: show_labels,label_size=True,11
            use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
            x_bins = parse_custom_bins_input(st.session_state.get(f"{m_id}_dist_x_bins","0,5,10,15,20,30")) if use_custom else None
            dist_df,labels=calc_stack_distribution(df_raw,selected_types,latest_year,calc_csm_amort_ratio,x_bins_custom=x_bins)
            if dist_df.empty: st.warning("无法计算CSM摊销比率分布"); return
            create_stack_chart_and_table(dist_df,labels,"CSM摊销比率",latest_year,show_labels,label_size)
        
        # 4. 新业务IFRS利润率堆叠分布图
        elif m_id == "industry_nb_ifrs_margin":
            if not print_mode:
                c1,c2=st.columns(2)
                with c1: show_labels=st.toggle("显示数值标签",value=True,key=f"lab_{m_id}")
                with c2: label_size=st.slider("标签大小",8,16,11,key=f"sz_{m_id}")
            else: show_labels,label_size=True,11
            use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
            x_bins = parse_custom_bins_input(st.session_state.get(f"{m_id}_dist_x_bins","0,10,20,30,40,50")) if use_custom else None
            dist_df,labels=calc_stack_distribution(df_raw,selected_types,latest_year,calc_nb_ifrs_margin,x_bins_custom=x_bins)
            if dist_df.empty: st.warning("无法计算新业务IFRS利润率分布"); return
            create_stack_chart_and_table(dist_df,labels,"新业务IFRS利润率",latest_year,show_labels,label_size)

        # 净资产散点图
        elif m_id == "industry_equity_trend":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "期末股东权益", "净资产")
            if df_scatter.empty: st.warning("缺少期末股东权益字段"); return
            use_custom = st.session_state.get("equity_trend_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("equity_trend_scatter_x_bins", "0,100,500,1000,2000,5000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("equity_trend_scatter_y_bins", "-50,-20,0,20,50,100,200")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'净资产','y_col':'净资产增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "净资产区间（亿元）", "净资产增长率区间（%）")
        
        # CSM/净资产占比散点图（需要手写计算）
        elif m_id == "industry_csm_equity":
            df_curr = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)]
            df_prev = df_raw[df_raw['报告年份'].astype(str) == str(prev_year)]
            
            scatter_data = []
            for ct in selected_types:
                for co in df_raw[df_raw['公司类型']==ct]['公司'].unique():
                    def get_val(df, co, field): return df[(df['公司']==co) & (df['字段名']==field)]['(百万)人民币'].sum()
                    
                    # 最新年：税后CSM = CSM × 0.75，再计算占比
                    csm = get_val(df_curr, co, 'CSM期末余额')
                    equity = get_val(df_curr, co, '期末股东权益')
                    csm_after_tax = csm * 0.75
                    curr_ratio = csm_after_tax / equity * 100 if equity != 0 else 0
                    
                    # 前一年
                    csm_prev = get_val(df_prev, co, 'CSM期末余额')
                    equity_prev = get_val(df_prev, co, '期末股东权益')
                    csm_after_tax_prev = csm_prev * 0.75
                    prev_ratio = csm_after_tax_prev / equity_prev * 100 if equity_prev != 0 else 0
                    
                    change = curr_ratio - prev_ratio  # 百分点变化
                    
                    scatter_data.append({'公司': co, '公司类型': ct, 'CSM/净资产占比': curr_ratio, '占比变化': change})
            
            df_scatter = pd.DataFrame(scatter_data)
            if df_scatter.empty: st.warning("缺少CSM期末余额或期末股东权益字段"); return
            
            use_custom = st.session_state.get("csm_equity_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("csm_equity_scatter_x_bins", "0,20,40,60,80,100,120")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("csm_equity_scatter_y_bins", "-20,-10,-5,0,5,10,20")) if use_custom else None
            
            create_scatter_chart(df_scatter, {'x_col':'CSM/净资产占比','y_col':'占比变化','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "CSM/净资产占比（%）", "占比变化（百分点）")
        
        # 新业务CSM散点图
        elif m_id == "industry_nb_csm":
            df_scatter = calc_scatter_data(df_raw, selected_types, latest_year, prev_year, "新业务CSM（集团口径）", "新业务CSM")
            if df_scatter.empty: st.warning("缺少新业务CSM（集团口径）字段"); return
            use_custom = st.session_state.get("nb_csm_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("nb_csm_scatter_x_bins", "0,50,100,200,500,1000")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("nb_csm_scatter_y_bins", "-100,-50,-20,0,20,50,100")) if use_custom else None
            create_scatter_chart(df_scatter, {'x_col':'新业务CSM','y_col':'新业务CSM增长率','x_bins_custom':x_bins,'y_bins_custom':y_bins}, "新业务CSM区间（亿元）", "新业务CSM增长率区间（%）")
        
        # 图表：新业务LC亏损率散点热力图
        elif m_id == "industry_lc_loss_ratio":
            df_curr = df_raw[df_raw['报告年份'].astype(str) == str(latest_year)]
            df_prev = df_raw[df_raw['报告年份'].astype(str) == str(prev_year)]
            
            scatter_data = []
            for ct in selected_types:
                for co in df_raw[df_raw['公司类型'] == ct]['公司'].unique():
                    def get_val(df, co, field): 
                        val = df[(df['公司'] == co) & (df['字段名'] == field)]['(百万)人民币']
                        return val.iloc[0] if not val.empty else 0
                    
                    # 最新年：LC亏损率 = 新业务亏损合同（LC） / 新业务未来现金流入现值（亏损）
                    lc_curr = get_val(df_curr, co, '新业务亏损合同（LC）——非PAA')
                    pv_loss_curr = get_val(df_curr, co, '新业务未来现金流入现值（亏损）')
                    curr_ratio = lc_curr / pv_loss_curr * 100 if pv_loss_curr != 0 else np.nan
                    
                    # 前一年
                    lc_prev = get_val(df_prev, co, '新业务亏损合同（LC）——非PAA')
                    pv_loss_prev = get_val(df_prev, co, '新业务未来现金流入现值（亏损）')
                    prev_ratio = lc_prev / pv_loss_prev * 100 if pv_loss_prev != 0 else np.nan
                    
                    # 变化（百分点）
                    change = curr_ratio - prev_ratio if not np.isnan(curr_ratio) and not np.isnan(prev_ratio) else np.nan
                    
                    if not np.isnan(curr_ratio):
                        scatter_data.append({
                            '公司': co, 
                            '公司类型': ct, 
                            'LC亏损率': curr_ratio, 
                            'LC亏损率变化': change
                        })
            
            df_scatter = pd.DataFrame(scatter_data)
            if df_scatter.empty: 
                st.warning("缺少新业务亏损合同或新业务未来现金流入现值（亏损）字段")
                return
            
            use_custom = st.session_state.get("lc_loss_ratio_scatter_use_custom_bins", False)
            x_bins = parse_bins_input(st.session_state.get("lc_loss_ratio_scatter_x_bins", "0,10,20,30,40,50,60,70,80,90,100")) if use_custom else None
            y_bins = parse_bins_input(st.session_state.get("lc_loss_ratio_scatter_y_bins", "-50,-30,-20,-10,0,10,20,30,50")) if use_custom else None
            
            create_scatter_chart(df_scatter, {
                'x_col': 'LC亏损率', 
                'y_col': 'LC亏损率变化',
                'x_bins_custom': x_bins, 
                'y_bins_custom': y_bins
            }, "LC亏损率（%）", "LC亏损率变化（百分点）")

        # 新业务IFRS利润率堆叠分布图
        elif m_id == "industry_nb_ifrs_margin":
            if not print_mode:
                c1,c2=st.columns(2)
                with c1: show_labels=st.toggle("显示数值标签",value=True,key=f"lab_{m_id}")
                with c2: label_size=st.slider("标签大小",8,16,11,key=f"sz_{m_id}")
            else: show_labels,label_size=True,11
            use_custom = st.session_state.get(f"{m_id}_dist_use_custom_bins", False)
            x_bins = parse_custom_bins_input(st.session_state.get(f"{m_id}_dist_x_bins","0,10,20,30,40,50")) if use_custom else None
            dist_df,labels=calc_stack_distribution(df_raw,selected_types,latest_year,calc_nb_ifrs_margin,x_bins_custom=x_bins)
            if dist_df.empty: st.warning("无法计算新业务IFRS利润率分布"); return
            create_stack_chart_and_table(dist_df,labels,"新业务IFRS利润率",latest_year,show_labels,label_size)
        else:
            st.info(f"⏳ 模块 [{m_id}] 尚未配置底层绘图代码")
            
    # ==========================================
    # 第5层：报告包装器
    # ==========================================

    def render_report_module(m_id, print_mode):
        """完整报告模块渲染（标题 + 分析框 + 图表 + 注释）"""
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
                        if val and val.lower() != 'nan':
                            title_parts.append(val)
                    if title_parts:
                        full_title = " - ".join(title_parts)

        title_class = "page-break-title" if print_mode else ""
        st.markdown(
            f"<h3 class='{title_class}' style='color:#00338D; font-size:18px; margin-top:20px; margin-bottom:15px; border-bottom:2px solid #00338D; padding-bottom:8px;'>{full_title}</h3>",
            unsafe_allow_html=True
        )

        analysis_text = mod_data.get('analysis', '')
        if analysis_text:
            st.markdown(f"""
            <div style="background-color:#F0F4FA; border-left:4px solid #00338D; padding:15px; margin-bottom:20px; border-radius:4px;">
                <p style="margin:0; color:#1E3A8A; font-size:14px;">{analysis_text}</p>
            </div>
            """, unsafe_allow_html=True)

        render_pure_chart_entity(m_id, print_mode)

        note_text = mod_data.get('note', '')
        if note_text:
            st.markdown(f"""
            <div style="margin-top:20px; padding-left:5px;">
                <p style="margin:0; color:#888; font-size:12px; font-style:italic;">* 注释：{note_text}</p>
            </div>
            """, unsafe_allow_html=True)

    # ==========================================
    # 第6层：主分发逻辑
    # ==========================================

    if print_mode:
        for m_id in ordered_modules:
            render_report_module(m_id, print_mode=True)
    else:
        if active_m_id:
            render_report_module(active_m_id, print_mode=False)
        else:
            st.info("💡 请从左侧导航栏选择要查看的行业分析模块")




