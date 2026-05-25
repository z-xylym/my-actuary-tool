# ==========================================
    # 4. 侧边栏：Step 8 专属级联导航 (防越界护盾版)
    # ==========================================
    print_mode, active_m_id, active_level_context = False, None, ""
    with st.sidebar:
        st.markdown("<h3 style='color: #008578; font-size: 18px;'>行业报告导航</h3>", unsafe_allow_html=True)
        if 'df_notes_8' in st.session_state and st.session_state['df_notes_8'] is not None and not st.session_state['df_notes_8'].empty:
            df_n_8 = st.session_state['df_notes_8']
            first_levels = [x for x in df_n_8['一级分类'].unique() if x]
            main_nav = st.radio("📁 行业大类", first_levels + ["🖨️ 一键显示全部 (行业打印)"], key="s8_m")

            if main_nav == "🖨️ 一键显示全部 (行业打印)":
                print_mode = True
                st.info("💡 行业报告导出模式：打印目标选“另存为PDF”，勾选“背景图形”。")
                components.html("""<button onclick="window.parent.print()" style="width:100%; padding:12px; background:#008578; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold; font-size:14px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">导出 行业分析 PDF</button>""", height=60)
            else:
                df_sub1 = df_n_8[df_n_8['一级分类'] == main_nav]
                sec_levels = [x for x in df_sub1['二级分类'].unique() if x]

                # 🌟 修复点：全部替换为 matched 变量 + .empty 检查
                if len(sec_levels) == 0:
                    charts = [x for x in df_sub1['对应图表名称'].unique() if x]
                    chart_nav = st.radio("具体图表", charts, key="s8_c")
                    matched = df_sub1[df_sub1['对应图表名称'] == chart_nav]
                    active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
                else:
                    sub_nav = st.radio("📂 细分行业视角", ["全部"] + sec_levels, key="s8_s")
                    if sub_nav != "全部":
                        df_sub2 = df_sub1[df_sub1['二级分类'] == sub_nav]
                        charts = [x for x in df_sub2['对应图表名称'].unique() if x]
                        chart_nav = st.radio("📊 行业图表", charts, key="s8_c2")
                        matched = df_sub2[df_sub2['对应图表名称'] == chart_nav]
                        active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
                    else:
                        charts = [x for x in df_sub1['对应图表名称'].unique() if x]
                        chart_nav = st.radio("📊 行业图表 (全部)", charts, key="s8_c_all")
                        matched = df_sub1[df_sub1['对应图表名称'] == chart_nav]
                        active_m_id = matched.iloc[0]['模块ID'] if not matched.empty else None
        else:
            st.warning("⚠️ 请先加载 Step 8 行业注释表")
