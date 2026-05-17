/* ========== 3. 页面布局 ========== */
        
        @page {
            size: A4 portrait;
            margin: 10mm;
        }
        
        html, body {
            margin: 0 !important;
            padding: 0 !important;
        }
        
        .main .block-container {
            max-width: 100% !important; 
            width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        
        .stApp { margin-top: 0 !important; padding-top: 0 !important; }
        .stApp > div:first-child { margin-top: 0 !important; padding-top: 0 !important; }
        
        /* ========== 4. 图表与表格样式 ========== */
        
        .plotly-graph-div {
            page-break-inside: avoid !important;
            margin-bottom: 20px !important;
            display: flex !important;
            justify-content: center !important;
            /* 🚫 核心修复：彻底删除了坑人的 zoom 和 scale！全权交给上方 Python 的 710px 控制！ */
        }
        
        /* 附录的超长明细表属于原生 HTML table，依然需要用 zoom 稍微压一下 */
        div[data-testid="stDataFrame"] {
            zoom: 0.75 !important;
            page-break-inside: avoid !important;
            margin-bottom: 30px !important;
        }
