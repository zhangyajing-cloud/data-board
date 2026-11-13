import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl

# -----------------------------------------------------------------------------
# 1. 页面配置与 CSS / JS 飞书剪贴板导出交互
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="短剧/漫剧数据大盘分析工具", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stApp {
        background-color: #f8f9fa;
        color: #212529 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    h1, h2, h3, h4, h5, h6, p, span, label, div { color: #212529 !important; }
    
    .top-card {
        background-color: #ffffff;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 16px 20px;
        margin-bottom: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .top-card-title {
        font-size: 1.1rem;
        font-weight: 800;
        color: #212529 !important;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .report-header-area {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 10px;
        margin-bottom: 16px;
    }

    .report-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #212529 !important;
        border-left: 6px solid #212529;
        padding-left: 12px;
        margin-top: 28px;
        margin-bottom: 14px;
    }

    .report-subtitle {
        font-size: 1.05rem;
        font-weight: 700;
        color: #343a40 !important;
        margin-top: 16px;
        margin-bottom: 10px;
        padding-left: 8px;
    }

    /* 顶部飞书导出按钮 */
    .btn-feishu-global {
        background-color: #2b579a !important;
        color: #ffffff !important;
        border: none;
        padding: 8px 18px;
        font-size: 0.92rem;
        font-weight: 700 !important;
        border-radius: 6px;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.2s ease;
    }
    .btn-feishu-global:hover { background-color: #1a365d !important; }

    .stButton>button {
        width: 100%;
        background-color: #212529 !important;
        color: #ffffff !important;
        border: none;
        padding: 8px 16px;
        font-size: 0.95rem;
        font-weight: 700 !important;
        border-radius: 6px;
    }
    .stButton>button:hover { background-color: #343a40 !important; }

    div[data-baseweb="select"] * { color: #212529 !important; font-weight: 500; }

    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border: 2px dashed #adb5bd !important;
        border-radius: 8px !important;
        padding: 10px !important;
    }

    /* 全局透视表样式：100% 水平与垂直双向居中、紧凑行高 */
    .pivot-table-container {
        width: 100%;
        overflow-x: auto;
        margin-bottom: 20px;
        border-radius: 6px;
    }
    .full-expand-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 12.5px;
        background-color: #ffffff;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .full-expand-table th {
        background-color: #212529 !important;
        color: #ffffff !important;
        font-weight: bold;
        text-align: center !important;
        vertical-align: middle !important;
        padding: 6px 8px !important;
        border: 1px solid #dee2e6;
        white-space: nowrap;
    }
    .full-expand-table td {
        text-align: center !important;
        vertical-align: middle !important;
        padding: 6px 8px !important;
        border: 1px solid #dee2e6;
        color: #212529;
        white-space: nowrap;
    }
    
    /* 背景色与高亮设定 */
    .row-highlight-yellow {
        background-color: #ffe066 !important;
        color: #000000 !important;
        font-weight: bold;
    }
    .row-cp-subtotal {
        background-color: #fff9db !important;
        font-weight: bold;
    }
    .row-lang-subtotal {
        background-color: #fff9c4 !important;
        font-weight: bold;
    }
    .row-grand-total {
        background-color: #ffe066 !important;
        color: #000000 !important;
        font-weight: bold;
    }
    .cell-passed {
        background-color: #d4edda !important;
        color: #28a745 !important;
        font-weight: bold;
    }
    .cell-passed-text-only {
        color: #28a745 !important;
        font-weight: bold;
    }
    .cell-passed-underline {
        color: #28a745 !important;
        font-weight: bold;
        text-decoration: underline;
    }
    .cell-last-week {
        color: #6c757d !important;
    }

    /* Streamlit 原生 Tabs 组件定制 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
        padding-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 38px;
        background-color: #ffffff;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        color: #495057 !important;
        font-weight: 600;
        padding: 0 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #212529 !important;
        color: #ffffff !important;
        border-color: #212529 !important;
        font-weight: bold !important;
    }

    /* 文字复盘卡片 */
    .review-card {
        background-color: #ffffff;
        border-left: 4px solid #212529;
        border-radius: 6px;
        padding: 14px 18px;
        margin-top: 12px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .review-item {
        margin-bottom: 6px;
        font-size: 13.5px;
        line-height: 1.5;
        color: #212529 !important;
    }
</style>

<script>
function copyFeishuRichTextGlobal(btnElement) {
    var htmlContent = "";
    var containers = document.querySelectorAll(".pivot-table-container");
    if (containers.length === 0) {
        alert("暂无渲染好的表格可供导出！");
        return;
    }
    containers.forEach(function(container) {
        htmlContent += container.innerHTML + "<br><br>";
    });

    var blobHtml = new Blob([htmlContent], { type: "text/html" });
    var blobText = new Blob([htmlContent.replace(/<[^>]+>/g, '')], { type: "text/plain" });

    if (navigator.clipboard && window.ClipboardItem) {
        var data = [new ClipboardItem({
            "text/html": blobHtml,
            "text/plain": blobText
        })];
        navigator.clipboard.write(data).then(function() {
            var originalText = btnElement.innerHTML;
            btnElement.innerHTML = "✅ 已复制！请直接在飞书文档按 Cmd+V 粘贴";
            btnElement.style.backgroundColor = "#059669";
            setTimeout(function() {
                btnElement.innerHTML = originalText;
                btnElement.style.backgroundColor = "#2b579a";
            }, 3000);
        }).catch(function(err) {
            fallbackCopyText(htmlContent, btnElement);
        });
    } else {
        fallbackCopyText(htmlContent, btnElement);
    }
}

function fallbackCopyText(text, btnElement) {
    var textarea = document.createElement("textarea");
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
    
    var originalText = btnElement.innerHTML;
    btnElement.innerHTML = "✅ 已复制 HTML 文本！可在飞书粘贴";
    btnElement.style.backgroundColor = "#059669";
    setTimeout(function() {
        btnElement.innerHTML = originalText;
        btnElement.style.backgroundColor = "#2b579a";
    }, 3000);
}
</script>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. 数据清洗与工具函数
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner="⚡ 正在极速加载 Excel 数据至内存...")
def parse_excel_fast(file_bytes):
    if file_bytes is None: return None
    df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    return df

def clean_num(series):
    if series is None: return pd.Series(0)
    return pd.to_numeric(
        series.astype(str).str.replace('$', '', regex=False)
                          .str.replace(',', '', regex=False)
                          .str.replace('%', '', regex=False).str.strip(), 
        errors='coerce'
    ).fillna(0)

def find_best_match(df_cols, keywords, exact_first=True):
    if exact_first:
        for kw in keywords:
            for col in df_cols:
                if kw.lower() == str(col).strip().lower(): return col
    for kw in keywords:
        for col in df_cols:
            if kw.lower() in str(col).lower(): return col
    return "不映射/无此列"

def format_diff_with_arrow(val_pct, is_pp=False):
    if pd.isna(val_pct) or val_pct == 0:
        return "-", "color:#212529;"
    
    suffix = "pp" if is_pp else "%"
    if val_pct > 0:
        return f"+{val_pct:.2f}{suffix} ↑", "color:#28a745;font-weight:bold;"
    else:
        return f"{val_pct:.2f}{suffix} ↓", "color:#dc3545;font-weight:bold;"

def get_account_timezone(account_name):
    if pd.isna(account_name):
        return '美西账户'
    acc_str = str(account_name)
    if '+8-' in acc_str:
        return '+8 时区账户'
    elif '+0-' in acc_str:
        return '0 时区账户'
    else:
        return '美西账户'

# -----------------------------------------------------------------------------
# 3. 页面 Header 与全局导出控制区
# -----------------------------------------------------------------------------
st.markdown('''
<div class="report-header-area">
    <h2 style="margin:0; font-weight:800;">📊 短剧/漫剧数据全维度分析对比工具</h2>
    <button class="btn-feishu-global" onclick="copyFeishuRichTextGlobal(this)">🚀 导出完整报告至飞书文档</button>
</div>
''', unsafe_allow_html=True)

lang_tabs = st.tabs(["🌐 全部地区/双语种", "🇹🇼 繁体中文", "🇰🇷 韩语"])

st.markdown('<div class="top-card">', unsafe_allow_html=True)
st.markdown('<div class="top-card-title">📁 第一步：上传数据表格 (上周基准表 vs 本周对比表)</div>', unsafe_allow_html=True)

col_u1, col_u2 = st.columns(2)
with col_u1: file_w1 = st.file_uploader("上传【上周】基准表", type=["xlsx", "xls"])
with col_u2: file_w2 = st.file_uploader("上传【本周】对比表", type=["xlsx", "xls"])

st.markdown('</div>', unsafe_allow_html=True)

if file_w1 and file_w2:
    raw_df_w1 = parse_excel_fast(file_w1.getvalue())
    raw_df_w2 = parse_excel_fast(file_w2.getvalue())
    all_cols = ["不映射/无此列"] + list(dict.fromkeys(list(raw_df_w1.columns) + list(raw_df_w2.columns)))
    
    st.markdown('<div class="top-card">', unsafe_allow_html=True)
    st.markdown('<div class="top-card-title">⚙️ 第二步：指标维度映射选择 (已自动智能识别)</div>', unsafe_allow_html=True)
    
    m1, m2, m3, m4 = st.columns(4)
    with m1: col_lang = st.selectbox("语言/渠道列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['语言', '平台', '渠道', '地区'])))
    with m2: col_spend = st.selectbox("花费/消耗列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['花费', '消耗', 'cost'])))
    with m3: col_d0_rev = st.selectbox("D0 Revenue列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['d0_revenue', 'd0 revenue', 'd0收入'])))
    with m4: col_d0_target = st.selectbox("D0 目标收入列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['D0目标收入', '目标收入'])))

    m5, m6, m7, m8 = st.columns(4)
    with m5: col_act = st.selectbox("激活用户数列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['首日激活用户数', '首日激活', '激活用户数', '激活'])))
    with m6: col_pay = st.selectbox("付费用户数列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['首日付费用户数', '首日付费', '付费用户数', '付费'])))
    with m7: col_reg_rate = st.selectbox("新注册用户占比:", all_cols, index=all_cols.index(find_best_match(all_cols, ['新注册用户占比', '新注册占比'])))
    with m8: col_pay_reg_rate = st.selectbox("付费新注册占比:", all_cols, index=all_cols.index(find_best_match(all_cols, ['付费用户新注册用户占比', '拉新付费占比', '付费新注册占比'])))

    st.markdown("##### 🎭 模块二追加字段映射 (剧集测试与明细复盘)")
    f1, f2, f3, f4 = st.columns(4)
    with f1: col_title = st.selectbox("剧目/书籍名称列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['剧目名称', '书籍名称', '剧目', '书名'])))
    with f2: col_cn_title = st.selectbox("中文推广剧目名称列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['中文推广剧目', '中文剧名', '中文名', '中文'])))
    with f3: col_en_title = st.selectbox("英文推广剧目名称列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['英文推广剧目', '英文剧名', '英文名', '英文'])))
    with f4: col_book_source = st.selectbox("书籍来源列 (表格1类型):", all_cols, index=all_cols.index(find_best_match(all_cols, ['书籍来源', '剧目来源'])))

    f5, f6, f7, f8 = st.columns(4)
    with f5: col_copyright = st.selectbox("版权来源列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['版权来源', '版权'])))
    with f6: col_manga_type = st.selectbox("漫剧类型列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['漫剧类型', '类型'])))
    with f7: col_manga_cls = st.selectbox("漫剧分类列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['漫剧分类', '分类'])))
    with f8: col_optimizer = st.selectbox("优化师列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['优化师', '负责人', '优化师姓名'])))

    st.markdown("##### 🌐 板块一/二/每日趋势/Pixel & 优化方式/操作系统扩展映射")
    e1, e2, e3, e4 = st.columns(4)
    with e1: col_account = st.selectbox("账户名称/投放账户列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['账户名称', '投放账户', '账户', 'account'])))
    with e2: col_platform = st.selectbox("投放平台列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['投放平台', '平台', 'media_source'])))
    with e3: col_target = st.selectbox("推广目标列 (w2a/app):", all_cols, index=all_cols.index(find_best_match(all_cols, ['推广目标', '目标', 'campaign_type'])))
    with e4: col_pixel = st.selectbox("Pixel 名称列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['pixel名称', 'pixel', '像素'])))

    e5, e6, e7, e8 = st.columns(4)
    with e5: col_opt_type = st.selectbox("优化方式列 (AEO/VO):", all_cols, index=all_cols.index(find_best_match(all_cols, ['优化方式', '优化目标', 'optimization_goal'])))
    with e6: col_bid_val = st.selectbox("手动出价金额/系数列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['手动出价的金额/系数', '手动出价', '出价系数', 'bid'])))
    with e7: col_date = st.selectbox("时间/日期列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['日期', '时间', 'date'])))
    with e8: col_os = st.selectbox("操作系统 (广告)列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['操作系统 (广告)', '操作系统', '系统', 'os', 'platform_os'])))

    st.markdown('</div>', unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # ⚡ 4. 向量化数据预计算缓存
    # -------------------------------------------------------------------------
    @st.cache_data(ttl=3600, show_spinner=False)
    def prepare_precomputed_data(raw_w1, raw_w2, col_map_tuple):
        (c_lang, c_spend, c_rev, c_tar, c_act, c_pay, c_reg, c_pay_reg, c_title, c_date) = col_map_tuple
        w2 = raw_w2.copy()
        w1 = raw_w1.copy()

        w2['spend'] = clean_num(w2[c_spend]) if c_spend in w2 else 0.0
        w2['d0_rev'] = clean_num(w2[c_rev]) if c_rev in w2 else 0.0
        w2['d0_target'] = clean_num(w2[c_tar]) if c_tar in w2 else 0.0
        w2['act'] = clean_num(w2[c_act]) if c_act in w2 else 0.0
        w2['pay'] = clean_num(w2[c_pay]) if c_pay in w2 else 0.0

        w1['spend'] = clean_num(w1[c_spend]) if c_spend in w1 else 0.0
        w1['d0_rev'] = clean_num(w1[c_rev]) if c_rev in w1 else 0.0
        w1['d0_target'] = clean_num(w1[c_tar]) if c_tar in w1 else 0.0
        w1['act'] = clean_num(w1[c_act]) if c_act in w1 else 0.0
        w1['pay'] = clean_num(w1[c_pay]) if c_pay in w1 else 0.0

        # 剔除零花费
        w2 = w2[w2['spend'] > 0].copy()
        w1 = w1[w1['spend'] > 0].copy()

        for df_tmp in [w1, w2]:
            reg_r = clean_num(df_tmp[c_reg]) if c_reg in df_tmp else pd.Series(0, index=df_tmp.index)
            pay_reg_r = clean_num(df_tmp[c_pay_reg]) if c_pay_reg in df_tmp else pd.Series(0, index=df_tmp.index)
            if reg_r.max() > 1.0: reg_r /= 100.0
            if pay_reg_r.max() > 1.0: pay_reg_r /= 100.0
            df_tmp['new_reg'] = df_tmp['act'] * reg_r
            df_tmp['new_pay'] = df_tmp['pay'] * pay_reg_r

        # 首投日期计算
        if c_title != "不映射/无此列" and c_title in w2.columns:
            w2[c_title] = w2[c_title].astype(str).str.strip()
            if c_title in w1.columns:
                w1[c_title] = w1[c_title].astype(str).str.strip()
            
            if c_date != "不映射/无此列" and c_date in w2.columns:
                all_hist = pd.concat([w1, w2], ignore_index=True)
                all_hist['date_dt'] = pd.to_datetime(all_hist[c_date], errors='coerce')
                first_launch = all_hist.dropna(subset=['date_dt']).groupby(c_title)['date_dt'].min().reset_index()
                first_launch.rename(columns={'date_dt': '首次投放日期'}, inplace=True)

                w2['date_dt'] = pd.to_datetime(w2[c_date], errors='coerce')
                valid_w2_dates = w2['date_dt'].dropna()
                w_start = valid_w2_dates.min() if not valid_w2_dates.empty else pd.Timestamp.min
                w_end = valid_w2_dates.max() if not valid_w2_dates.empty else pd.Timestamp.max

                w2 = pd.merge(w2, first_launch, on=c_title, how='left')
                w2['is_new_drama'] = w2['首次投放日期'].apply(lambda x: True if (pd.notna(x) and w_start <= x <= w_end) else False)
            else:
                w1_titles = set(w1[c_title].dropna().unique()) if c_title in w1.columns else set()
                w2['is_new_drama'] = w2[c_title].apply(lambda x: x not in w1_titles)
        else:
            w2['is_new_drama'] = False

        return w1, w2

    col_tuple = (col_lang, col_spend, col_d0_rev, col_d0_target, col_act, col_pay, col_reg_rate, col_pay_reg_rate, col_title, col_date)
    pre_w1, pre_w2 = prepare_precomputed_data(raw_df_w1, raw_df_w2, col_tuple)

    # -------------------------------------------------------------------------
    # 5. 动态渲染全量视图
    # -------------------------------------------------------------------------
    def render_dashboard_content(target_lang_filter=None):
        if target_lang_filter and col_lang != "不映射/无此列" and col_lang in pre_w2.columns:
            df_w2 = pre_w2[pre_w2[col_lang].astype(str).str.contains(target_lang_filter, na=False)].copy()
            df_w1 = pre_w1[pre_w1[col_lang].astype(str).str.contains(target_lang_filter, na=False)].copy() if col_lang in pre_w1.columns else pre_w1.copy()
        else:
            df_w2 = pre_w2.copy()
            df_w1 = pre_w1.copy()

        # =====================================================================
        # 一、整体数据情况
        # =====================================================================
        st.markdown('<div class="report-title">一、整体数据情况</div>', unsafe_allow_html=True)

        def get_totals(df):
            sp = df['spend'].sum()
            rev = df['d0_rev'].sum()
            tar = df['d0_target'].sum()
            act = df['act'].sum()
            pay = df['pay'].sum()
            return {
                'spend': sp, 
                'd0_rate': rev / tar if tar > 0 else 0.0, 
                'roas': rev / sp if sp > 0 else 0.0, 
                'pay_reg_pct': df['new_pay'].sum() / pay if pay > 0 else 0.0, 
                'reg_pct': df['new_reg'].sum() / act if act > 0 else 0.0
            }

        tot_w1 = get_totals(df_w1)
        tot_w2 = get_totals(df_w2)

        st.markdown('<div class="report-subtitle">📊 1. 大盘本周环比情况</div>', unsafe_allow_html=True)
        
        summary_rows = [
            ("花费 (消耗金额)", tot_w1['spend'], tot_w2['spend'], "money"),
            ("D0 达标率", tot_w1['d0_rate'], tot_w2['d0_rate'], "percent"),
            ("D0 ROAS", tot_w1['roas'], tot_w2['roas'], "percent"),
            ("付费用户新注册占比", tot_w1['pay_reg_pct'], tot_w2['pay_reg_pct'], "percent"),
            ("新注册用户占比", tot_w1['reg_pct'], tot_w2['reg_pct'], "percent"),
        ]

        html_dapan = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
        html_dapan.append('<th>指标维度</th><th>本周</th><th>上周</th><th>双周环比 (涨跌幅)</th></tr></thead><tbody>')

        is_first_row = True
        for name, v1, v2, kind in summary_rows:
            diff_abs = v2 - v1
            if kind == "money":
                f1, f2 = f"${v1:,.2f}", f"${v2:,.2f}"
                ratio = (diff_abs / v1 * 100) if v1 > 0 else 0
                diff_str, diff_style = format_diff_with_arrow(ratio)
            else:
                f1, f2 = f"{v1*100:.2f}%", f"{v2*100:.2f}%"
                ratio = (diff_abs / v1 * 100) if v1 > 0 else 0
                pp = diff_abs * 100
                diff_str_pct, diff_style = format_diff_with_arrow(ratio)
                pp_str, _ = format_diff_with_arrow(pp, is_pp=True)
                diff_str = f"{diff_str_pct} ({pp_str})"

            row_class = ' class="row-highlight-yellow"' if is_first_row else ''
            is_first_row = False

            html_dapan.append(f'<tr{row_class}>')
            html_dapan.append(f'<td style="font-weight:bold;">{name}</td>')
            html_dapan.append(f'<td>{f2}</td>')
            html_dapan.append(f'<td class="cell-last-week">{f1}</td>')
            html_dapan.append(f'<td style="{diff_style}">{diff_str}</td>')
            html_dapan.append('</tr>')

        html_dapan.append('</tbody></table></div>')
        st.markdown("".join(html_dapan), unsafe_allow_html=True)

        st.markdown('<div class="report-subtitle">📊 2. 双语种本周数据情况</div>', unsafe_allow_html=True)

        if col_lang != "不映射/无此列" and col_lang in df_w2.columns:
            g2 = df_w2.groupby(col_lang, as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum', 'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'}).sort_values('spend', ascending=False)
            g1 = df_w1.groupby(col_lang, as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum'}) if col_lang in df_w1.columns else pd.DataFrame(columns=[col_lang, 'spend', 'd0_rev', 'd0_target'])
            
            m = pd.merge(g2, g1, on=col_lang, how='left', suffixes=('', '_w1')).fillna(0)
            total_spend_w2 = m['spend'].sum()
            m['消耗占比'] = m['spend'] / total_spend_w2 if total_spend_w2 > 0 else 0
            m['消耗环比'] = (m['spend'] - m['spend_w1']) / m['spend_w1'].replace(0, np.nan)
            m['本周达标率'] = m['d0_rev'] / m['d0_target'].replace(0, np.nan)
            d0_rate_w1 = m['d0_rev_w1'] / m['d0_target_w1'].replace(0, np.nan)
            m['达标率环比'] = m['本周达标率'] - d0_rate_w1
            m['本周 D0 ROAS'] = m['d0_rev'] / m['spend'].replace(0, np.nan)
            roas_w1 = m['d0_rev_w1'] / m['spend_w1'].replace(0, np.nan)
            m['D0 ROAS 环比'] = m['本周 D0 ROAS'] - roas_w1
            m['新注册用户率'] = m['new_reg'] / m['act'].replace(0, np.nan)
            m['新注册用户付费率'] = m['new_pay'] / m['pay'].replace(0, np.nan)
            
            total_sp_w1 = m['spend_w1'].sum()
            total_rev_w2, total_rev_w1 = m['d0_rev'].sum(), m['d0_rev_w1'].sum()
            total_tar_w2, total_tar_w1 = m['d0_target'].sum(), m['d0_target_w1'].sum()
            total_act, total_pay = m['act'].sum(), m['pay'].sum()
            total_new_reg, total_new_pay = m['new_reg'].sum(), m['new_pay'].sum()
            
            tot_d0_w2 = total_rev_w2 / total_tar_w2 if total_tar_w2 > 0 else 0
            tot_d0_w1 = total_rev_w1 / total_tar_w1 if total_tar_w1 > 0 else 0
            tot_roas_w2 = total_rev_w2 / total_spend_w2 if total_spend_w2 > 0 else 0
            tot_roas_w1 = total_rev_w1 / total_sp_w1 if total_sp_w1 > 0 else 0
            
            html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
            html_list.append('<th>剧目/书籍语言</th><th>本周消耗 ($)</th><th>消耗占比</th><th>消耗环比</th><th>本周达标率 (%)</th><th>达标率环比</th><th>本周 D0 ROAS (%)</th><th>D0 ROAS 环比</th><th>新注册用户率 (%)</th><th>新注册用户付费率 (%)</th></tr></thead><tbody>')

            for _, r in m.iterrows():
                sp = r['spend']
                p_pct = r['消耗占比'] * 100
                p_diff = r['消耗环比'] * 100 if pd.notnull(r['消耗环比']) else 0
                d0_v = r['本周达标率'] * 100
                d0_diff = r['达标率环比'] * 100
                roas_v = r['本周 D0 ROAS'] * 100
                roas_diff = r['D0 ROAS 环比'] * 100
                reg_v = r['新注册用户率'] * 100
                pay_reg_v = r['新注册用户付费率'] * 100

                passed_cell_class = ' class="cell-passed"' if d0_v >= 85.0 else (' class="cell-passed-text-only"' if d0_v >= 80.0 else '')
                
                p_diff_str, p_diff_style = format_diff_with_arrow(p_diff)
                d0_diff_str, d0_diff_style = format_diff_with_arrow(d0_diff)
                roas_diff_str, roas_diff_style = format_diff_with_arrow(roas_diff)

                html_list.append('<tr>')
                html_list.append(f'<td style="font-weight:bold;">{r[col_lang]}</td>')
                html_list.append(f'<td>${sp:,.2f}</td>')
                html_list.append(f'<td>{p_pct:.2f}%</td>')
                html_list.append(f'<td style="{p_diff_style}">{p_diff_str}</td>')
                html_list.append(f'<td{passed_cell_class}>{d0_v:.2f}%</td>')
                html_list.append(f'<td style="{d0_diff_style}">{d0_diff_str}</td>')
                html_list.append(f'<td>{roas_v:.2f}%</td>')
                html_list.append(f'<td style="{roas_diff_style}">{roas_diff_str}</td>')
                html_list.append(f'<td>{reg_v:.2f}%</td>')
                html_list.append(f'<td>{pay_reg_v:.2f}%</td>')
                html_list.append('</tr>')

            tot_p_diff = (total_spend_w2 - total_sp_w1) / total_sp_w1 * 100 if total_sp_w1 > 0 else 0
            tot_d0_v = tot_d0_w2 * 100
            tot_d0_diff = (tot_d0_w2 - tot_d0_w1) * 100
            tot_roas_v = tot_roas_w2 * 100
            tot_roas_diff = (tot_roas_w2 - tot_roas_w1) * 100
            tot_reg_v = (total_new_reg / total_act * 100) if total_act > 0 else 0
            tot_pay_reg_v = (total_new_pay / total_pay * 100) if total_pay > 0 else 0

            tot_passed_class = ' class="cell-passed-underline"' if tot_d0_v >= 85.0 else (' class="cell-passed-text-only"' if tot_d0_v >= 80.0 else '')
            
            tot_p_str, tot_p_style = format_diff_with_arrow(tot_p_diff)
            tot_d0_str, tot_d0_style = format_diff_with_arrow(tot_d0_diff)
            tot_roas_str, tot_roas_style = format_diff_with_arrow(tot_roas_diff)

            html_list.append('<tr class="row-highlight-yellow"><td>总计</td>')
            html_list.append(f'<td>${total_spend_w2:,.2f}</td><td>100.00%</td>')
            html_list.append(f'<td style="{tot_p_style}">{tot_p_str}</td>')
            html_list.append(f'<td{tot_passed_class}>{tot_d0_v:.2f}%</td>')
            html_list.append(f'<td style="{tot_d0_style}">{tot_d0_str}</td>')
            html_list.append(f'<td>{tot_roas_v:.2f}%</td>')
            html_list.append(f'<td style="{tot_roas_style}">{tot_roas_str}</td>')
            html_list.append(f'<td>{tot_reg_v:.2f}%</td>')
            html_list.append(f'<td>{tot_pay_reg_v:.2f}%</td></tr>')

            html_list.append('</tbody></table></div>')
            st.markdown("".join(html_list), unsafe_allow_html=True)

        # =====================================================================
        # 二、具体投放数据和剧集情况复盘
        # =====================================================================
        st.markdown('<div class="report-title">二、具体投放数据和剧集情况复盘</div>', unsafe_allow_html=True)

        required_m2_cols = [col_lang, col_copyright, col_manga_type, col_manga_cls, col_title, col_spend, col_d0_rev, col_d0_target]
        missing_cols = [c for c in required_m2_cols if c == "不映射/无此列" or c not in df_w2.columns]

        if missing_cols:
            st.warning(f"⚠️ 需要完成追加字段映射，请在上方“第二步”设置：剧目名称、版权来源、漫剧类型、漫剧分类！")
        else:
            w2 = df_w2.copy()
            w2['cn_title'] = w2[col_cn_title].astype(str).str.strip() if col_cn_title in w2 else "-"
            w2['en_title'] = w2[col_en_title].astype(str).str.strip() if col_en_title in w2 else "-"
            
            # 📌 专门清洗【书籍来源】列用于表格 1
            if col_book_source != "不映射/无此列" and col_book_source in w2.columns:
                w2['book_source_clean'] = w2[col_book_source].fillna('未知来源').astype(str).str.strip()
            else:
                w2['book_source_clean'] = w2[col_manga_type].fillna('未知来源').astype(str).str.strip()

            show_level = w2.groupby([col_lang, col_copyright, col_manga_type, col_manga_cls, col_title, 'cn_title', 'en_title', 'book_source_clean'], as_index=False).agg({
                'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum',
                'is_new_drama': 'first'
            })
            show_level = show_level[show_level['spend'] > 0]
            show_level['d0_rate'] = np.where(show_level['d0_target'] > 0, show_level['d0_rev'] / show_level['d0_target'], 0)
            show_level['is_passed'] = show_level['d0_rate'] >= 0.80
            show_level['is_new_passed'] = show_level['is_new_drama'] & show_level['is_passed']

            # =================================================================
            # 🚨 修复完成：📊 1. 本周剧集类型情况及过测情况 - 语种类型汇总表 (按【书籍来源】聚合)
            # =================================================================
            st.markdown('<div class="report-subtitle">📊 1. 本周剧集类型情况及过测情况 - 语种类型汇总表</div>', unsafe_allow_html=True)

            def render_lang_type_summary_table(show_level):
                if show_level.empty: return

                html_lt = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_lt.append('<th style="width: 12%;">剧目/书籍语言</th>')
                html_lt.append('<th style="width: 18%;">类型</th>')
                html_lt.append('<th style="width: 12%;">消耗占比 (%)</th>')
                html_lt.append('<th style="width: 14%;">花费 ($)</th>')
                html_lt.append('<th style="width: 12%;">D0达标率 (%)</th>')
                html_lt.append('<th style="width: 10%;">剧集数量</th>')
                html_lt.append('<th style="width: 10%;">新剧数量</th>')
                html_lt.append('<th style="width: 12%;">新剧测出率 (%)</th>')
                html_lt.append('</tr></thead><tbody>')

                lang_grouped = sorted(show_level.groupby(col_lang), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    total_lang_cost = group_lang['spend'].sum()
                    
                    # 📌 100% 绑定底表的【书籍来源】列 (book_source_clean) 进行聚合计算
                    type_rows = []
                    for source_val, g_type in group_lang.groupby('book_source_clean'):
                        t_sp = g_type['spend'].sum()
                        t_rev = g_type['d0_rev'].sum()
                        t_tar = g_type['d0_target'].sum()
                        
                        t_d0_rate = (t_rev / t_tar * 100) if t_tar > 0 else 0.0
                        t_shows = g_type[col_title].nunique()
                        t_new_shows = g_type[g_type['is_new_drama']][col_title].nunique()
                        t_new_passed = g_type[g_type['is_new_passed']][col_title].nunique()
                        t_new_rate = (t_new_passed / t_new_shows * 100) if t_new_shows > 0 else 0.0
                        t_cost_pct = (t_sp / total_lang_cost * 100) if total_lang_cost > 0 else 0.0

                        type_rows.append({
                            'source': source_val,
                            'cost_pct': t_cost_pct,
                            'spend': t_sp,
                            'd0_rate': t_d0_rate,
                            'shows': t_shows,
                            'new_shows': t_new_shows,
                            'new_rate': t_new_rate
                        })

                    type_rows.sort(key=lambda x: x['spend'], reverse=True)
                    rowspan_cnt = len(type_rows)
                    first_r = True

                    for r in type_rows:
                        d0_v = r['d0_rate']
                        passed_cls = ' class="cell-passed"' if d0_v >= 85.0 else (' class="cell-passed-text-only"' if d0_v >= 80.0 else '')
                        new_rate_str = f"{r['new_rate']:.2f}%" if r['new_shows'] > 0 else "0.00%"

                        html_lt.append('<tr>')
                        if show_lang_col and first_r:
                            html_lt.append(f'<td rowspan="{rowspan_cnt}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_r = False

                        html_lt.append(f'<td style="font-weight:600;">{r["source"]}</td>')
                        html_lt.append(f'<td>{r["cost_pct"]:.2f}%</td>')
                        html_lt.append(f'<td style="font-weight:600;">${r["spend"]:,.2f}</td>')
                        html_lt.append(f'<td{passed_cls}>{d0_v:.2f}%</td>')
                        html_lt.append(f'<td>{r["shows"]}</td>')
                        html_lt.append(f'<td>{r["new_shows"]}</td>')
                        html_lt.append(f'<td>{new_rate_str}</td>')
                        html_lt.append('</tr>')

                    # 语种小计行
                    l_rev = group_lang['d0_rev'].sum()
                    l_tar = group_lang['d0_target'].sum()
                    l_d0 = (l_rev / l_tar * 100) if l_tar > 0 else 0.0
                    l_shows = group_lang[col_title].nunique()
                    l_new_shows = group_lang[group_lang['is_new_drama']][col_title].nunique()
                    l_new_passed = group_lang[group_lang['is_new_passed']][col_title].nunique()
                    l_new_rate = (l_new_passed / l_new_shows * 100) if l_new_shows > 0 else 0.0
                    l_passed_cls = ' class="cell-passed-underline"' if l_d0 >= 85.0 else (' class="cell-passed-text-only"' if l_d0 >= 80.0 else '')
                    
                    span_sub = 2 if show_lang_col else 1
                    html_lt.append(f'<tr class="row-lang-subtotal">')
                    html_lt.append(f'<td colspan="{span_sub}">{lang_val} 小计</td>')
                    html_lt.append(f'<td>100.00%</td>')
                    html_lt.append(f'<td>${total_lang_cost:,.2f}</td>')
                    html_lt.append(f'<td{l_passed_cls}>{l_d0:.2f}%</td>')
                    html_lt.append(f'<td>{l_shows}</td>')
                    html_lt.append(f'<td>{l_new_shows}</td>')
                    html_lt.append(f'<td>{l_new_rate:.2f}%</td>')
                    html_lt.append('</tr>')

                html_lt.append('</tbody></table></div>')
                st.markdown("".join(html_lt), unsafe_allow_html=True)

            render_lang_type_summary_table(show_level)

            # 📊 2. 本周剧集类型情况及过测情况 - 类型拆分表
            st.markdown('<div class="report-subtitle">📊 2. 本周剧集类型情况及过测情况 - 类型拆分表</div>', unsafe_allow_html=True)

            def render_table2_4level_pivot(show_level):
                tbl2_raw = []
                for (lang_val, cp_val, mtype_val, mcls_val), group in show_level.groupby([col_lang, col_copyright, col_manga_type, col_manga_cls]):
                    spend_total = group['spend'].sum()
                    total_shows = group[col_title].nunique()
                    new_shows = group[group['is_new_drama']][col_title].nunique()
                    new_passed = group[group['is_new_passed']][col_title].nunique()
                    active_passed = group[group['is_passed']][col_title].nunique()

                    tbl2_raw.append({
                        'lang': lang_val,
                        'cp': cp_val,
                        'mtype': mtype_val,
                        'mcls': mcls_val,
                        'spend': spend_total,
                        'total_shows': total_shows,
                        'new_shows': new_shows,
                        'new_passed': new_passed,
                        'active_passed': active_passed
                    })

                df_raw = pd.DataFrame(tbl2_raw)
                if df_raw.empty: return

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th style="width: 12%;">剧目/书籍语言</th>')
                html_list.append('<th style="width: 10%;">版权来源</th><th style="width: 10%;">漫剧类型</th><th style="width: 10%;">漫剧分类</th><th style="width: 12%;">本周花费 ($)</th><th style="width: 8%;">在投总数</th><th style="width: 8%;">新剧数量</th><th style="width: 8%;">新剧测出数</th><th style="width: 10%;">新剧测出率 (%)</th><th style="width: 12%;">在投剧目达标率 (%)</th></tr></thead><tbody>')

                g_spend, g_tot_shows, g_new_shows, g_new_passed, g_act_passed = 0, 0, 0, 0, 0
                lang_grouped = sorted(df_raw.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    lang_spend, lang_tot, lang_new, lang_n_pass, lang_a_pass = 0, 0, 0, 0, 0
                    lang_rows_count = len(group_lang) + group_lang['cp'].nunique()
                    first_lang_row = True
                    cp_grouped = sorted(group_lang.groupby('cp'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                    for cp_val, group_cp in cp_grouped:
                        cp_sorted = group_cp.sort_values(by='spend', ascending=False)
                        cp_spend, cp_tot, cp_new, cp_n_pass, cp_a_pass = cp_sorted['spend'].sum(), cp_sorted['total_shows'].sum(), cp_sorted['new_shows'].sum(), cp_sorted['new_passed'].sum(), cp_sorted['active_passed'].sum()
                        lang_spend += cp_spend; lang_tot += cp_tot; lang_new += cp_new; lang_n_pass += cp_n_pass; lang_a_pass += cp_a_pass
                        cp_rows_count = len(cp_sorted)
                        first_cp_row = True

                        for _, row in cp_sorted.iterrows():
                            sp, n_shows, n_pass, a_shows, a_pass = row['spend'], row['new_shows'], row['new_passed'], row['total_shows'], row['active_passed']
                            n_rate = (n_pass / n_shows * 100) if n_shows > 0 else 0.0
                            a_rate = (a_pass / a_shows * 100) if a_shows > 0 else 0.0
                            a_rate_class = ' class="cell-passed"' if a_rate >= 85.0 else (' class="cell-passed-text-only"' if a_rate >= 80.0 else '')

                            html_list.append('<tr>')
                            if show_lang_col and first_lang_row:
                                html_list.append(f'<td rowspan="{lang_rows_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                                first_lang_row = False
                            if first_cp_row:
                                html_list.append(f'<td rowspan="{cp_rows_count}" style="background-color:#ffffff;">{cp_val}</td>')
                                first_cp_row = False

                            html_list.append(f'<td>{row["mtype"]}</td><td>{row["mcls"]}</td><td style="font-weight:600;">${sp:,.2f}</td><td>{a_shows}</td><td>{n_shows}</td><td>{n_pass}</td><td>{n_rate:.2f}%</td><td{a_rate_class}>{a_rate:.2f}%</td></tr>')

                        cp_n_rate = (cp_n_pass / cp_new * 100) if cp_new > 0 else 0.0
                        cp_a_rate = (cp_a_pass / cp_tot * 100) if cp_tot > 0 else 0.0
                        cp_a_class = ' class="cell-passed-underline"' if cp_a_rate >= 85.0 else (' class="cell-passed-text-only"' if cp_a_rate >= 80.0 else '')
                        span_val = 3 if show_lang_col else 2
                        html_list.append(f'<tr class="row-cp-subtotal"><td colspan="{span_val}">{cp_val} 小计</td><td>${cp_spend:,.2f}</td><td>{cp_tot}</td><td>{cp_new}</td><td>{cp_n_pass}</td><td>{cp_n_rate:.2f}%</td><td{cp_a_class}>{cp_a_rate:.2f}%</td></tr>')

                    g_spend += lang_spend; g_tot_shows += lang_tot; g_new_shows += lang_new; g_new_passed += lang_n_pass; g_act_passed += lang_a_pass
                    lang_n_rate = (lang_n_pass / lang_new * 100) if lang_new > 0 else 0.0
                    lang_a_rate = (lang_a_pass / lang_tot * 100) if lang_tot > 0 else 0.0
                    lang_a_class = ' class="cell-passed-underline"' if lang_a_rate >= 85.0 else (' class="cell-passed-text-only"' if lang_a_rate >= 80.0 else '')
                    span_val_lang = 4 if show_lang_col else 3
                    html_list.append(f'<tr class="row-lang-subtotal"><td colspan="{span_val_lang}">{lang_val} 小计</td><td>${lang_spend:,.2f}</td><td>{lang_tot}</td><td>{lang_new}</td><td>{lang_n_pass}</td><td>{lang_n_rate:.2f}%</td><td{lang_a_class}>{lang_a_rate:.2f}%</td></tr>')

                grand_n_rate = (g_new_passed / g_new_shows * 100) if g_new_shows > 0 else 0.0
                grand_a_rate = (g_act_passed / g_tot_shows * 100) if g_tot_shows > 0 else 0.0
                grand_a_class = ' class="cell-passed-underline"' if grand_a_rate >= 85.0 else (' class="cell-passed-text-only"' if grand_a_rate >= 80.0 else '')
                span_val_tot = 4 if show_lang_col else 3
                html_list.append(f'<tr class="row-grand-total"><td colspan="{span_val_tot}">总计</td><td>${g_spend:,.2f}</td><td>{g_tot_shows}</td><td>{g_new_shows}</td><td>{g_new_passed}</td><td>{grand_n_rate:.2f}%</td><td{grand_a_class}>{grand_a_rate:.2f}%</td></tr></tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_table2_4level_pivot(show_level)

            # 📊 3. 本周新剧数据汇总对比
            st.markdown('<div class="report-subtitle">📊 3. 本周新剧数据汇总对比（新剧定义：首次投放日期属于近 7 天本周内）</div>', unsafe_allow_html=True)

            def render_new_old_summary(show_level):
                if show_level.empty: return
                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr><th style="width: 15%;">是否新剧</th><th style="width: 20%;">花费 ($)</th><th style="width: 15%;">D0 达标率 (%)</th><th style="width: 15%;">CPI ($)</th><th style="width: 15%;">CPP ($)</th><th style="width: 10%;">注册拉新占比 (%)</th><th style="width: 10%;">拉新付费占比 (%)</th></tr></thead><tbody>']
                
                def calc_metrics(group):
                    sp, rev, tar, act_num, pay_num, n_reg, n_pay = group['spend'].sum(), group['d0_rev'].sum(), group['d0_target'].sum(), group['act'].sum(), group['pay'].sum(), group['new_reg'].sum(), group['new_pay'].sum()
                    return sp, (rev / tar * 100) if tar > 0 else 0.0, (sp / act_num) if act_num > 0 else 0.0, (sp / pay_num) if pay_num > 0 else 0.0, (n_reg / act_num * 100) if act_num > 0 else 0.0, (n_pay / pay_num * 100) if pay_num > 0 else 0.0

                sp_old, d0_old, cpi_old, cpp_old, reg_old, pay_old = calc_metrics(show_level[~show_level['is_new_drama']])
                sp_new, d0_new, cpi_new, cpp_new, reg_new, pay_new = calc_metrics(show_level[show_level['is_new_drama']])
                rows_data = sorted([('否', sp_old, d0_old, cpi_old, cpp_old, reg_old, pay_old), ('是', sp_new, d0_new, cpi_new, cpp_new, reg_new, pay_new)], key=lambda x: x[1], reverse=True)

                for is_n, sp_v, d0_v, cpi_v, cpp_v, reg_v, pay_v in rows_data:
                    passed_class = ' class="cell-passed"' if d0_v >= 85.0 else (' class="cell-passed-text-only"' if d0_v >= 80.0 else '')
                    html_list.append(f'<tr><td>{is_n}</td><td>${sp_v:,.2f}</td><td{passed_class}>{d0_v:.2f}%</td><td>${cpi_v:.2f}</td><td>${cpp_v:.2f}</td><td>{reg_v:.2f}%</td><td>{pay_v:.2f}%</td></tr>')

                sp_tot, d0_tot, cpi_tot, cpp_tot, reg_tot, pay_tot = calc_metrics(show_level)
                tot_passed_class = ' class="cell-passed-underline"' if d0_tot >= 85.0 else (' class="cell-passed-text-only"' if d0_tot >= 80.0 else '')
                html_list.append(f'<tr class="row-grand-total"><td>总计</td><td>${sp_tot:,.2f}</td><td{tot_passed_class}>{d0_tot:.2f}%</td><td>${cpi_tot:.2f}</td><td>${cpp_tot:.2f}</td><td>{reg_tot:.2f}%</td><td>{pay_tot:.2f}%</td></tr></tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_new_old_summary(show_level)

            # 📊 4. 本周每部剧集数据明细透视
            st.markdown('<div class="report-subtitle">📊 4. 本周每部剧集数据明细透视</div>', unsafe_allow_html=True)

            def render_table3_pivot(show_level):
                if show_level.empty: return
                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col: html_list.append('<th>语言</th>')
                html_list.append('<th>序号</th><th>剧目名称</th><th>中文推广剧目/书籍名称</th><th>英文推广剧目/书籍名称</th><th>版权来源</th><th>漫剧类型</th><th>漫剧分类</th><th>本周总花费 ($)</th><th>D0 达标率 (%)</th><th>D0 ROAS (%)</th><th>CPI ($)</th><th>付费率 (%)</th><th>CPP ($)</th><th>新注册占比 (%)</th><th>付费新注册占比 (%)</th></tr></thead><tbody>')

                g_spend, g_d0_rev, g_d0_tar, g_act, g_pay, g_new_reg, g_new_pay = 0, 0, 0, 0, 0, 0, 0
                lang_grouped = sorted(show_level.groupby(col_lang), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values(by='spend', ascending=False)
                    lang_spend, lang_d0_rev, lang_d0_tar, lang_act, lang_pay, lang_new_reg, lang_new_pay = group_sorted['spend'].sum(), group_sorted['d0_rev'].sum(), group_sorted['d0_target'].sum(), group_sorted['act'].sum(), group_sorted['pay'].sum(), group_sorted['new_reg'].sum(), group_sorted['new_pay'].sum()
                    g_spend += lang_spend; g_d0_rev += lang_d0_rev; g_d0_tar += lang_d0_tar; g_act += lang_act; g_pay += lang_pay; g_new_reg += lang_new_reg; g_new_pay += lang_new_pay
                    rowspan_count = len(group_sorted); first_row = True; seq_num = 1

                    for _, row in group_sorted.iterrows():
                        sp, d0_r, rev, act_num, pay_num, n_reg, n_pay = row['spend'], row['d0_rate'], row['d0_rev'], row['act'], row['pay'], row['new_reg'], row['new_pay']
                        d0_pct = d0_r * 100
                        passed_cell_class = ' class="cell-passed"' if d0_pct >= 85.0 else (' class="cell-passed-text-only"' if d0_pct >= 80.0 else '')
                        html_list.append('<tr>')
                        if show_lang_col and first_row: html_list.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>'); first_row = False
                        html_list.append(f'<td>{seq_num}</td><td>{row[col_title]}</td><td>{row["cn_title"]}</td><td>{row["en_title"]}</td><td>{row[col_copyright]}</td><td>{row[col_manga_type]}</td><td>{row[col_manga_cls]}</td><td style="font-weight:600;">${sp:,.2f}</td><td{passed_cell_class}>{d0_pct:.2f}%</td><td>{(rev/sp*100) if sp>0 else 0:.2f}%</td><td>${(sp/act_num) if act_num>0 else 0:.2f}</td><td>{(pay_num/act_num*100) if act_num>0 else 0:.2f}%</td><td>${(sp/pay_num) if pay_num>0 else 0:.2f}</td><td>{(n_reg/act_num*100) if act_num>0 else 0:.2f}%</td><td>{(n_pay/pay_num*100) if pay_num>0 else 0:.2f}%</td></tr>')
                        seq_num += 1

                    lang_d0_rate = (lang_d0_rev / lang_d0_tar * 100) if lang_d0_tar > 0 else 0.0
                    lang_d0_class = ' class="cell-passed-underline"' if lang_d0_rate >= 85.0 else (' class="cell-passed-text-only"' if lang_d0_rate >= 80.0 else '')
                    span_val_lang = 8 if show_lang_col else 7
                    html_list.append(f'<tr class="row-lang-subtotal"><td colspan="{span_val_lang}">{lang_val} 小计</td><td>${lang_spend:,.2f}</td><td{lang_d0_class}>{lang_d0_rate:.2f}%</td><td>{(lang_d0_rev/lang_spend*100) if lang_spend>0 else 0:.2f}%</td><td>${(lang_spend/lang_act) if lang_act>0 else 0:.2f}</td><td>{(lang_pay/lang_act*100) if lang_act>0 else 0:.2f}%</td><td>${(lang_spend/lang_pay) if lang_pay>0 else 0:.2f}</td><td>{(lang_new_reg/lang_act*100) if lang_act>0 else 0:.2f}%</td><td>{(lang_new_pay/lang_pay*100) if lang_pay>0 else 0:.2f}%</td></tr>')

                g_d0_rate = (g_d0_rev / g_d0_tar * 100) if g_d0_tar > 0 else 0.0
                g_d0_class = ' class="cell-passed-underline"' if g_d0_rate >= 85.0 else (' class="cell-passed-text-only"' if g_d0_rate >= 80.0 else '')
                span_val_tot = 8 if show_lang_col else 7
                html_list.append(f'<tr class="row-grand-total"><td colspan="{span_val_tot}">总计</td><td>${g_spend:,.2f}</td><td{g_d0_class}>{g_d0_rate:.2f}%</td><td>{(g_d0_rev/g_spend*100) if g_spend>0 else 0:.2f}%</td><td>${(g_spend/g_act) if g_act>0 else 0:.2f}</td><td>{(g_pay/g_act*100) if g_act>0 else 0:.2f}%</td><td>${(g_spend/g_pay) if g_pay>0 else 0:.2f}</td><td>{(g_new_reg/g_act*100) if g_act>0 else 0:.2f}%</td><td>{(g_new_pay/g_pay*100) if g_pay>0 else 0:.2f}%</td></tr></tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_table3_pivot(show_level)

            # 📊 5. 本周投放维度拆分
            st.markdown('<div class="report-subtitle">📊 5. 本周投放维度拆分</div>', unsafe_allow_html=True)

            p_w2 = df_w2.copy()
            p_w1 = df_w1.copy()
            p_w2['lang'] = p_w2[col_lang].astype(str).str.strip() if col_lang in p_w2 else "未知"
            p_w1['lang'] = p_w1[col_lang].astype(str).str.strip() if col_lang in p_w1 else "未知"
            p_w2['tz_class'] = p_w2[col_account].apply(get_account_timezone) if col_account in p_w2 and col_account != "不映射/无此列" else '美西账户'

            def process_lang_dim_block(dim_col, dim_name):
                if dim_col == "不映射/无此列" or dim_col not in p_w2.columns: return
                st.markdown(f"**📌 本周{dim_name}指标汇总与消耗占比透视：**")
                g2 = p_w2.groupby(['lang', dim_col], as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum', 'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'})
                tot_spend2_all = g2['spend'].sum()

                html_up = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col: html_up.append('<th>剧目/书籍语言</th>')
                html_up.append(f'<th>{dim_name}</th><th>花费 ($)</th><th>消耗占比 (%)</th><th>D0 达标率 (%)</th><th>ROAS (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比 (%)</th><th>付费拉新占比 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0
                lang_grouped = sorted(g2.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values('spend', ascending=False)
                    rowspan_count = len(group_sorted); first_row = True
                    l_sp, l_rev, l_tar, l_act, l_pay, l_n_reg, l_n_pay = group_sorted['spend'].sum(), group_sorted['d0_rev'].sum(), group_sorted['d0_target'].sum(), group_sorted['act'].sum(), group_sorted['pay'].sum(), group_sorted['new_reg'].sum(), group_sorted['new_pay'].sum()
                    grand_sp += l_sp; grand_rev += l_rev; grand_tar += l_tar; grand_act += l_act; grand_pay += l_pay; grand_n_reg += l_n_reg; grand_n_pay += l_n_pay

                    for _, r in group_sorted.iterrows():
                        sp = r['spend']
                        d0_r = (r['d0_rev'] / r['d0_target'] * 100) if r['d0_target'] > 0 else 0
                        passed_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')
                        html_up.append('<tr>')
                        if show_lang_col and first_row: html_up.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>'); first_row = False
                        html_up.append(f'<td>{r[dim_col]}</td><td>${sp:,.2f}</td><td>{(sp/tot_spend2_all*100) if tot_spend2_all>0 else 0:.2f}%</td><td{passed_class}>{d0_r:.2f}%</td><td>{(r["d0_rev"]/sp*100) if sp>0 else 0:.2f}%</td><td>${(sp/r["act"]) if r["act"]>0 else 0:.2f}</td><td>${(sp/r["pay"]) if r["pay"]>0 else 0:.2f}</td><td>{(r["new_reg"]/r["act"]*100) if r["act"]>0 else 0:.2f}%</td><td>{(r["new_pay"]/r["pay"]*100) if r["pay"]>0 else 0:.2f}%</td></tr>')

                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')
                    span_val_up = 2 if show_lang_col else 1
                    html_up.append(f'<tr class="row-lang-subtotal"><td colspan="{span_val_up}">{lang_val} 小计</td><td>${l_sp:,.2f}</td><td>{(l_sp/tot_spend2_all*100) if tot_spend2_all>0 else 0:.2f}%</td><td{l_passed_class}>{l_d0_r:.2f}%</td><td>{(l_rev/l_sp*100) if l_sp>0 else 0:.2f}%</td><td>${(l_sp/l_act) if l_act>0 else 0:.2f}</td><td>${(l_sp/l_pay) if l_pay>0 else 0:.2f}</td><td>{(l_n_reg/l_act*100) if l_act>0 else 0:.2f}%</td><td>{(l_n_pay/l_pay*100) if l_pay>0 else 0:.2f}%</td></tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')
                span_val_tot = 2 if show_lang_col else 1
                html_up.append(f'<tr class="row-grand-total"><td colspan="{span_val_tot}">总计</td><td>${grand_sp:,.2f}</td><td>100.00%</td><td{g_passed_class}>{g_d0:.2f}%</td><td>{(grand_rev/grand_sp*100) if grand_sp>0 else 0:.2f}%</td><td>${(grand_sp/grand_act) if grand_act>0 else 0:.2f}</td><td>${(grand_sp/grand_pay) if grand_pay>0 else 0:.2f}</td><td>{(grand_n_reg/grand_act*100) if grand_act>0 else 0:.2f}%</td><td>{(grand_n_pay/grand_pay*100) if grand_pay>0 else 0:.2f}%</td></tr></tbody></table></div>')
                st.markdown("".join(html_up), unsafe_allow_html=True)

            st.markdown("**【板块一】渠道维度双周环比对比 (FB / TT)**")
            process_lang_dim_block(col_platform, "投放平台")

            st.markdown("**【板块二】推广目标维度 & 操作系统端成效透视**")
            process_lang_dim_block(col_target, "推广目标")
            process_lang_dim_block(col_os, "操作系统 (广告)")

            st.markdown("**【板块三】Pixel 名称维度数据透视**")
            process_lang_dim_block(col_pixel, "Pixel 名称")

            st.markdown("**【板块四】账户时区维度数据透视**")
            process_lang_dim_block('tz_class', "账户时区分类")

            # =================================================================
            # 三、团队管理
            # =================================================================
            st.markdown('<div class="report-title">三、团队管理</div>', unsafe_allow_html=True)
            OPTIMIZER_WEEKLY_KPI = {'CZM': 63000.0, 'lel': 45360.0, 'LTT': 50400.0, 'Mia': 42000.0, 'Fanny': 42000.0, 'Nora': 30450.0}

            if col_optimizer != "不映射/无此列" and col_optimizer in df_w2.columns:
                w2_opt = df_w2.copy()
                w2_opt['optimizer'] = w2_opt[col_optimizer].astype(str).str.strip()
                g2_opt = w2_opt.groupby('optimizer', as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum', 'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'}).sort_values('spend', ascending=False)

                html_3a = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr><th>优化师</th><th>花费 ($)</th><th>D0 达标率 (%)</th><th>周 KPI (7天) ($)</th><th>完成率 (%)</th><th>D0 ROAS (%)</th><th>付费用户新注册用户占比 (%)</th><th>新注册用户占比 (%)</th></tr></thead><tbody>']
                tot_sp, tot_rev, tot_tar, tot_act, tot_pay, tot_n_reg, tot_n_pay = 0, 0, 0, 0, 0, 0, 0

                for _, r in g2_opt.iterrows():
                    opt = r['optimizer']; sp = r['spend']; rev = r['d0_rev']; tar = r['d0_target']; act_n = r['act']; pay_n = r['pay']
                    tot_sp += sp; tot_rev += rev; tot_tar += tar; tot_act += act_n; tot_pay += pay_n; tot_n_reg += r['new_reg']; tot_n_pay += r['new_pay']
                    d0_r = (rev / tar * 100) if tar > 0 else 0.0
                    kpi_target = OPTIMIZER_WEEKLY_KPI.get(opt, None)
                    kpi_str = f"${kpi_target:,.2f}" if kpi_target else "—"
                    comp_str = f"{(sp / kpi_target * 100):.2f}%" if kpi_target else "—"
                    passed_cls = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')
                    html_3a.append(f'<tr><td>{opt}</td><td>${sp:,.2f}</td><td{passed_cls}>{d0_r:.2f}%</td><td>{kpi_str}</td><td>{comp_str}</td><td>{(rev/sp*100) if sp>0 else 0:.2f}%</td><td>{(r["new_pay"]/pay_n*100) if pay_n>0 else 0:.2f}%</td><td>{(r["new_reg"]/act_n*100) if act_n>0 else 0:.2f}%</td></tr>')

                tot_d0 = (tot_rev / tot_tar * 100) if tot_tar > 0 else 0.0
                tot_passed_cls = ' class="cell-passed-underline"' if tot_d0 >= 85.0 else (' class="cell-passed-text-only"' if tot_d0 >= 80.0 else '')
                html_3a.append(f'<tr class="row-highlight-yellow"><td>总计</td><td>${tot_sp:,.2f}</td><td{tot_passed_cls}>{tot_d0:.2f}%</td><td>—</td><td>—</td><td>{(tot_rev/tot_sp*100) if tot_sp>0 else 0:.2f}%</td><td>{(tot_n_pay/tot_pay*100) if tot_pay>0 else 0:.2f}%</td><td>{(tot_n_reg/tot_act*100) if tot_act>0 else 0:.2f}%</td></tr></tbody></table></div>')
                st.markdown("".join(html_3a), unsafe_allow_html=True)

            # 智能文字复盘卡片
            zi_group = show_level[show_level[col_copyright].str.contains('自制|自办', na=False)]
            cai_group = show_level[show_level[col_copyright].str.contains('采购|外包', na=False)]
            zi_pass = (zi_group['is_new_passed'].sum() / zi_group['is_new_drama'].sum() * 100) if zi_group['is_new_drama'].sum() > 0 else 0
            cai_pass = (cai_group['is_new_passed'].sum() / cai_group['is_new_drama'].sum() * 100) if cai_group['is_new_drama'].sum() > 0 else 0
            waste_sp = show_level[show_level['is_new_drama'] & (~show_level['is_passed'])]['spend'].sum()

            st.markdown(f"""
            <div class="review-card">
                <div class="review-item"><b>1. 自制剧 vs 采购剧过测率对比：</b> 本周自制新剧过测率为 <b>{zi_pass:.1f}%</b>，采购新剧过测率为 <b>{cai_pass:.1f}%</b>。</div>
                <div class="review-item"><b>2. 新剧测试空耗归因：</b> 本周未过测新剧共耗资 <b>${waste_sp:,.2f}</b>，建议优化前3秒留存与开头素材。</div>
                <div class="review-item"><b>3. 在投剧目与团队表现：</b> 头部优化师 KPI 稳步推进，老剧持续提供稳健大盘流水支撑。</div>
            </div>
            """, unsafe_allow_html=True)

    with lang_tabs[0]: render_dashboard_content(target_lang_filter=None)
    with lang_tabs[1]: render_dashboard_content(target_lang_filter="繁体中文")
    with lang_tabs[2]: render_dashboard_content(target_lang_filter="韩语")

else:
    st.info("👆 请在顶部上传【上周】和【本周】两份 Excel 表格。")
