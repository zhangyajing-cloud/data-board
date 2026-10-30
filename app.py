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
        background-color: #212529 !important; /* 深炭灰色表头 */
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
        background-color: #ffe066 !important; /* 第一行/总计暖黄色背景 */
        color: #000000 !important;
        font-weight: bold;
    }
    .row-cp-subtotal {
        background-color: #fff9db !important; /* 版权小计极浅黄色 */
        font-weight: bold;
    }
    .row-lang-subtotal {
        background-color: #fff9c4 !important; /* 语种小计浅黄色 */
        font-weight: bold;
    }
    .row-grand-total {
        background-color: #ffe066 !important; /* 全表总计暖黄色 */
        color: #000000 !important;
        font-weight: bold;
    }
    .cell-passed {
        background-color: #d4edda !important; /* 淡绿背景 */
        color: #28a745 !important; /* 绿色文字 */
        font-weight: bold;
    }
    .cell-passed-text-only {
        color: #28a745 !important; /* 仅仅绿色文字加粗 */
        font-weight: bold;
    }
    .cell-passed-underline {
        color: #28a745 !important; /* 绿色文字加粗 + 下划线 */
        font-weight: bold;
        text-decoration: underline;
    }
    .cell-last-week {
        color: #6c757d !important; /* 浅灰色文字 */
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
@st.cache_data(ttl=3600, show_spinner=False)
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

def find_best_match(df_cols, keywords):
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

# 📌 账户时区分类判断规则函数
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

# 📌 顶部【语种/地区 Tab 选项卡导航组件】
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
    with f4: col_copyright = st.selectbox("版权来源列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['版权来源', '版权', '来源'])))

    f5, f6, f7, f8 = st.columns(4)
    with f5: col_manga_type = st.selectbox("漫剧类型列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['漫剧类型', '类型'])))
    with f6: col_manga_cls = st.selectbox("漫剧分类列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['漫剧分类', '分类'])))
    with f7: col_optimizer = st.selectbox("优化师列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['优化师', '负责人', '优化师姓名'])))
    with f8: col_account = st.selectbox("账户名称/投放账户列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['账户名称', '投放账户', '账户', 'account'])))

    st.markdown("##### 🌐 板块一/二/每日趋势/Pixel & 优化方式/操作系统扩展映射 (严格绑定)")
    e1, e2, e3 = st.columns(3)
    with e1: col_platform = st.selectbox("投放平台列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['投放平台', '平台', 'media_source'])))
    with e2: col_target = st.selectbox("推广目标列 (w2a/app):", all_cols, index=all_cols.index(find_best_match(all_cols, ['推广目标', '目标', 'campaign_type'])))
    with e3: col_pixel = st.selectbox("Pixel 名称列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['pixel名称', 'pixel', '像素'])))

    e4, e5, e6, e7 = st.columns(4)
    with e4: col_opt_type = st.selectbox("优化方式列 (AEO/VO):", all_cols, index=all_cols.index(find_best_match(all_cols, ['优化方式', '优化目标', 'optimization_goal'])))
    with e5: col_bid_val = st.selectbox("手动出价金额/系数列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['手动出价的金额/系数', '手动出价', '出价系数', 'bid'])))
    with e6: col_date = st.selectbox("时间/日期列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['日期', '时间', 'date'])))
    with e7: col_os = st.selectbox("操作系统 (广告)列:", all_cols, index=all_cols.index(find_best_match(all_cols, ['操作系统 (广告)', '操作系统', '系统', 'os', 'platform_os'])))

    st.markdown('</div>', unsafe_allow_html=True)

    # 📌 渲染函数定义：支持基于当前选中的 Tab 语种环境全板动态渲染
    def render_dashboard_content(target_lang_filter=None):
        # 1. 过滤主数据集
        if target_lang_filter and col_lang != "不映射/无此列" and col_lang in raw_df_w2.columns:
            df_w2 = raw_df_w2[raw_df_w2[col_lang].astype(str).str.contains(target_lang_filter, na=False)].copy()
            df_w1 = raw_df_w1[raw_df_w1[col_lang].astype(str).str.contains(target_lang_filter, na=False)].copy() if col_lang in raw_df_w1.columns else raw_df_w1.copy()
        else:
            df_w2 = raw_df_w2.copy()
            df_w1 = raw_df_w1.copy()

        # =================================----------------------------------------
        # 一、整体数据情况
        # =================================----------------------------------------
        st.markdown('<div class="report-title">一、整体数据情况</div>', unsafe_allow_html=True)

        def get_totals(df):
            sp = clean_num(df[col_spend]).sum() if col_spend in df else 0.0
            rev = clean_num(df[col_d0_rev]).sum() if col_d0_rev in df else 0.0
            tar = clean_num(df[col_d0_target]).sum() if col_d0_target in df else 0.0
            act = clean_num(df[col_act]).sum() if col_act in df else 0.0
            pay = clean_num(df[col_pay]).sum() if col_pay in df else 0.0
            
            reg_r = clean_num(df[col_reg_rate]) if col_reg_rate in df else pd.Series(0)
            pay_reg_r = clean_num(df[col_pay_reg_rate]) if col_pay_reg_rate in df else pd.Series(0)
            if reg_r.max() > 1.0: reg_r /= 100.0
            if pay_reg_r.max() > 1.0: pay_reg_r /= 100.0
            
            act_s = clean_num(df[col_act]) if col_act in df else pd.Series(0)
            pay_s = clean_num(df[col_pay]) if col_pay in df else pd.Series(0)
            
            new_reg = (act_s * reg_r).sum()
            new_pay = (pay_s * pay_reg_r).sum()
            
            return {
                'spend': sp, 
                'd0_rate': rev / tar if tar > 0 else 0.0, 
                'roas': rev / sp if sp > 0 else 0.0, 
                'pay_reg_pct': new_pay / pay if pay > 0 else 0.0, 
                'reg_pct': new_reg / act if act > 0 else 0.0
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

        def process_lang_breakdown(df_w1, df_w2):
            if col_lang == "不映射/无此列" or col_lang not in df_w2.columns: return
            df2 = df_w2.copy()
            df2['spend'] = clean_num(df2[col_spend]) if col_spend in df2 else 0
            df2['d0_rev'] = clean_num(df2[col_d0_rev]) if col_d0_rev in df2 else 0
            df2['d0_target'] = clean_num(df2[col_d0_target]) if col_d0_target in df2 else 0
            df2['act'] = clean_num(df2[col_act]) if col_act in df2 else 0
            df2['pay'] = clean_num(df2[col_pay]) if col_pay in df2 else 0
            
            reg_r2 = clean_num(df2[col_reg_rate]) if col_reg_rate in df2 else pd.Series(0)
            pay_reg_r2 = clean_num(df2[col_pay_reg_rate]) if col_pay_reg_rate in df2 else pd.Series(0)
            if reg_r2.max() > 1.0: reg_r2 /= 100.0
            if pay_reg_r2.max() > 1.0: pay_reg_r2 /= 100.0
            
            df2['new_reg'] = df2['act'] * reg_r2
            df2['new_pay'] = df2['pay'] * pay_reg_r2
            
            g2 = df2.groupby(col_lang, as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum', 'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'}).sort_values('spend', ascending=False)
            
            df1 = df_w1.copy()
            df1['spend'] = clean_num(df1[col_spend]) if col_spend in df1 else 0
            df1['d0_rev'] = clean_num(df1[col_d0_rev]) if col_d0_rev in df1 else 0
            df1['d0_target'] = clean_num(df1[col_d0_target]) if col_d0_target in df1 else 0
            g1 = df1.groupby(col_lang, as_index=False).agg({'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum'})
            
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

        process_lang_breakdown(df_w1, df_w2)

        # =================================----------------------------------------
        # 二、具体投放数据和剧集情况复盘
        # =================================----------------------------------------
        st.markdown('<div class="report-title">二、具体投放数据和剧集情况复盘</div>', unsafe_allow_html=True)

        required_m2_cols = [col_lang, col_copyright, col_manga_type, col_manga_cls, col_title, col_spend, col_d0_rev, col_d0_target]
        missing_cols = [c for c in required_m2_cols if c == "不映射/无此列" or c not in df_w2.columns]

        if missing_cols:
            st.warning(f"⚠️ 需要完成追加字段映射，请在上方“第二步”设置：剧目名称、版权来源、漫剧类型、漫剧分类！")
        else:
            w2 = df_w2.copy()
            w1 = df_w1.copy()
            w2[col_title] = w2[col_title].astype(str).str.strip()
            w1[col_title] = w1[col_title].astype(str).str.strip() if col_title in w1 else "-"
            
            w2['cn_title'] = w2[col_cn_title].astype(str).str.strip() if col_cn_title in w2 else "-"
            w2['en_title'] = w2[col_en_title].astype(str).str.strip() if col_en_title in w2 else "-"
            
            w2['spend'] = clean_num(w2[col_spend])
            w2['d0_rev'] = clean_num(w2[col_d0_rev])
            w2['d0_target'] = clean_num(w2[col_d0_target])
            w2['act'] = clean_num(w2[col_act]) if col_act in w2 else 0
            w2['pay'] = clean_num(w2[col_pay]) if col_pay in w2 else 0

            w1['spend'] = clean_num(w1[col_spend]) if col_spend in w1 else 0

            # 📌 剔除零花费
            w2 = w2[w2['spend'] > 0]
            w1 = w1[w1['spend'] > 0]

            # ---------------------------------------------------------------------
            # 🚨【新剧（New Drama）精准判定规则】
            # 从 W1 与 W2 全量底表中提取每个剧目的【历史最早投放日期 Min(日期)】
            # 当且仅当: 首次投放日期 >= 本周起始日期 (week_start) 且 <= 本周截止日期 (week_end) 时判定为本周新剧
            # ---------------------------------------------------------------------
            all_history_df = pd.concat([w1, w2], ignore_index=True)
            if col_date != "不映射/无此列" and col_date in all_history_df.columns:
                all_history_df['date_dt'] = pd.to_datetime(all_history_df[col_date], errors='coerce')
                first_launch_df = all_history_df.dropna(subset=['date_dt']).groupby(col_title)['date_dt'].min().reset_index()
                first_launch_df.rename(columns={'date_dt': '首次投放日期'}, inplace=True)
                
                # 计算本周起止日期
                w2['date_dt'] = pd.to_datetime(w2[col_date], errors='coerce')
                valid_w2_dates = w2['date_dt'].dropna()
                if not valid_w2_dates.empty:
                    week_start = valid_w2_dates.min()
                    week_end = valid_w2_dates.max()
                else:
                    week_start = pd.Timestamp.min
                    week_end = pd.Timestamp.max
                
                w2 = pd.merge(w2, first_launch_df, on=col_title, how='left')
                w2['is_new_drama'] = w2['首次投放日期'].apply(
                    lambda x: True if (pd.notna(x) and week_start <= x <= week_end) else False
                )
            else:
                # 备用兼容逻辑：上周无消耗且本周有消耗
                w1_titles = set(w1[col_title].dropna().unique())
                w2['is_new_drama'] = w2[col_title].apply(lambda x: x not in w1_titles)

            reg_r2 = clean_num(w2[col_reg_rate]) if col_reg_rate in w2 else pd.Series(0)
            pay_reg_r2 = clean_num(w2[col_pay_reg_rate]) if col_pay_reg_rate in w2 else pd.Series(0)
            if reg_r2.max() > 1.0: reg_r2 /= 100.0
            if pay_reg_r2.max() > 1.0: pay_reg_r2 /= 100.0
            w2['new_reg'] = w2['act'] * reg_r2
            w2['new_pay'] = w2['pay'] * pay_reg_r2

            show_level = w2.groupby([col_lang, col_copyright, col_manga_type, col_manga_cls, col_title, 'cn_title', 'en_title'], as_index=False).agg({
                'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum',
                'is_new_drama': 'first'
            })
            show_level = show_level[show_level['spend'] > 0]

            show_level['d0_rate'] = np.where(show_level['d0_target'] > 0, show_level['d0_rev'] / show_level['d0_target'], 0)
            show_level['is_passed'] = show_level['d0_rate'] >= 0.8  # 测出标准 D0 达标率 >= 80%
            show_level['is_new_passed'] = show_level['is_new_drama'] & show_level['is_passed']

            # 📊 1. 本周剧集类型情况及过测情况 - 类型拆分表 (按照最新指令规则重构)
            st.markdown('<div class="report-subtitle">📊 1. 本周剧集类型情况及过测情况 - 类型拆分表</div>', unsafe_allow_html=True)

            def render_table2_4level_pivot(show_level):
                tbl2_raw = []
                for (lang_val, cp_val, mtype_val, mcls_val), group in show_level.groupby([col_lang, col_copyright, col_manga_type, col_manga_cls]):
                    spend_total = group['spend'].sum()
                    total_shows = group[col_title].nunique()
                    
                    # 📌 精确按新规则计算：新剧数量、新剧测出数、在投达标数
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
                html_list.append('<th style="width: 10%;">版权来源</th>')
                html_list.append('<th style="width: 10%;">漫剧类型</th>')
                html_list.append('<th style="width: 10%;">漫剧分类</th>')
                html_list.append('<th style="width: 12%;">本周花费 ($)</th>')
                html_list.append('<th style="width: 8%;">在投总数</th>')
                html_list.append('<th style="width: 8%;">新剧数量</th>')
                html_list.append('<th style="width: 8%;">新剧测出数</th>')
                html_list.append('<th style="width: 10%;">新剧测出率 (%)</th>')
                html_list.append('<th style="width: 12%;">在投剧目达标率 (%)</th>')
                html_list.append('</tr></thead>')
                html_list.append('<tbody>')

                g_spend, g_tot_shows, g_new_shows, g_new_passed, g_act_passed = 0, 0, 0, 0, 0

                lang_grouped = sorted(df_raw.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    lang_spend, lang_tot, lang_new, lang_n_pass, lang_a_pass = 0, 0, 0, 0, 0
                    
                    lang_rows_count = len(group_lang) + group_lang['cp'].nunique()
                    first_lang_row = True

                    cp_grouped = sorted(group_lang.groupby('cp'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                    for cp_val, group_cp in cp_grouped:
                        cp_sorted = group_cp.sort_values(by='spend', ascending=False)
                        
                        cp_spend = cp_sorted['spend'].sum()
                        cp_tot = cp_sorted['total_shows'].sum()
                        cp_new = cp_sorted['new_shows'].sum()
                        cp_n_pass = cp_sorted['new_passed'].sum()
                        cp_a_pass = cp_sorted['active_passed'].sum()

                        lang_spend += cp_spend
                        lang_tot += cp_tot
                        lang_new += cp_new
                        lang_n_pass += cp_n_pass
                        lang_a_pass += cp_a_pass

                        cp_rows_count = len(cp_sorted)
                        first_cp_row = True

                        for _, row in cp_sorted.iterrows():
                            sp = row['spend']
                            n_shows = row['new_shows']
                            n_pass = row['new_passed']
                            a_shows = row['total_shows']
                            a_pass = row['active_passed']

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

                            html_list.append(f'<td>{row["mtype"]}</td>')
                            html_list.append(f'<td>{row["mcls"]}</td>')
                            html_list.append(f'<td style="font-weight:600;">${sp:,.2f}</td>')
                            html_list.append(f'<td>{a_shows}</td>')
                            html_list.append(f'<td>{n_shows}</td>')
                            html_list.append(f'<td>{n_pass}</td>')
                            html_list.append(f'<td>{n_rate:.2f}%</td>')
                            html_list.append(f'<td{a_rate_class}>{a_rate:.2f}%</td>')
                            html_list.append('</tr>')

                        cp_n_rate = (cp_n_pass / cp_new * 100) if cp_new > 0 else 0.0
                        cp_a_rate = (cp_a_pass / cp_tot * 100) if cp_tot > 0 else 0.0
                        cp_a_class = ' class="cell-passed-underline"' if cp_a_rate >= 85.0 else (' class="cell-passed-text-only"' if cp_a_rate >= 80.0 else '')

                        span_val = 3 if show_lang_col else 2
                        html_list.append('<tr class="row-cp-subtotal">')
                        html_list.append(f'<td colspan="{span_val}">{cp_val} 小计</td>')
                        html_list.append(f'<td>${cp_spend:,.2f}</td>')
                        html_list.append(f'<td>{cp_tot}</td>')
                        html_list.append(f'<td>{cp_new}</td>')
                        html_list.append(f'<td>{cp_n_pass}</td>')
                        html_list.append(f'<td>{cp_n_rate:.2f}%</td>')
                        html_list.append(f'<td{cp_a_class}>{cp_a_rate:.2f}%</td>')
                        html_list.append('</tr>')

                    g_spend += lang_spend
                    g_tot_shows += lang_tot
                    g_new_shows += lang_new
                    g_new_passed += lang_n_pass
                    g_act_passed += lang_a_pass

                    lang_n_rate = (lang_n_pass / lang_new * 100) if lang_new > 0 else 0.0
                    lang_a_rate = (lang_a_pass / lang_tot * 100) if lang_tot > 0 else 0.0
                    lang_a_class = ' class="cell-passed-underline"' if lang_a_rate >= 85.0 else (' class="cell-passed-text-only"' if lang_a_rate >= 80.0 else '')

                    span_val_lang = 4 if show_lang_col else 3
                    html_list.append('<tr class="row-lang-subtotal">')
                    html_list.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_list.append(f'<td>${lang_spend:,.2f}</td>')
                    html_list.append(f'<td>{lang_tot}</td>')
                    html_list.append(f'<td>{lang_new}</td>')
                    html_list.append(f'<td>{lang_n_pass}</td>')
                    html_list.append(f'<td>{lang_n_rate:.2f}%</td>')
                    html_list.append(f'<td{lang_a_class}>{lang_a_rate:.2f}%</td>')
                    html_list.append('</tr>')

                grand_n_rate = (g_new_passed / g_new_shows * 100) if g_new_shows > 0 else 0.0
                grand_a_rate = (g_act_passed / g_tot_shows * 100) if g_tot_shows > 0 else 0.0
                grand_a_class = ' class="cell-passed-underline"' if grand_a_rate >= 85.0 else (' class="cell-passed-text-only"' if grand_a_rate >= 80.0 else '')

                span_val_tot = 4 if show_lang_col else 3
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${g_spend:,.2f}</td>')
                html_list.append(f'<td>{g_tot_shows}</td>')
                html_list.append(f'<td>{g_new_shows}</td>')
                html_list.append(f'<td>{g_new_passed}</td>')
                html_list.append(f'<td>{grand_n_rate:.2f}%</td>')
                html_list.append(f'<td{grand_a_class}>{grand_a_rate:.2f}%</td>')
                html_list.append('</tr>')

                html_list.append('</tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_table2_4level_pivot(show_level)

            # 📊 2. 本周新剧数据汇总对比
            st.markdown('<div class="report-subtitle">📊 2. 本周新剧数据汇总对比（新剧定义：首次投放日期属于近 7 天本周内）</div>', unsafe_allow_html=True)

            def render_new_old_summary(show_level):
                if show_level.empty: return

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                html_list.append('<th style="width: 15%;">是否新剧</th>')
                html_list.append('<th style="width: 20%;">花费 ($)</th>')
                html_list.append('<th style="width: 15%;">D0 达标率 (%)</th>')
                html_list.append('<th style="width: 15%;">CPI ($)</th>')
                html_list.append('<th style="width: 15%;">CPP ($)</th>')
                html_list.append('<th style="width: 10%;">注册拉新占比 (%)</th>')
                html_list.append('<th style="width: 10%;">拉新付费占比 (%)</th>')
                html_list.append('</tr></thead>')
                html_list.append('<tbody>')

                old_shows = show_level[~show_level['is_new_drama']]
                new_shows = show_level[show_level['is_new_drama']]

                def calc_metrics(group):
                    sp = group['spend'].sum()
                    rev = group['d0_rev'].sum()
                    tar = group['d0_target'].sum()
                    act_num = group['act'].sum()
                    pay_num = group['pay'].sum()
                    n_reg = group['new_reg'].sum()
                    n_pay = group['new_pay'].sum()

                    d0_rate = (rev / tar * 100) if tar > 0 else 0.0
                    cpi = (sp / act_num) if act_num > 0 else 0.0
                    cpp = (sp / pay_num) if pay_num > 0 else 0.0
                    reg_pct = (n_reg / act_num * 100) if act_num > 0 else 0.0
                    pay_pct = (n_pay / pay_num * 100) if pay_num > 0 else 0.0

                    return sp, d0_rate, cpi, cpp, reg_pct, pay_pct

                sp_old, d0_old, cpi_old, cpp_old, reg_old, pay_old = calc_metrics(old_shows)
                sp_new, d0_new, cpi_new, cpp_new, reg_new, pay_new = calc_metrics(new_shows)

                rows_data = [
                    ('否', sp_old, d0_old, cpi_old, cpp_old, reg_old, pay_old),
                    ('是', sp_new, d0_new, cpi_new, cpp_new, reg_new, pay_new)
                ]
                rows_data.sort(key=lambda x: x[1], reverse=True)

                for is_n, sp_v, d0_v, cpi_v, cpp_v, reg_v, pay_v in rows_data:
                    passed_class = ' class="cell-passed"' if d0_v >= 85.0 else (' class="cell-passed-text-only"' if d0_v >= 80.0 else '')
                    html_list.append(f'<tr><td>{is_n}</td>')
                    html_list.append(f'<td>${sp_v:,.2f}</td>')
                    html_list.append(f'<td{passed_class}>{d0_v:.2f}%</td>')
                    html_list.append(f'<td>${cpi_v:.2f}</td>')
                    html_list.append(f'<td>${cpp_v:.2f}</td>')
                    html_list.append(f'<td>{reg_v:.2f}%</td>')
                    html_list.append(f'<td>{pay_v:.2f}%</td></tr>')

                sp_tot, d0_tot, cpi_tot, cpp_tot, reg_tot, pay_tot = calc_metrics(show_level)
                tot_passed_class = ' class="cell-passed-underline"' if d0_tot >= 85.0 else (' class="cell-passed-text-only"' if d0_tot >= 80.0 else '')
                html_list.append('<tr class="row-grand-total"><td>总计</td>')
                html_list.append(f'<td>${sp_tot:,.2f}</td>')
                html_list.append(f'<td{tot_passed_class}>{d0_tot:.2f}%</td>')
                html_list.append(f'<td>${cpi_tot:.2f}</td>')
                html_list.append(f'<td>${cpp_tot:.2f}</td>')
                html_list.append(f'<td>{reg_tot:.2f}%</td>')
                html_list.append(f'<td>{pay_tot:.2f}%</td></tr>')

                html_list.append('</tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_new_old_summary(show_level)

            # 📊 3. 本周每部剧集数据明细透视
            st.markdown('<div class="report-subtitle">📊 3. 本周每部剧集数据明细透视</div>', unsafe_allow_html=True)

            def render_table3_pivot(show_level):
                if show_level.empty: return

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th>语言</th>')
                html_list.append('<th>序号</th>')
                html_list.append('<th>剧目名称</th>')
                html_list.append('<th>中文推广剧目/书籍名称</th>')
                html_list.append('<th>英文推广剧目/书籍名称</th>')
                html_list.append('<th>版权来源</th>')
                html_list.append('<th>漫剧类型</th>')
                html_list.append('<th>漫剧分类</th>')
                html_list.append('<th>本周总花费 ($)</th>')
                html_list.append('<th>D0 达标率 (%)</th>')
                html_list.append('<th>D0 ROAS (%)</th>')
                html_list.append('<th>CPI ($)</th>')
                html_list.append('<th>付费率 (%)</th>')
                html_list.append('<th>CPP ($)</th>')
                html_list.append('<th>新注册占比 (%)</th>')
                html_list.append('<th>付费新注册占比 (%)</th>')
                html_list.append('</tr></thead>')
                html_list.append('<tbody>')

                g_spend, g_d0_rev, g_d0_tar, g_act, g_pay, g_new_reg, g_new_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(show_level.groupby(col_lang), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values(by='spend', ascending=False)
                    
                    lang_spend = group_sorted['spend'].sum()
                    lang_d0_rev = group_sorted['d0_rev'].sum()
                    lang_d0_tar = group_sorted['d0_target'].sum()
                    lang_act = group_sorted['act'].sum()
                    lang_pay = group_sorted['pay'].sum()
                    lang_new_reg = group_sorted['new_reg'].sum()
                    lang_new_pay = group_sorted['new_pay'].sum()

                    g_spend += lang_spend
                    g_d0_rev += lang_d0_rev
                    g_d0_tar += lang_d0_tar
                    g_act += lang_act
                    g_pay += lang_pay
                    g_new_reg += lang_new_reg
                    g_new_pay += lang_new_pay

                    rowspan_count = len(group_sorted)
                    first_row = True
                    seq_num = 1

                    for _, row in group_sorted.iterrows():
                        sp = row['spend']
                        d0_r = row['d0_rate']
                        rev = row['d0_rev']
                        act_num = row['act']
                        pay_num = row['pay']
                        n_reg = row['new_reg']
                        n_pay = row['new_pay']

                        roas_val = (rev / sp * 100) if sp > 0 else 0.0
                        cpi_val = (sp / act_num) if act_num > 0 else 0.0
                        pay_rate_val = (pay_num / act_num * 100) if act_num > 0 else 0.0
                        cpp_val = (sp / pay_num) if pay_num > 0 else 0.0
                        reg_rate_val = (n_reg / act_num * 100) if act_num > 0 else 0.0
                        pay_reg_rate_val = (n_pay / pay_num * 100) if pay_num > 0 else 0.0

                        d0_pct = d0_r * 100
                        passed_cell_class = ' class="cell-passed"' if d0_pct >= 85.0 else (' class="cell-passed-text-only"' if d0_pct >= 80.0 else '')

                        html_list.append('<tr>')
                        if show_lang_col and first_row:
                            html_list.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff; color:#212529;">{lang_val}</td>')
                            first_row = False

                        html_list.append(f'<td>{seq_num}</td>')
                        html_list.append(f'<td>{row[col_title]}</td>')
                        html_list.append(f'<td>{row["cn_title"]}</td>')
                        html_list.append(f'<td>{row["en_title"]}</td>')
                        html_list.append(f'<td>{row[col_copyright]}</td>')
                        html_list.append(f'<td>{row[col_manga_type]}</td>')
                        html_list.append(f'<td>{row[col_manga_cls]}</td>')
                        html_list.append(f'<td style="font-weight:600;">${sp:,.2f}</td>')
                        html_list.append(f'<td{passed_cell_class}>{d0_pct:.2f}%</td>')
                        html_list.append(f'<td>{roas_val:.2f}%</td>')
                        html_list.append(f'<td>${cpi_val:.2f}</td>')
                        html_list.append(f'<td>{pay_rate_val:.2f}%</td>')
                        html_list.append(f'<td>${cpp_val:.2f}</td>')
                        html_list.append(f'<td>{reg_rate_val:.2f}%</td>')
                        html_list.append(f'<td>{pay_reg_rate_val:.2f}%</td>')
                        html_list.append('</tr>')
                        seq_num += 1

                    lang_d0_rate = (lang_d0_rev / lang_d0_tar * 100) if lang_d0_tar > 0 else 0.0
                    lang_roas = (lang_d0_rev / lang_spend * 100) if lang_spend > 0 else 0.0
                    lang_cpi = (lang_spend / lang_act) if lang_act > 0 else 0.0
                    lang_pay_rate = (lang_pay / lang_act * 100) if lang_act > 0 else 0.0
                    lang_cpp = (lang_spend / lang_pay) if lang_pay > 0 else 0.0
                    lang_reg_rate = (lang_new_reg / lang_act * 100) if lang_act > 0 else 0.0
                    lang_pay_reg_rate = (lang_new_pay / lang_pay * 100) if lang_pay > 0 else 0.0
                    
                    lang_d0_class = ' class="cell-passed-underline"' if lang_d0_rate >= 85.0 else (' class="cell-passed-text-only"' if lang_d0_rate >= 80.0 else '')

                    span_val_lang = 8 if show_lang_col else 7
                    html_list.append('<tr class="row-lang-subtotal">')
                    html_list.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_list.append(f'<td>${lang_spend:,.2f}</td>')
                    html_list.append(f'<td{lang_d0_class}>{lang_d0_rate:.2f}%</td>')
                    html_list.append(f'<td>{lang_roas:.2f}%</td>')
                    html_list.append(f'<td>${lang_cpi:.2f}</td>')
                    html_list.append(f'<td>{lang_pay_rate:.2f}%</td>')
                    html_list.append(f'<td>${lang_cpp:.2f}</td>')
                    html_list.append(f'<td>{lang_reg_rate:.2f}%</td>')
                    html_list.append(f'<td>{lang_pay_reg_rate:.2f}%</td>')
                    html_list.append('</tr>')

                g_d0_rate = (g_d0_rev / g_d0_tar * 100) if g_d0_tar > 0 else 0.0
                g_roas = (g_d0_rev / g_spend * 100) if g_spend > 0 else 0.0
                g_cpi = (g_spend / g_act) if g_act > 0 else 0.0
                g_pay_rate = (g_pay / g_act * 100) if g_act > 0 else 0.0
                g_cpp = (g_spend / g_pay) if g_pay > 0 else 0.0
                g_reg_rate = (g_new_reg / g_act * 100) if g_act > 0 else 0.0
                g_pay_reg_rate = (g_new_pay / g_pay * 100) if g_pay > 0 else 0.0

                g_d0_class = ' class="cell-passed-underline"' if g_d0_rate >= 85.0 else (' class="cell-passed-text-only"' if g_d0_rate >= 80.0 else '')

                span_val_tot = 8 if show_lang_col else 7
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${g_spend:,.2f}</td>')
                html_list.append(f'<td{g_d0_class}>{g_d0_rate:.2f}%</td>')
                html_list.append(f'<td>{g_roas:.2f}%</td>')
                html_list.append(f'<td>${g_cpi:.2f}</td>')
                html_list.append(f'<td>{g_pay_rate:.2f}%</td>')
                html_list.append(f'<td>${g_cpp:.2f}</td>')
                html_list.append(f'<td>{g_reg_rate:.2f}%</td>')
                html_list.append(f'<td>{g_pay_reg_rate:.2f}%</td>')
                html_list.append('</tr>')

                html_list.append('</tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_table3_pivot(show_level)

            # ---------------------------------------------------------------------
            # 🚨 板块 4：本周投放维度拆分
            # ---------------------------------------------------------------------
            st.markdown('<div class="report-subtitle">📊 4. 本周投放维度拆分</div>', unsafe_allow_html=True)

            def prepare_df_with_metrics(df_raw):
                df = df_raw.copy()
                df['lang'] = df[col_lang].astype(str).str.strip() if col_lang in df else "未知"
                df['spend'] = clean_num(df[col_spend]) if col_spend in df else 0.0
                df['d0_rev'] = clean_num(df[col_d0_rev]) if col_d0_rev in df else 0.0
                df['d0_target'] = clean_num(df[col_d0_target]) if col_d0_target in df else 0.0
                df['act'] = clean_num(df[col_act]) if col_act in df else 0.0
                df['pay'] = clean_num(df[col_pay]) if col_pay in df else 0.0
                
                # 📌 剔除零花费
                df = df[df['spend'] > 0]
                
                if col_account in df and col_account != "不映射/无此列":
                    df['tz_class'] = df[col_account].apply(get_account_timezone)
                else:
                    df['tz_class'] = '美西账户'

                reg_r = clean_num(df[col_reg_rate]) if col_reg_rate in df else pd.Series(0)
                pay_reg_r = clean_num(df[col_pay_reg_rate]) if col_pay_reg_rate in df else pd.Series(0)
                if reg_r.max() > 1.0: reg_r /= 100.0
                if pay_reg_r.max() > 1.0: pay_reg_r /= 100.0
                
                df['new_reg'] = df['act'] * reg_r
                df['new_pay'] = df['pay'] * pay_reg_r
                return df

            p_w2 = prepare_df_with_metrics(df_w2)
            p_w1 = prepare_df_with_metrics(df_w1)

            def process_lang_dim_block(dim_col, dim_name):
                if dim_col == "不映射/无此列" or dim_col not in p_w2.columns:
                    st.warning(f"⚠️ 未映射 {dim_name} 列，请在上方“第二步”设置！")
                    return

                st.markdown(f"**📌 本周{dim_name}指标汇总与消耗占比透视：**")

                g2 = p_w2.groupby(['lang', dim_col], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                tot_spend2_all = g2['spend'].sum()

                html_up = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_up.append('<th>剧目/书籍语言</th>')
                html_up.append(f'<th>{dim_name}</th><th>花费 ($)</th><th>消耗占比 (%)</th><th>D0 达标率 (%)</th><th>ROAS (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比 (%)</th><th>付费拉新占比 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(g2.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values('spend', ascending=False)
                    rowspan_count = len(group_sorted)
                    first_row = True

                    l_sp = group_sorted['spend'].sum()
                    l_rev = group_sorted['d0_rev'].sum()
                    l_tar = group_sorted['d0_target'].sum()
                    l_act = group_sorted['act'].sum()
                    l_pay = group_sorted['pay'].sum()
                    l_n_reg = group_sorted['new_reg'].sum()
                    l_n_pay = group_sorted['new_pay'].sum()

                    grand_sp += l_sp
                    grand_rev += l_rev
                    grand_tar += l_tar
                    grand_act += l_act
                    grand_pay += l_pay
                    grand_n_reg += l_n_reg
                    grand_n_pay += l_n_pay

                    for _, r in group_sorted.iterrows():
                        sp = r['spend']
                        p_pct = (sp / tot_spend2_all * 100) if tot_spend2_all > 0 else 0
                        d0_r = (r['d0_rev'] / r['d0_target'] * 100) if r['d0_target'] > 0 else 0
                        roas_v = (r['d0_rev'] / sp * 100) if sp > 0 else 0
                        cpi_v = (sp / r['act']) if r['act'] > 0 else 0
                        cpp_v = (sp / r['pay']) if r['pay'] > 0 else 0
                        reg_pct = (r['new_reg'] / r['act'] * 100) if r['act'] > 0 else 0
                        pay_pct = (r['new_pay'] / r['pay'] * 100) if r['pay'] > 0 else 0

                        passed_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                        html_up.append('<tr>')
                        if show_lang_col and first_row:
                            html_up.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_row = False

                        html_up.append(f'<td>{r[dim_col]}</td>')
                        html_up.append(f'<td>${sp:,.2f}</td>')
                        html_up.append(f'<td>{p_pct:.2f}%</td>')
                        html_up.append(f'<td{passed_class}>{d0_r:.2f}%</td>')
                        html_up.append(f'<td>{roas_v:.2f}%</td>')
                        html_up.append(f'<td>${cpi_v:.2f}</td>')
                        html_up.append(f'<td>${cpp_v:.2f}</td>')
                        html_up.append(f'<td>{reg_pct:.2f}%</td>')
                        html_up.append(f'<td>{pay_pct:.2f}%</td>')
                        html_up.append('</tr>')

                    l_p_pct = (l_sp / tot_spend2_all * 100) if tot_spend2_all > 0 else 0
                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0
                    l_roas = (l_rev / l_sp * 100) if l_sp > 0 else 0
                    l_cpi = (l_sp / l_act) if l_act > 0 else 0
                    l_cpp = (l_sp / l_pay) if l_pay > 0 else 0
                    l_reg_pct = (l_n_reg / l_act * 100) if l_act > 0 else 0
                    l_pay_pct = (l_n_pay / l_pay * 100) if l_pay > 0 else 0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')

                    span_val_up = 2 if show_lang_col else 1
                    html_up.append('<tr class="row-lang-subtotal">')
                    html_up.append(f'<td colspan="{span_val_up}">{lang_val} 小计</td>')
                    html_up.append(f'<td>${l_sp:,.2f}</td>')
                    html_up.append(f'<td>{l_p_pct:.2f}%</td>')
                    html_up.append(f'<td{l_passed_class}>{l_d0_r:.2f}%</td>')
                    html_up.append(f'<td>{l_roas:.2f}%</td>')
                    html_up.append(f'<td>${l_cpi:.2f}</td>')
                    html_up.append(f'<td>${l_cpp:.2f}</td>')
                    html_up.append(f'<td>{l_reg_pct:.2f}%</td>')
                    html_up.append(f'<td>{l_pay_pct:.2f}%</td>')
                    html_up.append('</tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0
                g_roas = (grand_rev / grand_sp * 100) if grand_sp > 0 else 0
                g_cpi = (grand_sp / grand_act) if grand_act > 0 else 0
                g_cpp = (grand_sp / grand_pay) if grand_pay > 0 else 0
                g_reg_p = (grand_n_reg / grand_act * 100) if grand_act > 0 else 0
                g_pay_p = (grand_n_pay / grand_pay * 100) if grand_pay > 0 else 0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')

                span_val_tot = 2 if show_lang_col else 1
                html_up.append('<tr class="row-grand-total">')
                html_up.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_up.append(f'<td>${grand_sp:,.2f}</td>')
                html_up.append('<td>100.00%</td>')
                html_up.append(f'<td{g_passed_class}>{g_d0:.2f}%</td>')
                html_up.append(f'<td>{g_roas:.2f}%</td>')
                html_up.append(f'<td>${g_cpi:.2f}</td>')
                html_up.append(f'<td>${g_cpp:.2f}</td>')
                html_up.append(f'<td>{g_reg_p:.2f}%</td>')
                html_up.append(f'<td>{g_pay_p:.2f}%</td>')
                html_up.append('</tr></tbody></table></div>')

                st.markdown("".join(html_up), unsafe_allow_html=True)

                st.markdown(f"**📌 {dim_name}双周环比变动明细：**")

                g1 = p_w1.groupby(['lang', dim_col], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                html_down = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                if show_lang_col:
                    html_down.append('<th>剧目/书籍语言</th>')
                html_down.append(f'<th>{dim_name}</th><th>本周消耗 ($)</th><th>上周消耗 ($)</th><th>消耗环比 (%)</th><th>本周 D0 达标率 (%)</th><th>达标率环比 (%)</th><th>本周 D0 ROAS (%)</th><th>ROAS 环比 (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比 (%)</th><th>付费拉新占比 (%)</th></tr></thead><tbody>')

                for lang_val, group_lang2 in lang_grouped:
                    group_lang2_sorted = group_lang2.sort_values('spend', ascending=False)
                    lang_rows_count_down = len(group_lang2_sorted)
                    first_lang_row_down = True

                    for _, r2 in group_lang2_sorted.iterrows():
                        d_key = r2[dim_col]
                        match_1 = g1[(g1['lang'] == lang_val) & (g1[dim_col] == d_key)]

                        s2 = r2['spend']
                        s1 = match_1['spend'].sum() if not match_1.empty else 0.0

                        d0_2 = (r2['d0_rev'] / r2['d0_target'] * 100) if r2['d0_target'] > 0 else 0.0
                        d0_1 = (match_1['d0_rev'].sum() / match_1['d0_target'].sum() * 100) if not match_1.empty and match_1['d0_target'].sum() > 0 else 0.0

                        roas2 = (r2['d0_rev'] / s2 * 100) if s2 > 0 else 0.0
                        roas1 = (match_1['d0_rev'].sum() / s1 * 100) if not match_1.empty and s1 > 0 else 0.0

                        cpi2 = (s2 / r2['act']) if r2['act'] > 0 else 0.0
                        cpp2 = (s2 / r2['pay']) if r2['pay'] > 0 else 0.0

                        reg2 = (r2['new_reg'] / r2['act'] * 100) if r2['act'] > 0 else 0.0
                        pay2 = (r2['new_pay'] / r2['pay'] * 100) if r2['pay'] > 0 else 0.0

                        s_diff_pct = ((s2 - s1) / s1 * 100) if s1 > 0 else 0.0
                        d0_diff_pp = d0_2 - d0_1
                        roas_diff_pp = roas2 - roas1

                        s_diff_str, s_style = format_diff_with_arrow(s_diff_pct)
                        d0_diff_str, d0_style = format_diff_with_arrow(d0_diff_pp, is_pp=True)
                        roas_diff_str, roas_style = format_diff_with_arrow(roas_diff_pp, is_pp=True)

                        passed_class = ' class="cell-passed"' if d0_2 >= 85.0 else (' class="cell-passed-text-only"' if d0_2 >= 80.0 else '')

                        html_down.append('<tr>')
                        if show_lang_col and first_lang_row_down:
                            html_down.append(f'<td rowspan="{lang_rows_count_down}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_lang_row_down = False

                        html_down.append(f'<td style="font-weight:bold;">{d_key}</td>')
                        html_down.append(f'<td>${s2:,.2f}</td>')
                        html_down.append(f'<td class="cell-last-week">${s1:,.2f}</td>')
                        html_down.append(f'<td style="{s_style}">{s_diff_str}</td>')
                        html_down.append(f'<td{passed_class}>{d0_2:.2f}%</td>')
                        html_down.append(f'<td style="{d0_style}">{d0_diff_str}</td>')
                        html_down.append(f'<td>{roas2:.2f}%</td>')
                        html_down.append(f'<td style="{roas_style}">{roas_diff_str}</td>')
                        html_down.append(f'<td>${cpi2:.2f}</td>')
                        html_down.append(f'<td>${cpp2:.2f}</td>')
                        html_down.append(f'<td>{reg2:.2f}%</td>')
                        html_down.append(f'<td>{pay2:.2f}%</td>')
                        html_down.append('</tr>')

                html_down.append('</tbody></table></div>')
                st.markdown("".join(html_down), unsafe_allow_html=True)

            # 【板块一】渠道维度
            st.markdown("**【板块一】渠道维度双周环比对比 (FB / TT)**")
            process_lang_dim_block(col_platform, "投放平台")

            # 【板块二 - 表格 1】推广目标维度双周环比对比 (w2a / app)
            st.markdown("**【板块二 - 表格 1】推广目标维度双周环比对比 (w2a / app)**")
            process_lang_dim_block(col_target, "推广目标")

            # 【板块二 - 表格 2】操作系统端成效汇总与消耗占比透视 (iOS / Android)
            st.markdown("**【板块二 - 表格 2】操作系统端成效汇总与消耗占比透视**")
            def render_os_pivot_by_lang():
                if col_os == "不映射/无此列" or col_os not in p_w2.columns:
                    st.warning("⚠️ 未映射 操作系统 (广告) 列，请在上方“第二步”设置！")
                    return

                g_os = p_w2.groupby(['lang', col_os], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                tot_spend_all = g_os['spend'].sum()

                html_os = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_os.append('<th>剧目/书籍语言</th>')
                html_os.append('<th>操作系统 (广告)</th><th>花费 ($)</th><th>消耗占比 (%)</th><th>D0 达标率 (%)</th><th>ROAS (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比 (%)</th><th>付费拉新占比 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(g_os.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values('spend', ascending=False)
                    rowspan_count = len(group_sorted)
                    first_row = True

                    l_sp = group_sorted['spend'].sum()
                    l_rev = group_sorted['d0_rev'].sum()
                    l_tar = group_sorted['d0_target'].sum()
                    l_act = group_sorted['act'].sum()
                    l_pay = group_sorted['pay'].sum()
                    l_n_reg = group_sorted['new_reg'].sum()
                    l_n_pay = group_sorted['new_pay'].sum()

                    grand_sp += l_sp
                    grand_rev += l_rev
                    grand_tar += l_tar
                    grand_act += l_act
                    grand_pay += l_pay
                    grand_n_reg += l_n_reg
                    grand_n_pay += l_n_pay

                    for _, r in group_sorted.iterrows():
                        sp = r['spend']
                        p_pct = (sp / tot_spend_all * 100) if tot_spend_all > 0 else 0
                        d0_r = (r['d0_rev'] / r['d0_target'] * 100) if r['d0_target'] > 0 else 0
                        roas_v = (r['d0_rev'] / sp * 100) if sp > 0 else 0
                        cpi_v = (sp / r['act']) if r['act'] > 0 else 0
                        cpp_v = (sp / r['pay']) if r['pay'] > 0 else 0
                        reg_pct = (r['new_reg'] / r['act'] * 100) if r['act'] > 0 else 0
                        pay_pct = (r['new_pay'] / r['pay'] * 100) if r['pay'] > 0 else 0

                        passed_cell_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                        html_os.append('<tr>')
                        if show_lang_col and first_row:
                            html_os.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_row = False

                        html_os.append(f'<td>{r[col_os]}</td>')
                        html_os.append(f'<td>${sp:,.2f}</td>')
                        html_os.append(f'<td>{p_pct:.2f}%</td>')
                        html_os.append(f'<td{passed_cell_class}>{d0_r:.2f}%</td>')
                        html_os.append(f'<td>{roas_v:.2f}%</td>')
                        html_os.append(f'<td>${cpi_v:.2f}</td>')
                        html_os.append(f'<td>${cpp_v:.2f}</td>')
                        html_os.append(f'<td>{reg_pct:.2f}%</td>')
                        html_os.append(f'<td>{pay_pct:.2f}%</td>')
                        html_os.append('</tr>')

                    l_p_pct = (l_sp / tot_spend_all * 100) if tot_spend_all > 0 else 0
                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0
                    l_roas = (l_rev / l_sp * 100) if l_sp > 0 else 0
                    l_cpi = (l_sp / l_act) if l_act > 0 else 0
                    l_cpp = (l_sp / l_pay) if l_pay > 0 else 0
                    l_reg_pct = (l_n_reg / l_act * 100) if l_act > 0 else 0
                    l_pay_pct = (l_n_pay / l_pay * 100) if l_pay > 0 else 0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')

                    span_val_lang = 2 if show_lang_col else 1
                    html_os.append('<tr class="row-lang-subtotal">')
                    html_os.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_os.append(f'<td>${l_sp:,.2f}</td>')
                    html_os.append(f'<td>{l_p_pct:.2f}%</td>')
                    html_os.append(f'<td{l_passed_class}>{l_d0_r:.2f}%</td>')
                    html_os.append(f'<td>{l_roas:.2f}%</td>')
                    html_os.append(f'<td>${l_cpi:.2f}</td>')
                    html_os.append(f'<td>${l_cpp:.2f}</td>')
                    html_os.append(f'<td>{l_reg_pct:.2f}%</td>')
                    html_os.append(f'<td>{l_pay_pct:.2f}%</td>')
                    html_os.append('</tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0
                g_roas = (grand_rev / grand_sp * 100) if grand_sp > 0 else 0
                g_cpi = (grand_sp / grand_act) if grand_act > 0 else 0
                g_cpp = (grand_sp / grand_pay) if grand_pay > 0 else 0
                g_reg_p = (grand_n_reg / grand_act * 100) if grand_act > 0 else 0
                g_pay_p = (grand_n_pay / grand_pay * 100) if grand_pay > 0 else 0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')

                span_val_tot = 2 if show_lang_col else 1
                html_os.append('<tr class="row-grand-total">')
                html_os.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_os.append(f'<td>${grand_sp:,.2f}</td>')
                html_os.append('<td>100.00%</td>')
                html_os.append(f'<td{g_passed_class}>{g_d0:.2f}%</td>')
                html_os.append(f'<td>{g_roas:.2f}%</td>')
                html_os.append(f'<td>${g_cpi:.2f}</td>')
                html_os.append(f'<td>${g_cpp:.2f}</td>')
                html_os.append(f'<td>{g_reg_p:.2f}%</td>')
                html_os.append(f'<td>{g_pay_p:.2f}%</td>')
                html_os.append('</tr></tbody></table></div>')

                st.markdown("".join(html_os), unsafe_allow_html=True)

            render_os_pivot_by_lang()

            st.markdown("**【板块三】本周每日数据趋势透视**")

            def render_daily_trend_by_lang(df_w2):
                if col_date == "不映射/无此列" or col_date not in df_w2.columns:
                    st.warning("⚠️ 未映射 时间/日期 列，请在上方“第二步”设置！")
                    return

                df = p_w2.copy()
                try:
                    df['date_clean'] = pd.to_datetime(df[col_date]).dt.strftime('%Y-%m-%d')
                except:
                    df['date_clean'] = df[col_date].astype(str).str.strip()

                g_daily = df.groupby(['lang', 'date_clean'], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th>剧目/书籍语言</th>')
                html_list.append('<th>时间/日期</th><th>花费 ($)</th><th>D0 达标率 (%)</th><th>D0 ROAS (%)</th><th>付费用户新注册用户占比 (%)</th><th>新注册用户占比 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(g_daily.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values('date_clean', ascending=False)
                    rowspan_count = len(group_sorted)
                    first_row = True

                    l_sp = group_sorted['spend'].sum()
                    l_rev = group_sorted['d0_rev'].sum()
                    l_tar = group_sorted['d0_target'].sum()
                    l_act = group_sorted['act'].sum()
                    l_pay = group_sorted['pay'].sum()
                    l_n_reg = group_sorted['new_reg'].sum()
                    l_n_pay = group_sorted['new_pay'].sum()

                    grand_sp += l_sp
                    grand_rev += l_rev
                    grand_tar += l_tar
                    grand_act += l_act
                    grand_pay += l_pay
                    grand_n_reg += l_n_reg
                    grand_n_pay += l_n_pay

                    for _, r in group_sorted.iterrows():
                        sp = r['spend']
                        rev = r['d0_rev']
                        tar = r['d0_target']
                        act_n = r['act']
                        pay_n = r['pay']

                        d0_r = (rev / tar * 100) if tar > 0 else 0.0
                        roas_v = (rev / sp * 100) if sp > 0 else 0.0
                        pay_p = (r['new_pay'] / pay_n * 100) if pay_n > 0 else 0.0
                        reg_p = (r['new_reg'] / act_n * 100) if act_n > 0 else 0.0

                        passed_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                        html_list.append('<tr>')
                        if show_lang_col and first_row:
                            html_list.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_row = False

                        html_list.append(f'<td>{r["date_clean"]}</td>')
                        html_list.append(f'<td>${sp:,.2f}</td>')
                        html_list.append(f'<td{passed_class}>{d0_r:.2f}%</td>')
                        html_list.append(f'<td>{roas_v:.2f}%</td>')
                        html_list.append(f'<td>{pay_p:.2f}%</td>')
                        html_list.append(f'<td>{reg_p:.2f}%</td>')
                        html_list.append('</tr>')

                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0.0
                    l_roas = (l_rev / l_sp * 100) if l_sp > 0 else 0.0
                    l_pay_p = (l_n_pay / l_pay * 100) if l_pay > 0 else 0.0
                    l_reg_p = (l_n_reg / l_act * 100) if l_act > 0 else 0.0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')

                    span_val_lang = 2 if show_lang_col else 1
                    html_list.append('<tr class="row-lang-subtotal">')
                    html_list.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_list.append(f'<td>${l_sp:,.2f}</td>')
                    html_list.append(f'<td{l_passed_class}>{l_d0_r:.2f}%</td>')
                    html_list.append(f'<td>{l_roas:.2f}%</td>')
                    html_list.append(f'<td>{l_pay_p:.2f}%</td>')
                    html_list.append(f'<td>{l_reg_p:.2f}%</td>')
                    html_list.append('</tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0.0
                g_roas = (grand_rev / grand_sp * 100) if grand_sp > 0 else 0.0
                g_pay_p = (grand_n_pay / grand_pay * 100) if grand_pay > 0 else 0.0
                g_reg_p = (grand_n_reg / grand_act * 100) if grand_act > 0 else 0.0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')

                span_val_tot = 2 if show_lang_col else 1
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${grand_sp:,.2f}</td>')
                html_list.append(f'<td{g_passed_class}>{g_d0:.2f}%</td>')
                html_list.append(f'<td>{g_roas:.2f}%</td>')
                html_list.append(f'<td>{g_pay_p:.2f}%</td>')
                html_list.append(f'<td>{g_reg_p:.2f}%</td>')
                html_list.append('</tr></tbody></table></div>')

                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_daily_trend_by_lang(df_w2)

            st.markdown("**【板块四】本周投放平台 & Pixel 名称数据透视（按消耗降序）**")

            def render_pixel_pivot_by_lang(df_w2):
                if col_pixel == "不映射/无此列" or col_pixel not in df_w2.columns:
                    st.warning("⚠️ 未映射 Pixel 名称列，请在上方“第二步”设置！")
                    return

                g = p_w2.groupby(['lang', col_platform, col_pixel], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th>剧目/书籍语言</th>')
                html_list.append('<th>投放平台</th><th>pixel名称</th><th>求和项:花费 ($)</th><th>达标率 (%)</th><th>D0 ROAS (%)</th><th>CPI ($)</th><th>付费率 (%)</th><th>CPP ($)</th><th>新注册用户率 (%)</th><th>求和项:新注册用户付费率 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(g.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    l_sp = group_lang['spend'].sum()
                    l_rev = group_lang['d0_rev'].sum()
                    l_tar = group_lang['d0_target'].sum()
                    l_act = group_lang['act'].sum()
                    l_pay = group_lang['pay'].sum()
                    l_n_reg = group_lang['new_reg'].sum()
                    l_n_pay = group_lang['new_pay'].sum()

                    grand_sp += l_sp
                    grand_rev += l_rev
                    grand_tar += l_tar
                    grand_act += l_act
                    grand_pay += l_pay
                    grand_n_reg += l_n_reg
                    grand_n_pay += l_n_pay

                    p_grouped = sorted(group_lang.groupby(col_platform), key=lambda x: x[1]['spend'].sum(), reverse=True)
                    lang_rows_count = len(group_lang)
                    first_lang_row = True

                    for p_key, group_p in p_grouped:
                        group_sorted = group_p.sort_values('spend', ascending=False)
                        p_rows_count = len(group_sorted)
                        first_p_row = True

                        for _, r in group_sorted.iterrows():
                            sp = r['spend']
                            rev = r['d0_rev']
                            tar = r['d0_target']
                            act_n = r['act']
                            pay_n = r['pay']

                            d0_r = (rev / tar * 100) if tar > 0 else 0
                            roas_v = (rev / sp * 100) if sp > 0 else 0
                            cpi_v = (sp / act_n) if act_n > 0 else 0
                            pay_rate = (pay_n / act_n * 100) if act_n > 0 else 0
                            cpp_v = (sp / pay_n) if pay_n > 0 else 0
                            reg_p = (r['new_reg'] / act_n * 100) if act_n > 0 else 0
                            pay_p = (r['new_pay'] / pay_n * 100) if pay_n > 0 else 0

                            passed_cell_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                            html_list.append('<tr>')
                            if show_lang_col and first_lang_row:
                                html_list.append(f'<td rowspan="{lang_rows_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                                first_lang_row = False

                            if first_p_row:
                                html_list.append(f'<td rowspan="{p_rows_count}" style="font-weight:bold; background-color:#ffffff;">{p_key}</td>')
                                first_p_row = False

                            html_list.append(f'<td>{r[col_pixel]}</td>')
                            html_list.append(f'<td>${sp:,.2f}</td>')
                            html_list.append(f'<td{passed_cell_class}>{d0_r:.2f}%</td>')
                            html_list.append(f'<td>{roas_v:.2f}%</td>')
                            html_list.append(f'<td>${cpi_v:.2f}</td>')
                            html_list.append(f'<td>{pay_rate:.2f}%</td>')
                            html_list.append(f'<td>${cpp_v:.2f}</td>')
                            html_list.append(f'<td>{reg_p:.2f}%</td>')
                            html_list.append(f'<td>{pay_p:.2f}%</td>')
                            html_list.append('</tr>')

                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0.0
                    l_roas = (l_rev / l_sp * 100) if l_sp > 0 else 0.0
                    l_cpi = (l_sp / l_act) if l_act > 0 else 0.0
                    l_pay_rate = (l_pay / l_act * 100) if l_act > 0 else 0.0
                    l_cpp = (l_sp / l_pay) if l_pay > 0 else 0.0
                    l_reg_p = (l_n_reg / l_act * 100) if l_act > 0 else 0.0
                    l_pay_p = (l_n_pay / l_pay * 100) if l_pay > 0 else 0.0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')

                    span_val_lang = 3 if show_lang_col else 2
                    html_list.append('<tr class="row-lang-subtotal">')
                    html_list.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_list.append(f'<td>${l_sp:,.2f}</td>')
                    html_list.append(f'<td{l_passed_class}>{l_d0_r:.2f}%</td>')
                    html_list.append(f'<td>{l_roas:.2f}%</td>')
                    html_list.append(f'<td>${l_cpi:.2f}</td>')
                    html_list.append(f'<td>{l_pay_rate:.2f}%</td>')
                    html_list.append(f'<td>${l_cpp:.2f}</td>')
                    html_list.append(f'<td>{l_reg_p:.2f}%</td>')
                    html_list.append(f'<td>{l_pay_p:.2f}%</td>')
                    html_list.append('</tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0.0
                g_roas = (grand_rev / grand_sp * 100) if grand_sp > 0 else 0.0
                g_cpi = (grand_sp / grand_act) if grand_act > 0 else 0.0
                g_pay_rate = (grand_pay / grand_act * 100) if grand_act > 0 else 0.0
                g_cpp = (grand_sp / grand_pay) if grand_pay > 0 else 0.0
                g_reg_p = (grand_n_reg / grand_act * 100) if grand_act > 0 else 0.0
                g_pay_p = (grand_n_pay / grand_pay * 100) if grand_pay > 0 else 0.0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')

                span_val_tot = 3 if show_lang_col else 2
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${grand_sp:,.2f}</td>')
                html_list.append(f'<td{g_passed_class}>{g_d0:.2f}%</td>')
                html_list.append(f'<td>{g_roas:.2f}%</td>')
                html_list.append(f'<td>${g_cpi:.2f}</td>')
                html_list.append(f'<td>{g_pay_rate:.2f}%</td>')
                html_list.append(f'<td>${g_cpp:.2f}</td>')
                html_list.append(f'<td>{g_reg_p:.2f}%</td>')
                html_list.append(f'<td>{g_pay_p:.2f}%</td>')
                html_list.append('</tr></tbody></table></div>')

                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_pixel_pivot_by_lang(df_w2)

            st.markdown("**【板块五】本周优化方式&出价透视表（最新周 W2 数据透视）**")

            def render_optimization_pivot_5level_by_lang(df_w2):
                if col_opt_type == "不映射/无此列" or col_opt_type not in df_w2.columns:
                    st.warning("⚠️ 未映射 优化方式 列，请在上方“第二步”设置！")
                    return

                df = p_w2.copy()
                df['bid_val'] = df[col_bid_val].astype(str).str.strip() if col_bid_val in df else "-"

                g = df.groupby(['lang', col_platform, col_target, col_opt_type, 'bid_val'], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum'
                })

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th>剧目/书籍语言</th>')
                html_list.append('<th>投放平台</th><th>推广目标</th><th>优化方式</th><th>手动出价的金额/系数</th><th>花费 ($)</th><th>D0 达标率 (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>付费率 (%)</th><th>ARPPU ($)</th><th>ROAS (%)</th></tr></thead><tbody>')

                tot_sp, tot_rev, tot_tar, tot_act, tot_pay = 0, 0, 0, 0, 0

                lang_grouped = sorted(g.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    lang_rows_count = len(group_lang)
                    first_lang_row = True

                    p_grouped = sorted(group_lang.groupby(col_platform), key=lambda x: x[1]['spend'].sum(), reverse=True)

                    for p_key, group_p in p_grouped:
                        p_rows_count = len(group_p)
                        first_p_row = True

                        t_grouped = sorted(group_p.groupby(col_target), key=lambda x: x[1]['spend'].sum(), reverse=True)

                        for t_key, group_t in t_grouped:
                            t_rows_count = len(group_t)
                            first_t_row = True

                            o_grouped = sorted(group_t.groupby(col_opt_type), key=lambda x: x[1]['spend'].sum(), reverse=True)

                            for o_key, group_o in o_grouped:
                                group_sorted = group_o.sort_values('spend', ascending=False)
                                o_rows_count = len(group_sorted)
                                first_o_row = True

                                for _, r in group_sorted.iterrows():
                                    sp = r['spend']
                                    rev = r['d0_rev']
                                    tar = r['d0_target']
                                    act_n = r['act']
                                    pay_n = r['pay']

                                    tot_sp += sp
                                    tot_rev += rev
                                    tot_tar += tar
                                    tot_act += act_n
                                    tot_pay += pay_n

                                    d0_r = (rev / tar * 100) if tar > 0 else 0.0
                                    cpi_v = (sp / act_n) if act_n > 0 else 0.0
                                    cpp_v = (sp / pay_n) if pay_n > 0 else 0.0
                                    pay_rate = (pay_n / act_n * 100) if act_n > 0 else 0.0
                                    arppu_v = (rev / pay_n) if pay_n > 0 else 0.0
                                    roas_v = (rev / sp * 100) if sp > 0 else 0.0

                                    passed_cell_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                                    html_list.append('<tr>')
                                    if show_lang_col and first_lang_row:
                                        html_list.append(f'<td rowspan="{lang_rows_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                                        first_lang_row = False

                                    if first_p_row:
                                        html_list.append(f'<td rowspan="{p_rows_count}" style="font-weight:bold; background-color:#ffffff;">{p_key}</td>')
                                        first_p_row = False

                                    if first_t_row:
                                        html_list.append(f'<td rowspan="{t_rows_count}" style="background-color:#ffffff;">{t_key}</td>')
                                        first_t_row = False

                                    if first_o_row:
                                        html_list.append(f'<td rowspan="{o_rows_count}" style="font-weight:bold; background-color:#ffffff;">{o_key}</td>')
                                        first_o_row = False

                                    html_list.append(f'<td>{r["bid_val"]}</td>')
                                    html_list.append(f'<td>${sp:,.2f}</td>')
                                    html_list.append(f'<td{passed_cell_class}>{d0_r:.2f}%</td>')
                                    html_list.append(f'<td>${cpi_v:.2f}</td>')
                                    html_list.append(f'<td>${cpp_v:.2f}</td>')
                                    html_list.append(f'<td>{pay_rate:.2f}%</td>')
                                    html_list.append(f'<td>${arppu_v:.2f}</td>')
                                    html_list.append(f'<td>{roas_v:.2f}%</td>')
                                    html_list.append('</tr>')

                tot_d0 = (tot_rev / tot_tar * 100) if tot_tar > 0 else 0.0
                tot_cpi = (tot_sp / tot_act) if tot_act > 0 else 0.0
                tot_cpp = (tot_sp / tot_pay) if tot_pay > 0 else 0.0
                tot_pay_rate = (tot_pay / tot_act * 100) if tot_act > 0 else 0.0
                tot_arppu = (tot_rev / tot_pay) if tot_pay > 0 else 0.0
                tot_roas = (tot_rev / tot_sp * 100) if tot_sp > 0 else 0.0

                tot_passed_class = ' class="cell-passed-underline"' if tot_d0 >= 85.0 else (' class="cell-passed-text-only"' if tot_d0 >= 80.0 else '')

                span_val_tot = 5 if show_lang_col else 4
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${tot_sp:,.2f}</td>')
                html_list.append(f'<td{tot_passed_class}>{tot_d0:.2f}%</td>')
                html_list.append(f'<td>${tot_cpi:.2f}</td>')
                html_list.append(f'<td>${tot_cpp:.2f}</td>')
                html_list.append(f'<td>{tot_pay_rate:.2f}%</td>')
                html_list.append(f'<td>${tot_arppu:.2f}</td>')
                html_list.append(f'<td>{tot_roas:.2f}%</td>')
                html_list.append('</tr>')

                html_list.append('</tbody></table></div>')
                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_optimization_pivot_5level_by_lang(df_w2)

            # ---------------------------------------------------------------------
            # 🌐 【板块六】不同语种的账户时区数据透视表
            # ---------------------------------------------------------------------
            st.markdown("**【板块六】不同语种的账户时区数据透视表**")

            def render_timezone_pivot_by_lang(df_w2):
                if col_account == "不映射/无此列" or col_account not in df_w2.columns:
                    st.warning("⚠️ 未映射 账户名称 列，请在上方“第二步”设置！")
                    return

                g = p_w2.groupby(['lang', 'tz_class'], as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                tot_spend_all = g['spend'].sum()

                html_list = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                show_lang_col = (target_lang_filter is None)
                if show_lang_col:
                    html_list.append('<th>剧目/书籍语言</th>')
                html_list.append('<th>账户时区分类</th><th>花费 ($)</th><th>消耗占比 (%)</th><th>D0 达标率 (%)</th><th>ROAS (%)</th><th>CPI ($)</th><th>CPP ($)</th><th>注册拉新占比 (%)</th><th>付费拉新占比 (%)</th></tr></thead><tbody>')

                grand_sp, grand_rev, grand_tar, grand_act, grand_pay, grand_n_reg, grand_n_pay = 0, 0, 0, 0, 0, 0, 0

                lang_grouped = sorted(g.groupby('lang'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                for lang_val, group_lang in lang_grouped:
                    group_sorted = group_lang.sort_values('spend', ascending=False)
                    rowspan_count = len(group_sorted)
                    first_row = True

                    l_sp = group_sorted['spend'].sum()
                    l_rev = group_sorted['d0_rev'].sum()
                    l_tar = group_sorted['d0_target'].sum()
                    l_act = group_sorted['act'].sum()
                    l_pay = group_sorted['pay'].sum()
                    l_n_reg = group_sorted['new_reg'].sum()
                    l_n_pay = group_sorted['new_pay'].sum()

                    grand_sp += l_sp
                    grand_rev += l_rev
                    grand_tar += l_tar
                    grand_act += l_act
                    grand_pay += l_pay
                    grand_n_reg += l_n_reg
                    grand_n_pay += l_n_pay

                    for _, r in group_sorted.iterrows():
                        sp = r['spend']
                        p_pct = (sp / tot_spend_all * 100) if tot_spend_all > 0 else 0
                        d0_r = (r['d0_rev'] / r['d0_target'] * 100) if r['d0_target'] > 0 else 0
                        roas_v = (r['d0_rev'] / sp * 100) if sp > 0 else 0
                        cpi_v = (sp / r['act']) if r['act'] > 0 else 0
                        cpp_v = (sp / r['pay']) if r['pay'] > 0 else 0
                        reg_pct = (r['new_reg'] / r['act'] * 100) if r['act'] > 0 else 0
                        pay_pct = (r['new_pay'] / r['pay'] * 100) if r['pay'] > 0 else 0

                        passed_cell_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                        html_list.append('<tr>')
                        if show_lang_col and first_row:
                            html_list.append(f'<td rowspan="{rowspan_count}" style="font-weight:bold; background-color:#ffffff;">{lang_val}</td>')
                            first_row = False

                        html_list.append(f'<td>{r["tz_class"]}</td>')
                        html_list.append(f'<td>${sp:,.2f}</td>')
                        html_list.append(f'<td>{p_pct:.2f}%</td>')
                        html_list.append(f'<td{passed_cell_class}>{d0_r:.2f}%</td>')
                        html_list.append(f'<td>{roas_v:.2f}%</td>')
                        html_list.append(f'<td>${cpi_v:.2f}</td>')
                        html_list.append(f'<td>${cpp_v:.2f}</td>')
                        html_list.append(f'<td>{reg_pct:.2f}%</td>')
                        html_list.append(f'<td>{pay_pct:.2f}%</td>')
                        html_list.append('</tr>')

                    l_p_pct = (l_sp / tot_spend_all * 100) if tot_spend_all > 0 else 0
                    l_d0_r = (l_rev / l_tar * 100) if l_tar > 0 else 0
                    l_roas = (l_rev / l_sp * 100) if l_sp > 0 else 0
                    l_cpi = (l_sp / l_act) if l_act > 0 else 0
                    l_cpp = (l_sp / l_pay) if l_pay > 0 else 0
                    l_reg_pct = (l_n_reg / l_act * 100) if l_act > 0 else 0
                    l_pay_pct = (l_n_pay / l_pay * 100) if l_pay > 0 else 0
                    l_passed_class = ' class="cell-passed-underline"' if l_d0_r >= 85.0 else (' class="cell-passed-text-only"' if l_d0_r >= 80.0 else '')

                    span_val_lang = 2 if show_lang_col else 1
                    html_list.append('<tr class="row-lang-subtotal">')
                    html_list.append(f'<td colspan="{span_val_lang}">{lang_val} 小计</td>')
                    html_list.append(f'<td>${l_sp:,.2f}</td>')
                    html_list.append(f'<td>{l_p_pct:.2f}%</td>')
                    html_list.append(f'<td{l_passed_class}>{l_d0_r:.2f}%</td>')
                    html_list.append(f'<td>{l_roas:.2f}%</td>')
                    html_list.append(f'<td>${l_cpi:.2f}</td>')
                    html_list.append(f'<td>${l_cpp:.2f}</td>')
                    html_list.append(f'<td>{l_reg_pct:.2f}%</td>')
                    html_list.append(f'<td>{l_pay_pct:.2f}%</td>')
                    html_list.append('</tr>')

                g_d0 = (grand_rev / grand_tar * 100) if grand_tar > 0 else 0
                g_roas = (grand_rev / grand_sp * 100) if grand_sp > 0 else 0
                g_cpi = (grand_sp / grand_act) if grand_act > 0 else 0
                g_cpp = (grand_sp / grand_pay) if grand_pay > 0 else 0
                g_reg_p = (grand_n_reg / grand_act * 100) if grand_act > 0 else 0
                g_pay_p = (grand_n_pay / grand_pay * 100) if grand_pay > 0 else 0
                g_passed_class = ' class="cell-passed-underline"' if g_d0 >= 85.0 else (' class="cell-passed-text-only"' if g_d0 >= 80.0 else '')

                span_val_tot = 2 if show_lang_col else 1
                html_list.append('<tr class="row-grand-total">')
                html_list.append(f'<td colspan="{span_val_tot}">总计</td>')
                html_list.append(f'<td>${grand_sp:,.2f}</td>')
                html_list.append('<td>100.00%</td>')
                html_list.append(f'<td{g_passed_class}>{g_d0:.2f}%</td>')
                html_list.append(f'<td>{g_roas:.2f}%</td>')
                html_list.append(f'<td>${g_cpi:.2f}</td>')
                html_list.append(f'<td>${g_cpp:.2f}</td>')
                html_list.append(f'<td>{g_reg_p:.2f}%</td>')
                html_list.append(f'<td>{g_pay_p:.2f}%</td>')
                html_list.append('</tr></tbody></table></div>')

                st.markdown("".join(html_list), unsafe_allow_html=True)

            render_timezone_pivot_by_lang(df_w2)

            # =================================----------------------------------------
            # 三、团队管理
            # =================================----------------------------------------
            st.markdown('<div class="report-title">三、团队管理</div>', unsafe_allow_html=True)

            OPTIMIZER_WEEKLY_KPI = {
                'CZM': 63000.0,
                'lel': 45360.0,
                'LTT': 50400.0,
                'Mia': 42000.0,
                'Fanny': 42000.0,
                'Nora': 30450.0
            }

            def render_team_management_section(df_w1, df_w2):
                if col_optimizer == "不映射/无此列" or col_optimizer not in df_w2.columns:
                    st.warning("⚠️ 未映射 优化师 列，请在上方“第二步”设置！")
                    return

                w2 = df_w2.copy()
                w1 = df_w1.copy()

                w2['optimizer'] = w2[col_optimizer].astype(str).str.strip()
                w2['spend'] = clean_num(w2[col_spend])
                w2['d0_rev'] = clean_num(w2[col_d0_rev])
                w2['d0_target'] = clean_num(w2[col_d0_target])
                w2['act'] = clean_num(w2[col_act]) if col_act in w2 else 0
                w2['pay'] = clean_num(w2[col_pay]) if col_pay in w2 else 0
                
                # 📌 剔除零花费
                w2 = w2[w2['spend'] > 0]
                
                reg_r2 = clean_num(w2[col_reg_rate]) if col_reg_rate in w2 else pd.Series(0)
                pay_reg_r2 = clean_num(w2[col_pay_reg_rate]) if col_pay_reg_rate in w2 else pd.Series(0)
                if reg_r2.max() > 1.0: reg_r2 /= 100.0
                if pay_reg_r2.max() > 1.0: pay_reg_r2 /= 100.0
                w2['new_reg'] = w2['act'] * reg_r2
                w2['new_pay'] = w2['pay'] * pay_reg_r2

                w1['optimizer'] = w1[col_optimizer].astype(str).str.strip()
                w1['spend'] = clean_num(w1[col_spend])
                w1['d0_rev'] = clean_num(w1[col_d0_rev])
                w1['d0_target'] = clean_num(w1[col_d0_target])
                w1['act'] = clean_num(w1[col_act]) if col_act in w1 else 0
                w1['pay'] = clean_num(w1[col_pay]) if col_pay in w1 else 0
                
                w1 = w1[w1['spend'] > 0]
                
                reg_r1 = clean_num(w1[col_reg_rate]) if col_reg_rate in w1 else pd.Series(0)
                pay_reg_r1 = clean_num(w1[col_pay_reg_rate]) if col_pay_reg_rate in w1 else pd.Series(0)
                if reg_r1.max() > 1.0: reg_r1 /= 100.0
                if pay_reg_r1.max() > 1.0: pay_reg_r1 /= 100.0
                w1['new_reg'] = w1['act'] * reg_r1
                w1['new_pay'] = w1['pay'] * pay_reg_r1

                st.markdown('<div class="report-subtitle">📊 1. 本周组内优化师数据汇总透视表（含 KPI 完成率）</div>', unsafe_allow_html=True)

                g2 = w2.groupby('optimizer', as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                }).sort_values('spend', ascending=False)

                html_3a = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                html_3a.append('<th>优化师</th><th>花费 ($)</th><th>D0 达标率 (%)</th><th>周 KPI (7天) ($)</th><th>完成率 (%)</th><th>D0 ROAS (%)</th><th>付费用户新注册用户占比 (%)</th><th>新注册用户占比 (%)</th></tr></thead><tbody>')

                tot_sp, tot_rev, tot_tar, tot_act, tot_pay, tot_n_reg, tot_n_pay = 0, 0, 0, 0, 0, 0, 0

                for _, r in g2.iterrows():
                    opt = r['optimizer']
                    sp = r['spend']
                    rev = r['d0_rev']
                    tar = r['d0_target']
                    act_n = r['act']
                    pay_n = r['pay']

                    tot_sp += sp
                    tot_rev += rev
                    tot_tar += tar
                    tot_act += act_n
                    tot_pay += pay_n
                    tot_n_reg += r['new_reg']
                    tot_n_pay += r['new_pay']

                    d0_r = (rev / tar * 100) if tar > 0 else 0.0
                    roas_v = (rev / sp * 100) if sp > 0 else 0.0
                    pay_pct = (r['new_pay'] / pay_n * 100) if pay_n > 0 else 0.0
                    reg_pct = (r['new_reg'] / act_n * 100) if act_n > 0 else 0.0

                    kpi_target = OPTIMIZER_WEEKLY_KPI.get(opt, None)
                    if kpi_target:
                        kpi_str = f"${kpi_target:,.2f}"
                        completion_v = (sp / kpi_target * 100)
                        completion_str = f"{completion_v:.2f}%"
                    else:
                        kpi_str = "—"
                        completion_str = "—"

                    passed_cell_class = ' class="cell-passed"' if d0_r >= 85.0 else (' class="cell-passed-text-only"' if d0_r >= 80.0 else '')

                    html_3a.append(f'<tr><td>{opt}</td>')
                    html_3a.append(f'<td>${sp:,.2f}</td>')
                    html_3a.append(f'<td{passed_cell_class}>{d0_r:.2f}%</td>')
                    html_3a.append(f'<td>{kpi_str}</td>')
                    html_3a.append(f'<td>{completion_str}</td>')
                    html_3a.append(f'<td>{roas_v:.2f}%</td>')
                    html_3a.append(f'<td>{pay_pct:.2f}%</td>')
                    html_3a.append(f'<td>{reg_pct:.2f}%</td></tr>')

                tot_d0 = (tot_rev / tot_tar * 100) if tot_tar > 0 else 0.0
                tot_roas = (tot_rev / tot_sp * 100) if tot_sp > 0 else 0.0
                tot_pay_p = (tot_n_pay / tot_pay * 100) if tot_pay > 0 else 0.0
                tot_reg_p = (tot_n_reg / tot_act * 100) if tot_act > 0 else 0.0

                tot_passed_class = ' class="cell-passed-underline"' if tot_d0 >= 85.0 else (' class="cell-passed-text-only"' if tot_d0 >= 80.0 else '')

                html_3a.append('<tr class="row-highlight-yellow"><td>总计</td>')
                html_3a.append(f'<td>${tot_sp:,.2f}</td>')
                html_3a.append(f'<td{tot_passed_class}>{tot_d0:.2f}%</td>')
                html_3a.append('<td>—</td><td>—</td>')
                html_3a.append(f'<td>{tot_roas:.2f}%</td>')
                html_3a.append(f'<td>{tot_pay_p:.2f}%</td>')
                html_3a.append(f'<td>{tot_reg_p:.2f}%</td></tr>')

                html_3a.append('</tbody></table></div>')
                st.markdown("".join(html_3a), unsafe_allow_html=True)

                st.markdown('<div class="report-subtitle">📊 2. 优化师维度双周环比透视分析</div>', unsafe_allow_html=True)

                g1 = w1.groupby('optimizer', as_index=False).agg({
                    'spend': 'sum', 'd0_rev': 'sum', 'd0_target': 'sum',
                    'act': 'sum', 'pay': 'sum', 'new_reg': 'sum', 'new_pay': 'sum'
                })

                opt_grouped = sorted(g2.groupby('optimizer'), key=lambda x: x[1]['spend'].sum(), reverse=True)

                html_single = ['<div class="pivot-table-container"><table class="full-expand-table"><thead><tr>']
                html_single.append('<th>优化师</th><th>周期</th><th>消耗 ($)</th><th>消耗环比</th><th>D0 达标率 (%)</th><th>达标率环比</th><th>ROAS (%)</th><th>ROAS 环比</th><th>付费用户新注册用户占比 (%)</th><th>付费新注册占比环比</th></tr></thead><tbody>')

                for opt_key, group2 in opt_grouped:
                    group1 = g1[g1['optimizer'] == opt_key]

                    s2 = group2['spend'].sum()
                    rev2 = group2['d0_rev'].sum()
                    tar2 = group2['d0_target'].sum()
                    pay2_num = group2['pay'].sum()
                    n_pay2 = group2['new_pay'].sum()

                    d0_2 = (rev2 / tar2 * 100) if tar2 > 0 else 0.0
                    roas2 = (rev2 / s2 * 100) if s2 > 0 else 0.0
                    pay_pct2 = (n_pay2 / pay2_num * 100) if pay2_num > 0 else 0.0

                    s1 = group1['spend'].sum() if not group1.empty else 0.0
                    rev1 = group1['d0_rev'].sum() if not group1.empty else 0.0
                    tar1 = group1['d0_target'].sum() if not group1.empty else 0.0
                    pay1_num = group1['pay'].sum() if not group1.empty else 0.0
                    n_pay1 = group1['new_pay'].sum() if not group1.empty else 0.0

                    d0_1 = (rev1 / tar1 * 100) if tar1 > 0 and not group1.empty else 0.0
                    roas1 = (rev1 / s1 * 100) if s1 > 0 and not group1.empty else 0.0
                    pay_pct1 = (n_pay1 / pay1_num * 100) if pay1_num > 0 and not group1.empty else 0.0

                    spend_diff_pct = ((s2 - s1) / s1 * 100) if s1 > 0 else 0.0
                    d0_diff_pp = d0_2 - d0_1
                    roas_diff_pp = roas2 - roas1
                    pay_pct_diff_pp = pay_pct2 - pay_pct1

                    d0_2_class = ' class="cell-passed"' if d0_2 >= 85.0 else (' class="cell-passed-text-only"' if d0_2 >= 80.0 else '')
                    
                    s_diff_str, s_style = format_diff_with_arrow(spend_diff_pct)
                    d0_diff_str, d0_style = format_diff_with_arrow(d0_diff_pp, is_pp=True)
                    roas_diff_str, roas_style = format_diff_with_arrow(roas_diff_pp, is_pp=True)
                    pay_diff_str, pay_style = format_diff_with_arrow(pay_pct_diff_pp, is_pp=True)

                    html_single.append('<tr>')
                    html_single.append(f'<td rowspan="2" style="font-weight:bold; background-color:#ffffff;">{opt_key}</td>')
                    html_single.append('<td>本周</td>')
                    html_single.append(f'<td>${s2:,.2f}</td>')
                    html_single.append(f'<td style="{s_style}">{s_diff_str}</td>')
                    html_single.append(f'<td{d0_2_class}>{d0_2:.2f}%</td>')
                    html_single.append(f'<td style="{d0_style}">{d0_diff_str}</td>')
                    html_single.append(f'<td>{roas2:.2f}%</td>')
                    html_single.append(f'<td style="{roas_style}">{roas_diff_str}</td>')
                    html_single.append(f'<td>{pay_pct2:.2f}%</td>')
                    html_single.append(f'<td style="{pay_style}">{pay_diff_str}</td>')
                    html_single.append('</tr>')

                    d0_1_class = ' class="cell-passed"' if d0_1 >= 85.0 else (' class="cell-passed-text-only"' if d0_1 >= 80.0 else '')
                    html_single.append('<tr>')
                    html_single.append('<td>上周</td>')
                    html_single.append(f'<td class="cell-last-week">${s1:,.2f}</td><td>—</td>')
                    html_single.append(f'<td{d0_1_class}>{d0_1:.2f}%</td>')
                    html_single.append('<td>—</td>')
                    html_single.append(f'<td>{roas1:.2f}%</td><td>—</td>')
                    html_single.append(f'<td>{pay_pct1:.2f}%</td><td>—</td>')
                    html_single.append('</tr>')

                html_single.append('</tbody></table></div>')
                st.markdown("".join(html_single), unsafe_allow_html=True)

            render_team_management_section(df_w1, df_w2)

            st.markdown('<div class="report-subtitle">📝 文字复盘要点总结</div>', unsafe_allow_html=True)

            zi_group = show_level[show_level[col_copyright].str.contains('自制|自办', na=False)]
            cai_group = show_level[show_level[col_copyright].str.contains('采购|外包', na=False)]

            zi_pass_rate = (zi_group['is_new_passed'].sum() / zi_group['is_new_drama'].sum() * 100) if zi_group['is_new_drama'].sum() > 0 else 0
            cai_pass_rate = (cai_group['is_new_passed'].sum() / cai_group['is_new_drama'].sum() * 100) if cai_group['is_new_drama'].sum() > 0 else 0

            unpassed_new = show_level[show_level['is_new_drama'] & (~show_level['is_passed'])]
            waste_spend = unpassed_new['spend'].sum()

            st.markdown(f"""
            <div class="review-card">
                <div class="review-item"><b>1. 自制剧 vs 采购剧过测率对比：</b> 本周自制新剧过测率为 <b>{zi_pass_rate:.1f}%</b>，采购新剧过测率为 <b>{cai_pass_rate:.1f}%</b>。自制剧在题材贴合度及内容转化上表现更为稳定。</div>
                <div class="review-item"><b>2. 新剧测试空耗归因：</b> 本周未过测新剧共耗资 <b>${waste_spend:,.2f}</b>。空耗主要集中在部分前 3 秒留存较低的品类，建议针对此类未过测剧集缩短测试周期或微调素材开头。</div>
                <div class="review-item"><b>3. 在投剧目与团队表现：</b> 团队头部优化师花费贡献稳定，多位优化师 KPI 完成率超额达成；在投老剧整体达标率维持良好，持续提供稳健流水支撑。</div>
            </div>
            """, unsafe_allow_html=True)

    # 📌 Tabs 分发渲染
    with lang_tabs[0]:
        render_dashboard_content(target_lang_filter=None)
        
    with lang_tabs[1]:
        render_dashboard_content(target_lang_filter="繁体中文")
        
    with lang_tabs[2]:
        render_dashboard_content(target_lang_filter="韩语")

else:
    st.info("👆 请在顶部上传【上周】和【本周】两份 Excel 表格。")
