import io
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="短剧/漫剧综合数据看板", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """<style>
    h1 { font-size: 22px !important; margin-bottom: 10px !important; }
    h2, .stHeader { font-size: 18px !important; margin-top: 14px !important; margin-bottom: 8px !important; }
    h3, .stSubheader { font-size: 16px !important; margin-top: 10px !important; margin-bottom: 6px !important; }
    .dataframe-table, .custom-pivot-table { width: 100%; border-collapse: collapse; font-size: 13px; text-align: center; margin-bottom: 20px; }
    .dataframe-table th, .custom-pivot-table th { background-color: #F8F9FA; color: #333; font-weight: bold; padding: 8px 6px; border: 1px solid #DEE2E6; vertical-align: middle; }
    .dataframe-table td, .custom-pivot-table td { padding: 6px 4px; border: 1px solid #DEE2E6; vertical-align: middle; }
    .row-total { background-color: #FFE066 !important; font-weight: bold; }
    .col-highlight-yellow { background-color: #FFF9C4 !important; }
    .summary-d0-85 { color: #28A745 !important; font-weight: bold !important; text-decoration: underline !important; }
    .summary-d0-80 { color: #28A745 !important; font-weight: bold !important; }
    .detail-d0-85 { background-color: #D4EDDA !important; color: #28A745 !important; font-weight: bold !important; }
    .detail-d0-80 { color: #28A745 !important; font-weight: bold !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] { max-height: 160px; overflow-y: auto; }
    [data-testid="stSidebar"] [data-baseweb="tag"] { background-color: #FF4D4F !important; color: #FFF !important; margin: 2px; border-radius: 4px; font-weight: 500; font-size: 12px; }
    [data-testid="stSidebar"] [data-baseweb="tag"] span, [data-testid="stSidebar"] [data-baseweb="tag"] svg { color: #FFF !important; fill: #FFF !important; }
    .badge-warn-a { background-color: #FF4D4F; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-warn-b { background-color: #FA8C16; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-warn-c { background-color: #FADB14; color: #333; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .material-copy-card { background-color: #F0F5FF; border: 1px solid #ADC6FF; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; }
    .material-header-title { font-size: 16px; font-weight: bold; color: #1D39C4; margin-bottom: 8px; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; padding: 8px 16px; }
    </style>""",
    unsafe_allow_html=True,
)

st.title("🎬 短剧/漫剧综合数据看板")

def format_pct(val):
    if pd.isna(val) or val == "-" or val == "": return "0.00%"
    try:
        num = float(val)
        return f"{num * 100.0:.2f}%" if (abs(num) <= 2.0 and num != 0) else f"{num:.2f}%"
    except:
        return f"{val}"

def clean_rate_to_decimal(val):
    if pd.isna(val) or val == "-" or val == "": return 0.0
    try:
        num = float(val)
        return num / 100.0 if abs(num) > 2.0 else num
    except:
        return 0.0

@st.cache_data(ttl=7200, show_spinner="⚡ 正在构建数据缓存...")
def load_and_preprocess_excel(file_bytes):
    try: raw_df = pd.read_excel(io.BytesIO(file_bytes), engine="calamine")
    except:
        try: raw_df = pd.read_excel(io.BytesIO(file_bytes), engine="pyarrow")
        except: raw_df = pd.read_excel(io.BytesIO(file_bytes))

    defaults = {
        "优化师": "-", "花费": 0.0, "剧目/书籍语言": "未知语种", "剧目/书籍名称": "未知剧目",
        "剧目/书籍中文名称": "-", "剧目/书籍短ID": "-", "版权来源": "-", "pixel名称": "-",
        "优化方式": "-", "操作系统 (广告)": "-", "操作系统 (用户)": "-", "投放平台": "-",
        "推广目标": "-", "推广目标（属性）": "-", "广告类型": "-", "书籍类型": "未知类型",
        "漫剧类型": "未知类型", "首次投放的日期": "2020-01-01", "投放账户名称": "-",
        "广告创建时间": "2020-01-01", "素材名称": "-", "D0达标率": 0.0, "d0_roas": 0.0,
        "CPI": 0.0, "首日付费率": 0.0, "首日付费用户成本": 0.0, "Arppu- 付费用户人均付费金额": 0.0,
        "新注册用户占比": 0.0, "付费用户新注册用户占比": 0.0, "D0目标收入": 0.0, "d0_revenue": 0.0,
        "首日激活用户数": 0.0, "首日付费用户数": 0.0,
    }
    for k, v in defaults.items():
        if k not in raw_df.columns: raw_df[k] = v

    df = raw_df[(raw_df["优化师"] != "-") & (raw_df["花费"] > 0)].copy()
    df["日期_dt"] = pd.to_datetime(df["日期"])
    df["日期_str"] = df["日期_dt"].dt.strftime("%Y-%m-%d")
    df["首次投放_str"] = pd.to_datetime(df["首次投放的日期"]).dt.strftime("%Y-%m-%d")

    for nc in ["花费", "D0目标收入", "d0_revenue", "首日激活用户数", "首日付费用户数", "CPI", "首日付费用户成本", "Arppu- 付费用户人均付费金额"]:
        df[nc] = pd.to_numeric(df[nc], errors="coerce").fillna(0.0)

    for rc in ["D0达标率", "d0_roas", "首日付费率", "新注册用户占比", "付费用户新注册用户占比"]:
        df[rc] = pd.to_numeric(df[rc], errors="coerce").fillna(0.0).apply(clean_rate_to_decimal)

    df["账户时区"] = df["投放账户名称"].apply(lambda x: "+8 时区账户" if "+8-" in str(x) else ("0 时区账户" if "+0-" in str(x) else "美西账户"))
    return df

@st.cache_data(ttl=3600)
def get_cached_filter_options(df):
    t_col = "推广目标（属性）" if "推广目标（属性）" in df.columns else "推广目标"
    p_col = "书籍类型" if "书籍类型" in df.columns else "漫剧类型"
    return {
        "dramas": sorted(df["剧目/书籍短ID"].dropna().astype(str).unique().tolist()),
        "accounts": sorted(df["投放账户名称"].dropna().astype(str).unique().tolist()),
        "opts": sorted([o for o in df["优化师"].dropna().unique() if o != "-"]),
        "langs": sorted(df["剧目/书籍语言"].dropna().unique().tolist()) if "剧目/书籍语言" in df.columns else [],
        "copyrights": sorted(df["版权来源"].dropna().unique().tolist()) if "版权来源" in df.columns else [],
        "comic_types": sorted(df["漫剧类型"].dropna().unique().tolist()) if "漫剧类型" in df.columns else [],
        "platforms": sorted(df["投放平台"].dropna().unique().tolist()),
        "targets": sorted(df[t_col].dropna().unique().tolist()) if t_col in df.columns else [],
        "opt_modes": sorted(df["优化方式"].dropna().unique().tolist()) if "优化方式" in df.columns else [],
        "pixels": sorted(df["pixel名称"].dropna().astype(str).unique().tolist()),
        "types": sorted(df[p_col].dropna().unique().tolist()) if p_col in df.columns else [],
    }

@st.cache_data
def apply_fast_filter(df, sel_dramas, sel_accounts, sel_opts, sel_langs, sel_copyrights, sel_comic_types, sel_platforms, sel_targets, sel_opt_modes, sel_pixels, sel_types, start_date, end_date):
    mask = np.ones(len(df), dtype=bool)
    if sel_dramas: mask &= df["剧目/书籍短ID"].astype(str).isin(sel_dramas).values
    if sel_accounts: mask &= df["投放账户名称"].isin(sel_accounts).values
    if sel_opts: mask &= df["优化师"].isin(sel_opts).values
    if sel_langs and "剧目/书籍语言" in df.columns: mask &= df["剧目/书籍语言"].isin(sel_langs).values
    if sel_copyrights and "版权来源" in df.columns: mask &= df["版权来源"].isin(sel_copyrights).values
    if sel_comic_types and "漫剧类型" in df.columns: mask &= df["漫剧类型"].isin(sel_comic_types).values
    if sel_platforms: mask &= df["投放平台"].isin(sel_platforms).values
    t_col = "推广目标（属性）" if "推广目标（属性）" in df.columns else "推广目标"
    if sel_targets and t_col in df.columns: mask &= df[t_col].isin(sel_targets).values
    if sel_opt_modes and "优化方式" in df.columns: mask &= df["优化方式"].isin(sel_opt_modes).values
    if sel_pixels: mask &= df["pixel名称"].isin(sel_pixels).values
    p_col = "书籍类型" if "书籍类型" in df.columns else "漫剧类型"
    if sel_types and p_col in df.columns: mask &= df[p_col].isin(sel_types).values
    if start_date and end_date and "日期_dt" in df.columns:
        mask &= ((df["日期_dt"] >= pd.to_datetime(start_date)) & (df["日期_dt"] <= pd.to_datetime(end_date))).values
    return df[mask]

def compute_metrics(sub_df):
    c_cost = sub_df["花费"].sum()
    sum_tar = sub_df["D0目标收入"].sum() if "D0目标收入" in sub_df.columns else 0
    sum_rev = sub_df["d0_revenue"].sum() if "d0_revenue" in sub_df.columns else 0
    d0_reach = sum_rev / sum_tar if sum_tar > 0 else ((sub_df["D0达标率"] * sub_df["花费"]).sum() / (c_cost or 1))
    d0_roas = (sub_df["d0_roas"] * sub_df["花费"]).sum() / (c_cost or 1)
    sum_act = sub_df["首日激活用户数"].sum() if "首日激活用户数" in sub_df.columns else 0
    cpi = (c_cost / sum_act) if sum_act > 0 else ((sub_df["CPI"] * sub_df["花费"]).sum() / (c_cost or 1))
    sum_pay = sub_df["首日付费用户数"].sum() if "首日付费用户数" in sub_df.columns else 0
    cpp = (c_cost / sum_pay) if sum_pay > 0 else ((sub_df["首日付费用户成本"] * sub_df["花费"]).sum() / (c_cost or 1))
    pay_rate = (sum_pay / sum_act) if (sum_act > 0 and sum_pay > 0) else ((sub_df["首日付费率"] * sub_df["花费"]).sum() / (c_cost or 1))
    arppu = (sub_df["Arppu- 付费用户人均付费金额"] * sub_df["花费"]).sum() / (c_cost or 1)
    reg_ratio = ((sub_df["首日激活用户数"] * sub_df["新注册用户占比"]).sum() / sum_act) if sum_act > 0 else ((sub_df["新注册用户占比"] * sub_df["花费"]).sum() / (c_cost or 1))
    pay_reg_ratio = ((sub_df["首日付费用户数"] * sub_df["付费用户新注册用户占比"]).sum() / sum_pay) if sum_pay > 0 else ((sub_df["付费用户新注册用户占比"] * sub_df["花费"]).sum() / (c_cost or 1))

    return {
        "花费": c_cost, "达标率": d0_reach, "D0 ROAS": d0_roas, "CPI": cpi, "首日付费率": pay_rate,
        "CPP": cpp, "ARPPU": arppu, "新注册用户占比": reg_ratio, "付费用户新注册用户占比": pay_reg_ratio,
        "首日激活用户数": sum_act, "首日付费用户数": sum_pay
    }

@st.cache_data(ttl=3600, show_spinner=False)
def generate_pivot_html(df_data, dims, metric_cols, total_title="大盘总计 / 全局汇总"):
    dims = [d for d in dims if d in df_data.columns]
    if not dims or df_data.empty: return "<p style='color:#666;'>无符合条件的数据</p>"
    total_cost = df_data["花费"].sum()

    def build_tree(sub, idx):
        if idx >= len(dims): return []
        cdim = dims[idx]
        nodes = []
        grouped = sub.groupby(cdim, observed=True)
        sorted_keys = grouped["花费"].sum().sort_values(ascending=False).index.tolist()
        for k in sorted_keys:
            s_df = grouped.get_group(k)
            node = {"val": k, "cost": s_df["花费"].sum()}
            if idx == len(dims) - 1:
                node["leaf_data"] = compute_metrics(s_df)
                node["rowspan"] = 1
            else:
                children = build_tree(s_df, idx + 1)
                node["children"] = children
                node["rowspan"] = sum(c["rowspan"] for c in children)
            nodes.append(node)
        return nodes

    tree = build_tree(df_data, 0)

    def flatten(nodes, prefix=None):
        if prefix is None: prefix = []
        rows = []
        for n in nodes:
            curr = prefix + [{"val": n["val"], "rowspan": n["rowspan"]}]
            if "children" in n: rows.extend(flatten(n["children"], curr))
            else: rows.append({"dims": curr, "metrics": n["leaf_data"]})
        return rows

    flat_rows = flatten(tree)
    html = '<table class="dataframe-table custom-pivot-table"><thead><tr>'
    for d in dims: html += f"<th>{d}</th>"
    for mc in metric_cols: html += f"<th>{mc}</th>"
    html += "</tr></thead><tbody>"

    if total_cost > 0:
        tot_m = compute_metrics(df_data)
        d0_cell = 'class="summary-d0-85"' if tot_m['达标率'] >= 0.85 else ('class="summary-d0-80"' if tot_m['达标率'] >= 0.80 else '')
        html += f'<tr class="row-total"><td colspan="{len(dims)}">{total_title}</td>'
        for mc in metric_cols:
            if mc == "花费 ($)": html += f"<td>${total_cost:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]: html += "<td>100.00%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]: html += f"<td {d0_cell}>{format_pct(tot_m['达标率'])}</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]: html += f"<td>{format_pct(tot_m['D0 ROAS'])}</td>"
            elif mc == "CPI ($)": html += f"<td>${tot_m['CPI']:.2f}</td>"
            elif mc == "首日付费率 (%)": html += f"<td>{format_pct(tot_m['首日付费率'])}</td>"
            elif mc == "CPP ($)": html += f"<td>${tot_m['CPP']:.2f}</td>"
            elif mc == "ARPPU ($)": html += f"<td>${tot_m['ARPPU']:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]: html += f"<td>{format_pct(tot_m['新注册用户占比'])}</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]: html += f"<td>{format_pct(tot_m['付费用户新注册用户占比'])}</td>"
            elif mc == "首日激活用户数": html += f"<td>{int(tot_m['首日激活用户数']):,}</td>"
            elif mc == "首日付费用户数": html += f"<td>{int(tot_m['首日付费用户数']):,}</td>"
        html += "</tr>"

    rendered_dims = [0] * len(dims)
    yellow_dims = ["剧目/书籍语言", "优化师", "剧目/书籍名称", "剧目/书籍短ID", "投放平台", "账户时区", "操作系统 (广告)", "素材名称", "推广目标"]

    for r in flat_rows:
        html += "<tr>"
        for d_idx, d_info in enumerate(r["dims"]):
            if rendered_dims[d_idx] == 0:
                rs, val = d_info["rowspan"], d_info["val"]
                c_cls = 'class="col-highlight-yellow"' if dims[d_idx] in yellow_dims else ""
                html += f'<td rowspan="{rs}" {c_cls}>{val}</td>' if rs > 1 else f"<td {c_cls}>{val}</td>"
                rendered_dims[d_idx] = rs - 1
            else: rendered_dims[d_idx] -= 1

        m = r["metrics"]
        d0_val = m["达标率"]
        d0_cell = 'class="detail-d0-85"' if d0_val >= 0.85 else ('class="detail-d0-80"' if d0_val >= 0.80 else '')
        cost_ratio = (m["花费"] / total_cost * 100) if total_cost > 0 else 0

        for mc in metric_cols:
            if mc == "花费 ($)": html += f"<td>${m['花费']:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]: html += f"<td>{cost_ratio:.2f}%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]: html += f"<td {d0_cell}>{format_pct(d0_val)}</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]: html += f"<td>{format_pct(m['D0 ROAS'])}</td>"
            elif mc == "CPI ($)": html += f"<td>${m['CPI']:.2f}</td>"
            elif mc == "首日付费率 (%)": html += f"<td>{format_pct(m['首日付费率'])}</td>"
            elif mc == "CPP ($)": html += f"<td>${m['CPP']:.2f}</td>"
            elif mc == "ARPPU ($)": html += f"<td>${m['ARPPU']:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]: html += f"<td>{format_pct(m['新注册用户占比'])}</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]: html += f"<td>{format_pct(m['付费用户新注册用户占比'])}</td>"
            elif mc == "首日激活用户数": html += f"<td>{int(m['首日激活用户数']):,}</td>"
            elif mc == "首日付费用户数": html += f"<td>{int(m['首日付费用户数']):,}</td>"
            elif mc == "手动出价的金额/系数":
                bid = df_data[(df_data[dims[-1]] == r["dims"][-1]["val"])]["手动出价的金额/系数"].iloc[0] if "手动出价的金额/系数" in df_data.columns else "-"
                html += f"<td>{bid}</td>"
        html += "</tr>"
    return html + "</tbody></table>"

# -----------------------------------------------------------------------------
# 主控制逻辑
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("📂 请上传数据底表 (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = load_and_preprocess_excel(uploaded_file.getvalue())
    all_dates = sorted(df["日期_str"].unique(), reverse=True)
    latest_date = all_dates[0] if len(all_dates) > 0 else ""
    options = get_cached_filter_options(df)
    min_date, max_date = df["日期_dt"].min().date(), df["日期_dt"].max().date()
    first_launch_date = df["首次投放的日期"].min() if "首次投放的日期" in df.columns else latest_date

    with st.sidebar:
        st.header("🎛️ 多维切片筛选")
        with st.form("sidebar_filter_form"):
            st.form_submit_button("🚀 一键应用筛选", use_container_width=True)
            sel_dramas = st.multiselect("🔍 搜索/筛选 推广剧目/书籍短ID", options=options["dramas"], default=options["dramas"])
            date_range = st.date_input("选择时间范围", value=[min_date, max_date])
            sel_accounts = st.multiselect("筛选 投放账户名称", options=options["accounts"], default=options["accounts"])
            sel_opts = st.multiselect("筛选 优化师", options=options["opts"], default=options["opts"])
            sel_langs = st.multiselect("筛选 剧目/书籍语言", options=options["langs"], default=options["langs"])
            sel_copyrights = st.multiselect("筛选 版权来源", options=options["copyrights"], default=options["copyrights"])
            sel_comic_types = st.multiselect("筛选 漫剧类型", options=options["comic_types"], default=options["comic_types"])
            sel_platforms = st.multiselect("筛选 投放平台", options=options["platforms"], default=options["platforms"])
            sel_targets = st.multiselect("筛选 推广目标", options=options["targets"], default=options["targets"])
            sel_opt_modes = st.multiselect("筛选 优化方式", options=options["opt_modes"], default=options["opt_modes"])
            sel_pixels = st.multiselect("筛选 pixel名称", options=options["pixels"], default=options["pixels"])
            sel_types = st.multiselect("筛选 书籍类型", options=options["types"], default=options["types"])
            enable_read_mode = st.toggle("💡 阅读模式 (十字高亮)", value=True)

    start_d = date_range[0] if len(date_range) > 0 else min_date
    end_d = date_range[1] if len(date_range) > 1 else max_date
    df_filtered = apply_fast_filter(df, sel_dramas, sel_accounts, sel_opts, sel_langs, sel_copyrights, sel_comic_types, sel_platforms, sel_targets, sel_opt_modes, sel_pixels, sel_types, start_d, end_d)

    if enable_read_mode:
        st.markdown("""<script>
        (function() {
            function initCrosshair() {
                const tables = document.querySelectorAll('table.dataframe-table');
                tables.forEach(table => {
                    if (table.dataset.crosshairInjected) return;
                    table.dataset.crosshairInjected = "true";
                    table.addEventListener('mouseover', (e) => {
                        const td = e.target.closest('td, th');
                        if (!td) return;
                        table.querySelectorAll('.crosshair-row, .crosshair-col, .crosshair-focus').forEach(el => el.classList.remove('crosshair-row', 'crosshair-col', 'crosshair-focus'));
                        const tr = td.parentElement;
                        const idx = td.cellIndex;
                        td.classList.add('crosshair-focus');
                        tr.querySelectorAll('td, th').forEach(c => c.classList.add('crosshair-row'));
                        table.querySelectorAll('tr').forEach(row => { if (row.cells[idx]) row.cells[idx].classList.add('crosshair-col'); });
                    });
                    table.addEventListener('mouseleave', () => { table.querySelectorAll('.crosshair-row, .crosshair-col, .crosshair-focus').forEach(el => el.classList.remove('crosshair-row', 'crosshair-col', 'crosshair-focus')); });
                });
            }
            setTimeout(initCrosshair, 800);
        })();
        </script>""", unsafe_allow_html=True)

    sheet1, sheet2, sheet3, sheet4, sheet5, sheet6 = st.tabs(["📊 第一部分：整体数据看板", "📌 第二部分：多维度分析", "📈 第三部分：投放维度拆分", "⚙️ 第四部分：优化方式&出价深度拆分", "🎯 第五部分：单剧目深度透视", "🎬 第六部分：素材分析"])

    with sheet1:
        p1_t1, p1_t2 = st.tabs(["📊 总体全量", "📈 每日趋势"])
        with p1_t1:
            if df_filtered.empty: st.warning("无可展现数据。")
            else: st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "优化师"], ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="大盘总体 (全量汇总)"), unsafe_allow_html=True)
        with p1_t2:
            if not df_filtered.empty: st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "优化师", "日期_str"], ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="全期汇总"), unsafe_allow_html=True)

        if not df_filtered.empty:
            qualified = [s for s, g in df_filtered.groupby("剧目/书籍名称") if compute_metrics(g)["花费"] > 100 and compute_metrics(g)["达标率"] >= 0.80]
            df_q = df_filtered[df_filtered["剧目/书籍名称"].isin(qualified)]
            if not df_q.empty:
                st.markdown('##### 📌 表格 4: 达标优化剧集 <span style="color:red; font-weight:bold;">(重点优化剧目)</span>', unsafe_allow_html=True)
                st.markdown(generate_pivot_html(df_q, ["剧目/书籍语言", "投放平台", "剧目/书籍名称", "剧目/书籍短ID"], ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="重点优化剧集汇总"), unsafe_allow_html=True)

    with sheet2:
        if not df_filtered.empty:
            acc_m = ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)"]
            pix_m = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "CPP ($)", "首日付费率 (%)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
            st.subheader("📊 表格 1: 投放平台 × 账户 明细表")
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "账户时区", "投放账户名称"], acc_m, total_title="账户汇总"), unsafe_allow_html=True)
            st.subheader("🎯 表格 2: 投放平台 × Pixel 名称 明细表")
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "pixel名称", "投放平台", "推广目标", "优化方式"], pix_m, total_title="Pixel汇总"), unsafe_allow_html=True)
            st.subheader("🚀 表格 3: 投放平台 × 推广目标 × 账户名称 明细表")
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "推广目标", "投放账户名称"], acc_m, total_title="目标账户汇总"), unsafe_allow_html=True)

    with sheet3:
        p2_t1, p2_t2 = st.tabs(["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"])
        def render_p2(df_w, is_daily):
            if df_w.empty: return
            suff = f"(当日明细 - {latest_date})" if is_daily else "(全量汇总)"
            common_m = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
            st.markdown(generate_pivot_html(df_w, ["剧目/书籍语言", "版权来源", "漫剧类型" if "漫剧类型" in df_w.columns else "书籍类型"], common_m, total_title="类型拆分汇总"), unsafe_allow_html=True)
        with p2_t1: render_p2(df_filtered[df_filtered["日期_str"] == latest_date], True)
        with p2_t2: render_p2(df_filtered, False)

    with sheet4:
        st.subheader("⚙️ 第四部分：优化方式&出价深度拆分")
        if not df_filtered.empty:
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数"], ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="优化方式汇总"), unsafe_allow_html=True)

    with sheet5:
        if len(sel_dramas) == 1:
            single_id = sel_dramas[0]
            df_s = df[df["剧目/书籍短ID"].astype(str) == single_id]
            if not df_s.empty:
                st.header(f"🎯 {df_s['剧目/书籍名称'].iloc[0]} - {single_id} 深度透视")
                p5_m = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
                st.markdown(generate_pivot_html(df_s, ["投放平台"], p5_m, total_title="平台汇总"), unsafe_allow_html=True)
        else:
            st.info("💡 提示：请在左侧侧边栏【🔍 搜索/筛选 推广剧目/书籍短ID】中仅选择单个剧目！")

    with sheet6:
        m_t1, m_t2 = st.tabs(["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"])
        def render_mat(df_w, is_daily):
            if df_w.empty: return
            qual = []
            g_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "素材名称"]
            for k, sub in df_w.groupby(g_dims):
                m_res = compute_metrics(sub)
                if m_res["达标率"] >= 0.80 and m_res["花费"] > 0: qual.append(k)
            if qual:
                df_q = df_w[df_w.set_index(g_dims).index.isin(qual)]
                lines = []
                for (lang, name, sid), sub_d in df_q.groupby(["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID"]):
                    lines.append(f"{name} - {sid}\n达标素材清单:")
                    for mn in sorted(sub_d["素材名称"].dropna().unique().tolist()): lines.append(mn)
                    lines.append("")
                st.markdown('<div class="material-copy-card"><div class="material-header-title">📋 复制达标素材名称列表</div>', unsafe_allow_html=True)
                with st.expander("点击展开文本"): st.text_area("素材清单:", value="\n".join(lines), height=200)
                st.markdown("</div>", unsafe_allow_html=True)

                rows = []
                for k, sub in df_q.groupby(g_dims):
                    mr = compute_metrics(sub)
                    rows.append({
                        "剧目/书籍语言": k[0], "剧目/书籍名称": k[1], "剧目ID": k[2], "素材名称": k[3],
                        "花费 ($)": f"${mr['花费']:,.2f}",
                        "D0 达标率 (%)": format_pct(mr["达标率"]),
                        "首日付费用户数": int(mr["首日付费用户数"]),
                        "D0 ROAS (%)": format_pct(mr["D0 ROAS"]),
                        "CPI ($)": f"${mr['CPI']:.2f}",
                        "首日付费率 (%)": format_pct(mr["首日付费率"]),
                        "首日激活用户数": int(mr["首日激活用户数"]),
                        "新注册用户占比 (%)": format_pct(mr["新注册用户占比"]),
                        "付费用户新注册用户占比 (%)": format_pct(mr["付费用户新注册用户占比"])
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            else: st.info("暂无达标素材。")
        with m_t1: render_mat(df_filtered[df_filtered["日期_str"] == latest_date], True)
        with m_t2: render_mat(df_filtered, False)
