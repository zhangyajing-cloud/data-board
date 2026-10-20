import io
import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# 1. 页面基本配置与全局 CSS 样式渲染
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="短剧/漫剧综合数据看板",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* 标题字号全局调小，版面紧凑精干 */
    h1 { font-size: 22px !important; margin-bottom: 10px !important; }
    h2, .stHeader { font-size: 18px !important; margin-top: 14px !important; margin-bottom: 8px !important; }
    h3, .stSubheader { font-size: 16px !important; margin-top: 10px !important; margin-bottom: 6px !important; }
    h4, h5 { font-size: 14px !important; }

    /* 全局样式与自定义表格样式 */
    .dataframe-table, .custom-pivot-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        text-align: center;
        margin-bottom: 20px;
    }
    .dataframe-table th, .custom-pivot-table th {
        background-color: #F8F9FA;
        color: #333333;
        font-weight: bold;
        padding: 8px 6px;
        border: 1px solid #DEE2E6;
        vertical-align: middle;
    }
    .dataframe-table td, .custom-pivot-table td {
        padding: 6px 4px;
        border: 1px solid #DEE2E6;
        vertical-align: middle;
    }

    /* 汇总行背景色硬性保留 */
    .row-total { background-color: #FFE066 !important; font-weight: bold; }
    .row-subtotal { background-color: #FFF9C4 !important; font-weight: bold; }
    .col-highlight-yellow { background-color: #FFF9C4 !important; }

    /* 达标率（>80% / >85%）文本与背景样式分类定义 */
    .summary-d0-85 { color: #28A745 !important; font-weight: bold !important; text-decoration: underline !important; }
    .summary-d0-80 { color: #28A745 !important; font-weight: bold !important; }

    .detail-d0-85 { background-color: #D4EDDA !important; color: #28A745 !important; font-weight: bold !important; }
    .detail-d0-80 { color: #28A745 !important; font-weight: bold !important; }

    /* 侧边栏多选标签 (Tags) 红色卡片视觉重构 */
    [data-testid="stSidebar"] [data-baseweb="select"] { max-height: 160px; overflow-y: auto; }
    [data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #FF4D4F !important;
        color: #FFFFFF !important;
        margin: 2px;
        border-radius: 4px;
        font-weight: 500;
        font-size: 12px;
    }
    [data-testid="stSidebar"] [data-baseweb="tag"] span { color: #FFFFFF !important; }
    [data-testid="stSidebar"] [data-baseweb="tag"] svg { fill: #FFFFFF !important; }

    /* 预警标签颜色 */
    .badge-warn-a { background-color: #FF4D4F; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-warn-b { background-color: #FA8C16; color: white; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-warn-c { background-color: #FADB14; color: #333; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    
    .strategy-box { background-color: #F6FFED; border: 1px solid #B7EB8F; border-radius: 6px; padding: 12px 16px; margin-top: 10px; margin-bottom: 20px; font-size: 13px; line-height: 1.6; }
    .strategy-box-title { font-weight: bold; color: #52C41A; font-size: 14px; margin-bottom: 6px; }

    .material-copy-card { background-color: #F0F5FF; border: 1px solid #ADC6FF; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; }
    .material-header-title { font-size: 16px; font-weight: bold; color: #1D39C4; margin-bottom: 8px; }
    
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; padding: 8px 16px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🎬 短剧/漫剧综合数据看板")

# -----------------------------------------------------------------------------
# 一、 核心百分比格式化与平移 2 位修复逻辑
# -----------------------------------------------------------------------------
def format_pct(val):
    """【全局百分比平移修复】确保所有百分比指标向右平移 2 位显示 (例如 0.85 -> 85.00%, 1.12 -> 112.00%)"""
    if pd.isna(val) or val == "-" or val == "":
        return "0.00%"
    try:
        num = float(val)
        if abs(num) <= 2.0 and num != 0:
            return f"{num * 100.0:.2f}%"
        return f"{num:.2f}%"
    except:
        return f"{val}"


def clean_num(val):
    try:
        if pd.isna(val) or val == "-" or val == "":
            return 0.0
        return float(val)
    except:
        return 0.0


def clean_rate_to_decimal(val):
    """清理底层 Dataframe 比率，统一为标准的 0~1/1+ 纯小数"""
    if pd.isna(val) or val == "-" or val == "":
        return 0.0
    try:
        num = float(val)
        if abs(num) > 2.0:
            return num / 100.0
        return num
    except:
        return 0.0


# -----------------------------------------------------------------------------
# 缓存与数据读取
# -----------------------------------------------------------------------------
@st.cache_data(ttl=7200, show_spinner="⚡ 正在读入并构建全局数据缓存...")
def load_and_preprocess_excel(file_bytes):
    try:
        raw_df = pd.read_excel(io.BytesIO(file_bytes), engine="calamine")
    except Exception:
        try:
            raw_df = pd.read_excel(io.BytesIO(file_bytes), engine="pyarrow")
        except Exception:
            raw_df = pd.read_excel(io.BytesIO(file_bytes))

    required_defaults = {
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
    for col_name, default_val in required_defaults.items():
        if col_name not in raw_df.columns:
            raw_df[col_name] = default_val

    df = raw_df[(raw_df["优化师"] != "-") & (raw_df["花费"] > 0)].copy()

    df["日期_dt"] = pd.to_datetime(df["日期"])
    df["日期_str"] = df["日期_dt"].dt.strftime("%Y-%m-%d")
    df["首次投放_str"] = pd.to_datetime(df["首次投放的日期"]).dt.strftime("%Y-%m-%d")

    for num_c in ["花费", "D0目标收入", "d0_revenue", "首日激活用户数", "首日付费用户数", "CPI", "首日付费用户成本", "Arppu- 付费用户人均付费金额"]:
        df[num_c] = pd.to_numeric(df[num_c], errors="coerce").fillna(0.0)

    for rate_c in ["D0达标率", "d0_roas", "首日付费率", "新注册用户占比", "付费用户新注册用户占比"]:
        df[rate_c] = pd.to_numeric(df[rate_c], errors="coerce").fillna(0.0)
        df[rate_c] = df[rate_c].apply(clean_rate_to_decimal)

    def get_timezone(acc_name):
        acc = str(acc_name)
        if "+8-" in acc: return "+8 时区账户"
        elif "+0-" in acc: return "0 时区账户"
        return "美西账户"

    df["账户时区"] = df["投放账户名称"].apply(get_timezone)
    return df


@st.cache_data(ttl=3600)
def get_cached_filter_options(df):
    target_col = "推广目标（属性）" if "推广目标（属性）" in df.columns else "推广目标"
    type_col = "书籍类型" if "书籍类型" in df.columns else "漫剧类型"
    
    return {
        "dramas": sorted(df["剧目/书籍短ID"].dropna().astype(str).unique().tolist()),
        "accounts": sorted(df["投放账户名称"].dropna().astype(str).unique().tolist()),
        "opts": sorted([o for o in df["优化师"].dropna().unique() if o != "-"]),
        "langs": sorted(df["剧目/书籍语言"].dropna().unique().tolist()) if "剧目/书籍语言" in df.columns else [],
        "copyrights": sorted(df["版权来源"].dropna().unique().tolist()) if "版权来源" in df.columns else [],
        "comic_types": sorted(df["漫剧类型"].dropna().unique().tolist()) if "漫剧类型" in df.columns else [],
        "platforms": sorted(df["投放平台"].dropna().unique().tolist()),
        "targets": sorted(df[target_col].dropna().unique().tolist()) if target_col in df.columns else [],
        "opt_modes": sorted(df["优化方式"].dropna().unique().tolist()) if "优化方式" in df.columns else [],
        "pixels": sorted(df["pixel名称"].dropna().astype(str).unique().tolist()),
        "types": sorted(df[type_col].dropna().unique().tolist()) if type_col in df.columns else [],
    }


@st.cache_data
def apply_fast_filter(
    df, sel_dramas, sel_accounts, sel_opts, sel_langs, sel_copyrights,
    sel_comic_types, sel_platforms, sel_targets, sel_opt_modes, sel_pixels,
    sel_types, start_date, end_date
):
    mask = np.ones(len(df), dtype=bool)

    if sel_dramas:
        mask &= df["剧目/书籍短ID"].astype(str).isin(sel_dramas).values
    if sel_accounts:
        mask &= df["投放账户名称"].isin(sel_accounts).values
    if sel_opts:
        mask &= df["优化师"].isin(sel_opts).values
    if sel_langs and "剧目/书籍语言" in df.columns:
        mask &= df["剧目/书籍语言"].isin(sel_langs).values
    if sel_copyrights and "版权来源" in df.columns:
        mask &= df["版权来源"].isin(sel_copyrights).values
    if sel_comic_types and "漫剧类型" in df.columns:
        mask &= df["漫剧类型"].isin(sel_comic_types).values
    if sel_platforms:
        mask &= df["投放平台"].isin(sel_platforms).values

    target_col = "推广目标（属性）" if "推广目标（属性）" in df.columns else "推广目标"
    if sel_targets and target_col in df.columns:
        mask &= df[target_col].isin(sel_targets).values

    if sel_opt_modes and "优化方式" in df.columns:
        mask &= df["优化方式"].isin(sel_opt_modes).values
    if sel_pixels:
        mask &= df["pixel名称"].isin(sel_pixels).values

    type_col = "书籍类型" if "书籍类型" in df.columns else "漫剧类型"
    if sel_types and type_col in df.columns:
        mask &= df[type_col].isin(sel_types).values

    if start_date and end_date and "日期_dt" in df.columns:
        s_date = pd.to_datetime(start_date)
        e_date = pd.to_datetime(end_date)
        mask &= ((df["日期_dt"] >= s_date) & (df["日期_dt"] <= e_date)).values

    return df[mask]


def compute_metrics(sub_df):
    c_cost = sub_df["花费"].sum()

    sum_d0_target = sub_df["D0目标收入"].sum() if "D0目标收入" in sub_df.columns else 0
    sum_d0_rev = sub_df["d0_revenue"].sum() if "d0_revenue" in sub_df.columns else 0
    
    if sum_d0_target > 0:
        d0_reach = sum_d0_rev / sum_d0_target
    else:
        d0_reach = (sub_df["D0达标率"] * sub_df["花费"]).sum() / (c_cost or 1)

    d0_roas = (sub_df["d0_roas"] * sub_df["花费"]).sum() / (c_cost or 1)

    sum_active = sub_df["首日激活用户数"].sum() if "首日激活用户数" in sub_df.columns else 0
    cpi = (c_cost / sum_active) if sum_active > 0 else ((sub_df["CPI"] * sub_df["花费"]).sum() / (c_cost or 1))

    sum_pay_users = sub_df["首日付费用户数"].sum() if "首日付费用户数" in sub_df.columns else 0
    cpp = (c_cost / sum_pay_users) if sum_pay_users > 0 else ((sub_df["首日付费用户成本"] * sub_df["花费"]).sum() / (c_cost or 1))

    pay_rate = (sum_pay_users / sum_active) if (sum_active > 0 and sum_pay_users > 0) else ((sub_df["首日付费率"] * sub_df["花费"]).sum() / (c_cost or 1))

    arppu = (sub_df["Arppu- 付费用户人均付费金额"] * sub_df["花费"]).sum() / (c_cost or 1)

    if sum_active > 0:
        reg_ratio = (sub_df["首日激活用户数"] * sub_df["新注册用户占比"]).sum() / sum_active
    else:
        reg_ratio = (sub_df["新注册用户占比"] * sub_df["花费"]).sum() / (c_cost or 1)

    if sum_pay_users > 0:
        pay_reg_ratio = (sub_df["首日付费用户数"] * sub_df["付费用户新注册用户占比"]).sum() / sum_pay_users
    else:
        pay_reg_ratio = (sub_df["付费用户新注册用户占比"] * sub_df["花费"]).sum() / (c_cost or 1)

    return {
        "花费": c_cost,
        "达标率": d0_reach,
        "D0 ROAS": d0_roas,
        "CPI": cpi,
        "首日付费率": pay_rate,
        "CPP": cpp,
        "ARPPU": arppu,
        "新注册用户占比": reg_ratio,
        "付费用户新注册用户占比": pay_reg_ratio,
        "首日激活用户数": sum_active,
        "首日付费用户数": sum_pay_users,
    }


# 透视 HTML 生成函数
@st.cache_data(ttl=3600, show_spinner=False)
def generate_pivot_html(
    df_data, dims, metric_cols, total_title="大盘总计 / 全局汇总"
):
    dims = [d for d in dims if d in df_data.columns]
    if not dims or df_data.empty:
        return "<p style='color:#666;'>无符合条件的数据或维度字段缺失</p>"

    total_cost = df_data["花费"].sum()

    def build_tree(data_subset, dim_index):
        if dim_index >= len(dims):
            return []

        current_dim = dims[dim_index]
        nodes = []
        grouped = data_subset.groupby(current_dim, observed=True)
        sorted_keys = grouped["花费"].sum().sort_values(ascending=False).index.tolist()

        for k in sorted_keys:
            sub_df = grouped.get_group(k)
            c_cost = sub_df["花费"].sum()

            node = {
                "dim": current_dim,
                "val": k,
                "cost": c_cost,
                "sub_df": sub_df,
            }

            if dim_index == len(dims) - 1:
                node["leaf_data"] = compute_metrics(sub_df)
                node["rowspan"] = 1
            else:
                children = build_tree(sub_df, dim_index + 1)
                node["children"] = children
                node["rowspan"] = sum(c["rowspan"] for c in children)

            nodes.append(node)
        return nodes

    tree = build_tree(df_data, 0)

    def flatten_tree_to_rows(nodes, current_row_prefix=None):
        if current_row_prefix is None:
            current_row_prefix = []
        flat_rows = []
        for n in nodes:
            prefix = current_row_prefix + [{"val": n["val"], "rowspan": n["rowspan"]}]
            if "children" in n:
                flat_rows.extend(flatten_tree_to_rows(n["children"], prefix))
            else:
                flat_rows.append({"dims": prefix, "metrics": n["leaf_data"]})
        return flat_rows

    flat_rows = flatten_tree_to_rows(tree)

    html = '<table class="dataframe-table custom-pivot-table"><thead><tr>'
    for d in dims:
        html += f"<th>{d}</th>"
    for mc in metric_cols:
        html += f"<th>{mc}</th>"
    html += "</tr></thead><tbody>"

    # 大盘总计行
    if total_cost > 0:
        tot_m = compute_metrics(df_data)
        tot_d0 = tot_m["达标率"]
        d0_cell = 'class="summary-d0-85"' if tot_d0 >= 0.85 else ('class="summary-d0-80"' if tot_d0 >= 0.80 else '')

        html += f'<tr class="row-total"><td colspan="{len(dims)}">{total_title}</td>'
        for mc in metric_cols:
            if mc == "花费 ($)":
                html += f"<td>${total_cost:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]:
                html += "<td>100.00%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]:
                html += f"<td {d0_cell}>{format_pct(tot_d0)}</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]:
                html += f"<td>{format_pct(tot_m['D0 ROAS'])}</td>"
            elif mc == "CPI ($)":
                html += f"<td>${tot_m['CPI']:.2f}</td>"
            elif mc == "首日付费率 (%)":
                html += f"<td>{format_pct(tot_m['首日付费率'])}</td>"
            elif mc == "CPP ($)":
                html += f"<td>${tot_m['CPP']:.2f}</td>"
            elif mc == "ARPPU ($)":
                html += f"<td>${tot_m['ARPPU']:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]:
                html += f"<td>{format_pct(tot_m['新注册用户占比'])}</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]:
                html += f"<td>{format_pct(tot_m['付费用户新注册用户占比'])}</td>"
            elif mc == "首日激活用户数":
                html += f"<td>{int(tot_m['首日激活用户数']):,}</td>"
            elif mc == "首日付费用户数":
                html += f"<td>{int(tot_m['首日付费用户数']):,}</td>"
        html += "</tr>"

    rendered_dims_counts = [0] * len(dims)
    yellow_dims = ["剧目/书籍语言", "优化师", "剧目/书籍名称", "剧目/书籍短ID", "投放平台", "账户时区", "操作系统 (广告)", "素材名称", "推广目标"]

    for row_idx, r in enumerate(flat_rows):
        html += "<tr>"
        for d_idx, dim_info in enumerate(r["dims"]):
            if rendered_dims_counts[d_idx] == 0:
                rs = dim_info["rowspan"]
                val = dim_info["val"]
                cell_class = 'class="col-highlight-yellow"' if dims[d_idx] in yellow_dims else ""
                html += f'<td rowspan="{rs}" {cell_class}>{val}</td>' if rs > 1 else f"<td {cell_class}>{val}</td>"
                rendered_dims_counts[d_idx] = rs - 1
            else:
                rendered_dims_counts[d_idx] -= 1

        m = r["metrics"]
        d0_val = m["达标率"]
        d0_cell = 'class="detail-d0-85"' if d0_val >= 0.85 else ('class="detail-d0-80"' if d0_val >= 0.80 else '')
        cost_ratio = (m["花费"] / total_cost * 100) if total_cost > 0 else 0

        for mc in metric_cols:
            if mc == "花费 ($)":
                html += f"<td>${m['花费']:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]:
                html += f"<td>{cost_ratio:.2f}%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]:
                html += f"<td {d0_cell}>{format_pct(d0_val)}</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]:
                html += f"<td>{format_pct(m['D0 ROAS'])}</td>"
            elif mc == "CPI ($)":
                html += f"<td>${m['CPI']:.2f}</td>"
            elif mc == "首日付费率 (%)":
                html += f"<td>{format_pct(m['首日付费率'])}</td>"
            elif mc == "CPP ($)":
                html += f"<td>${m['CPP']:.2f}</td>"
            elif mc == "ARPPU ($)":
                html += f"<td>${m['ARPPU']:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]:
                html += f"<td>{format_pct(m['新注册用户占比'])}</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]:
                html += f"<td>{format_pct(m['付费用户新注册用户占比'])}</td>"
            elif mc == "首日激活用户数":
                html += f"<td>{int(m['首日激活用户数']):,}</td>"
            elif mc == "首日付费用户数":
                html += f"<td>{int(m['首日付费用户数']):,}</td>"
            elif mc == "手动出价的金额/系数":
                bid_val = df_data[(df_data[dims[-1]] == r["dims"][-1]["val"])]["手动出价的金额/系数"].iloc[0] if "手动出价的金额/系数" in df_data.columns else "-"
                html += f"<td>{bid_val}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    return html


# -----------------------------------------------------------------------------
# 2. Excel 文件读取与数据预处理
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader("📂 请上传数据底表 (.xlsx / .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    df = load_and_preprocess_excel(file_bytes)

    all_dates = sorted(df["日期_str"].unique(), reverse=True)
    latest_date = all_dates[0] if len(all_dates) > 0 else ""

    options = get_cached_filter_options(df)
    min_date = df["日期_dt"].min().date()
    max_date = df["日期_dt"].max().date()
    first_launch_date = df["首次投放的日期"].min() if "首次投放的日期" in df.columns else latest_date

    with st.sidebar:
        st.header("🎛️ 多维切片筛选")

        with st.form("sidebar_filter_form"):
            submit_btn = st.form_submit_button("🚀 一键应用筛选", use_container_width=True)
            st.markdown("---")

            sel_dramas = st.multiselect("🔍 搜索/筛选 推广剧目/书籍短ID", options=options["dramas"], default=options["dramas"])
            st.markdown("##### 📅 投放时间筛选")
            st.caption(f"💡 首投起始日: {first_launch_date}")
            date_range = st.date_input("选择时间范围", value=[min_date, max_date])

            st.markdown("---")
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

            enable_read_mode = st.toggle("💡 开启/关闭阅读模式 (十字高亮)", value=True)

    start_d = date_range[0] if len(date_range) > 0 else min_date
    end_d = date_range[1] if len(date_range) > 1 else max_date

    df_filtered = apply_fast_filter(
        df, sel_dramas, sel_accounts, sel_opts, sel_langs, sel_copyrights,
        sel_comic_types, sel_platforms, sel_targets, sel_opt_modes, sel_pixels,
        sel_types, start_d, end_d
    )

    if enable_read_mode:
        st.markdown(
            """
        <script>
        (function() {
            function initCrosshair() {
                const tables = document.querySelectorAll('table.dataframe-table');
                tables.forEach(table => {
                    if (table.dataset.crosshairInjected) return;
                    table.dataset.crosshairInjected = "true";

                    table.addEventListener('mouseover', (e) => {
                        const td = e.target.closest('td, th');
                        if (!td) return;

                        table.querySelectorAll('.crosshair-row, .crosshair-col, .crosshair-focus')
                             .forEach(el => el.classList.remove('crosshair-row', 'crosshair-col', 'crosshair-focus'));

                        const tr = td.parentElement;
                        const colIndex = td.cellIndex;

                        td.classList.add('crosshair-focus');
                        tr.querySelectorAll('td, th').forEach(c => c.classList.add('crosshair-row'));

                        table.querySelectorAll('tr').forEach(row => {
                            if (row.cells[colIndex]) {
                                row.cells[colIndex].classList.add('crosshair-col');
                            }
                        });
                    });

                    table.addEventListener('mouseleave', () => {
                        table.querySelectorAll('.crosshair-row, .crosshair-col, .crosshair-focus')
                             .forEach(el => el.classList.remove('crosshair-row', 'crosshair-col', 'crosshair-focus'));
                    });
                });
            }
            setTimeout(initCrosshair, 800);
            document.addEventListener('DOMContentLoaded', initCrosshair);
        })();
        </script>
        """,
            unsafe_allow_html=True,
        )

    # =========================================================================
    # 顶部 6 大 Sheet 选项卡布局视图 (st.tabs)
    # =========================================================================
    sheet1, sheet2, sheet3, sheet4, sheet5, sheet6 = st.tabs(
        [
            "📊 第一部分：整体数据看板",
            "📌 第二部分：多维度分析",
            "📈 第三部分：投放维度拆分",
            "⚙️ 第四部分：优化方式&出价深度拆分",
            "🎯 第五部分：单剧目深度透视",
            "🎬 第六部分：素材分析",
        ]
    )

    # -------------------------------------------------------------------------
    # Sheet 1: 第一部分：整体数据看板
    # -------------------------------------------------------------------------
    with sheet1:
        part1_tab1, part1_tab2 = st.tabs(["📊 总体全量", "📈 每日趋势"])

        with part1_tab1:
            st.subheader("📊 板块一：数据实况表 (总体全量汇总)")
            if df_filtered.empty:
                st.warning("筛选条件下无可展现的数据。")
            else:
                p1_dims = ["剧目/书籍语言", "优化师"]
                p1_metrics = ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
                st.markdown(generate_pivot_html(df_filtered, p1_dims, p1_metrics, total_title="大盘总体 (全量汇总)"), unsafe_allow_html=True)

        with part1_tab2:
            st.subheader("📈 每日数据趋势情况表")
            if not df_filtered.empty:
                p2_dims = ["剧目/书籍语言", "优化师", "日期_str"]
                p2_metrics = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
                st.markdown(generate_pivot_html(df_filtered, p2_dims, p2_metrics, total_title="全期汇总"), unsafe_allow_html=True)

        st.subheader("📊 板块二：达标剧目 & 需重点优化")
        if not df_filtered.empty:
            qualified_shows = []
            for s_name, s_group in df_filtered.groupby("剧目/书籍名称"):
                s_m = compute_metrics(s_group)
                if s_m["花费"] > 100 and s_m["达标率"] >= 0.80:
                    qualified_shows.append(s_name)

            df_qualified = df_filtered[df_filtered["剧目/书籍名称"].isin(qualified_shows)].copy()

            if df_qualified.empty:
                st.info("当前筛选条件下，暂无同时满足【花费 > $100 且 D0 达标率 >= 80%】的重点优化剧集。")
            else:
                rank_dims = ["剧目/书籍语言", "投放平台", "剧目/书籍名称", "剧目/书籍短ID"]
                rank_metrics = ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
                st.markdown('##### 📌 表格 4: 达标优化剧集 <span style="color:red; font-weight:bold;">(优化师重点优化剧目)</span>', unsafe_allow_html=True)
                st.markdown(generate_pivot_html(df_qualified, rank_dims, rank_metrics, total_title="重点优化剧集大盘汇总"), unsafe_allow_html=True)

        st.subheader("🚨 板块三：持续低效/衰退剧目 & 需预警止损")
        if not df_filtered.empty:
            warning_rows = []

            for s_id, s_group in df_filtered.groupby("剧目/书籍短ID"):
                s_daily = s_group.groupby("日期_str").apply(compute_metrics).apply(pd.Series).reset_index().sort_values(by="日期_str", ascending=True)

                if len(s_daily) >= 2:
                    rates = s_daily["达标率"].tolist()
                    s_lang = s_group["剧目/书籍语言"].iloc[0]
                    s_name = s_group["剧目/书籍名称"].iloc[0]
                    tot_cost = s_group["花费"].sum()
                    tot_m = compute_metrics(s_group)

                    is_warn_a = any(rates[i] < 0.60 and rates[i + 1] < 0.60 for i in range(len(rates) - 1))
                    is_warn_b = len(rates) >= 3 and any(rates[i] < 0.70 and rates[i + 1] < 0.70 and rates[i + 2] < 0.70 for i in range(len(rates) - 2))
                    is_warn_c = len(rates) >= 4 and any(rates[i] < 0.76 and rates[i + 1] < 0.76 and rates[i + 2] < 0.76 and rates[i + 3] < 0.76 for i in range(len(rates) - 3))

                    if is_warn_a or is_warn_b or is_warn_c:
                        reasons = []
                        if is_warn_a: reasons.append('<span class="badge-warn-a">🚨 连续2天达标率<60%</span>')
                        if is_warn_b: reasons.append('<span class="badge-warn-b">⚠️ 连续3天达标率<70%</span>')
                        if is_warn_c: reasons.append('<span class="badge-warn-c">🔻 连续4天达标率<76%</span>')

                        warning_rows.append({"剧目/书籍语言": s_lang, "剧目/书籍名称": s_name, "剧目/书籍短ID": s_id, "花费 ($)": tot_cost, "D0 达标率 (%)": tot_m["达标率"], "预警原因/触发条件": " ".join(reasons)})

            if warning_rows:
                df_warn_res = pd.DataFrame(warning_rows).sort_values(by="花费 ($)", ascending=False)
                html_warn = """<table class="dataframe-table"><thead><tr><th>剧目/书籍语言</th><th>剧目/书籍名称</th><th>剧目/书籍短ID</th><th>花费 ($)</th><th>D0 达标率 (%)</th><th>预警原因/触发条件</th></tr></thead><tbody>"""
                for _, r in df_warn_res.iterrows():
                    html_warn += f"""<tr><td>{r['剧目/书籍语言']}</td><td>{r['剧目/书籍名称']}</td><td>{r['剧目/书籍短ID']}</td><td>${r['花费 ($)']:,.2f}</td><td style="color:red; font-weight:bold;">{format_pct(r['D0 达标率 (%)'])}</td><td style="text-align:left; padding-left:15px;">{r['预警原因/触发条件']}</td></tr>"""
                html_warn += "</tbody></table>"
                st.markdown(html_warn, unsafe_allow_html=True)
            else:
                st.success("🎉 当前全盘表现优异，未发现任何衰退预警止损剧目！")

    # -------------------------------------------------------------------------
    # Sheet 2: 第二部分：多维度分析
    # -------------------------------------------------------------------------
    with sheet2:
        st.subheader("📊 表格 1: 投放平台 × 账户 成效透视明细表")
        if not df_filtered.empty:
            acc_metrics = ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)"]
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "账户时区", "投放账户名称"], acc_metrics, total_title="大盘总计 / 账户汇总"), unsafe_allow_html=True)

        st.subheader("🎯 表格 2: 投放平台 × Pixel 名称成效透视明细表")
        if not df_filtered.empty:
            pix_metrics = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "CPP ($)", "首日付费率 (%)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "pixel名称", "投放平台", "推广目标", "优化方式"], pix_metrics, total_title="大盘总计 / Pixel汇总"), unsafe_allow_html=True)

        st.subheader("🚀 表格 3: 投放平台 × 推广目标 × 账户名称 成效透视明细表")
        if not df_filtered.empty:
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "推广目标", "投放账户名称"], acc_metrics, total_title="大盘总计 / 目标账户汇总"), unsafe_allow_html=True)

        st.subheader("🎭 表格 4: 投放平台 × 剧目类型成效透视明细表")
        if not df_filtered.empty:
            type_col_name = "漫剧类型" if "漫剧类型" in df_filtered.columns else "书籍类型"
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "版权来源", type_col_name], ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="大盘总计 / 类型汇总"), unsafe_allow_html=True)

        st.subheader("💻 表格 5: 投放平台 × OS端透视明细表")
        if not df_filtered.empty:
            st.markdown(generate_pivot_html(df_filtered, ["剧目/书籍语言", "投放平台", "操作系统 (广告)", "操作系统 (用户)"], ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="大盘总计 / OS汇总"), unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Sheet 3: 第三部分：投放维度拆分
    # -------------------------------------------------------------------------
    with sheet3:
        part2_tab1, part2_tab2 = st.tabs(["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"])

        def render_part2_content(df_p2_working, is_daily_mode):
            time_suffix = f"(当日明细 - {latest_date})" if is_daily_mode else "(全量汇总)"
            if df_p2_working.empty:
                st.warning("无可展示的数据")
                return

            type_col_name = "漫剧类型" if "漫剧类型" in df_p2_working.columns else "书籍类型"
            common_metrics = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]

            st.subheader(f"📊 表格 1：具体投放剧目类型情况拆分表 {time_suffix}")
            st.markdown(generate_pivot_html(df_p2_working, ["剧目/书籍语言", "版权来源", type_col_name], common_metrics, total_title="类型拆分大盘汇总"), unsafe_allow_html=True)

            st.subheader(f"🎬 表格 2：具体投放剧目情况表 {time_suffix}")
            t2_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "版权来源"]
            if "剧目/书籍中文名称" in df_p2_working.columns: t2_dims.insert(2, "剧目/书籍中文名称")
            st.markdown(generate_pivot_html(df_p2_working, t2_dims, common_metrics, total_title="剧目成效大盘汇总"), unsafe_allow_html=True)

            st.subheader(f"📈 表格 3: 具体投放剧目每日数据情况表 {time_suffix}")
            t3_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "日期_str"]
            if "剧目/书籍中文名称" in df_p2_working.columns: t3_dims.insert(2, "剧目/书籍中文名称")
            st.markdown(generate_pivot_html(df_p2_working, t3_dims, common_metrics, total_title="剧目每日趋势汇总"), unsafe_allow_html=True)

        with part2_tab1: render_part2_content(df_filtered[df_filtered["日期_str"] == latest_date].copy(), is_daily_mode=True)
        with part2_tab2: render_part2_content(df_filtered, is_daily_mode=False)

    # -------------------------------------------------------------------------
    # Sheet 4: 第四部分：优化方式&出价深度拆分
    # -------------------------------------------------------------------------
    with sheet4:
        part3_tab1, part3_tab2 = st.tabs(["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"])

        def render_part3_content(df_p3_working, is_daily_mode):
            suffix_title = "(当日实况)" if is_daily_mode else "(全量总计)"
            total_title = f"大盘总计 / 当日实况汇总" if is_daily_mode else f"大盘总计 / 全量总计汇总"

            st.subheader("📊 板块 1: 渠道维度情况 (FB / TT)")
            if not df_p3_working.empty:
                st.markdown(generate_pivot_html(df_p3_working, ["剧目/书籍语言", "优化师", "投放平台"], ["花费 ($)", "花费占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=total_title), unsafe_allow_html=True)

            st.subheader("⚙️ 板块 2: 优化方式维度情况")
            if not df_p3_working.empty:
                o1_dims = ["剧目/书籍语言", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数"]
                st.markdown(generate_pivot_html(df_p3_working, o1_dims, ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=total_title), unsafe_allow_html=True)

            st.markdown(f'##### 📌 表格 2: 极效高 ROAS 优化方式 & 出价透视表 (花费>$100 且 达标率>=80%) <span style="color:red; font-weight:bold; font-size:14px;">(💡 广告方向优化策略，可查看此表！)</span> {suffix_title}', unsafe_allow_html=True)
            if not df_p3_working.empty:
                tot_p3_cost = df_p3_working["花费"].sum()
                all_strategy_list = []

                for (d_name, l_v, p_v, t_v, m_v, b_v), sub_g in df_p3_working.groupby(["剧目/书籍名称", "剧目/书籍语言", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数"]):
                    m_sub = compute_metrics(sub_g)
                    c_ratio = (m_sub["花费"] / tot_p3_cost * 100) if tot_p3_cost > 0 else 0
                    all_strategy_list.append({"key": (l_v, p_v, t_v, m_v, b_v), "剧目/书籍名称": d_name, "剧目/书籍语言": l_v, "投放平台": p_v, "推广目标": t_v, "优化方式": m_v, "手动出价的金额/系数": b_v, "花费占比": c_ratio, "达标率": m_sub["达标率"], "花费": m_sub["花费"]})

                high_roas_keys = [item["key"] for item in all_strategy_list if item["花费"] > 100 and item["达标率"] >= 0.80]

                if high_roas_keys:
                    df_high_roas = df_p3_working[df_p3_working.set_index(["剧目/书籍语言", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数"]).index.isin(high_roas_keys)].copy()
                    st.markdown(generate_pivot_html(df_high_roas, ["剧目/书籍语言", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数"], ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="高ROAS策略大盘汇总"), unsafe_allow_html=True)
                else:
                    st.info("当前条件下，暂无同时满足【花费 > $100 且 达标率 >= 80%】的高 ROAS 出价策略。")

                st.markdown("##### 💡 广告方向策略提效看板")
                df_all_summary = pd.DataFrame(all_strategy_list)

                with st.container(border=True):
                    card1_tab1, card1_tab2 = st.tabs(["✅ 达标方向拓展策略 (达标率>80%, 花费占比>5%)", "🚨 不达标方向预算缩减 (达标率<70%, 花费占比>5%)"])
                    with card1_tab1:
                        df_good = df_all_summary[(df_all_summary["达标率"] > 0.80) & (df_all_summary["花费占比"] > 5.0)].sort_values(by="花费占比", ascending=False)
                        if not df_good.empty:
                            st.dataframe(df_good[["剧目/书籍名称", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数", "花费占比"]], use_container_width=True, hide_index=True, column_config={"花费占比": st.column_config.ProgressColumn("花费占比 (%)", help="拓展消耗占比进度条", format="%.2f%%", min_value=0, max_value=100)})
                        else: st.info("暂无符合【达标率>80% 且 花费占比>5%】的拓展方向策略。")

                    with card1_tab2:
                        df_bad = df_all_summary[(df_all_summary["达标率"] < 0.70) & (df_all_summary["花费占比"] > 5.0)].sort_values(by="花费占比", ascending=False)
                        if not df_bad.empty:
                            st.markdown('<span style="color:red; font-weight:bold;">🚨 以下方向达标率低于 70% 且消耗超 5%，建议立即下调预算或止损！</span>', unsafe_allow_html=True)
                            st.dataframe(df_bad[["剧目/书籍名称", "投放平台", "推广目标", "优化方式", "手动出价的金额/系数", "花费占比"]], use_container_width=True, hide_index=True, column_config={"花费占比": st.column_config.ProgressColumn("花费占比 (%)", help="低效浪费占比进度条", format="%.2f%%", min_value=0, max_value=100)})
                        else: st.success("🎉 未发现【达标率<70% 且 花费占比>5%】的低效浪费方向！")

            st.subheader("🎯 板块 3: 投放平台 & Pixel 名称数据维度情况")
            if not df_p3_working.empty:
                st.markdown(generate_pivot_html(df_p3_working, ["剧目/书籍语言", "投放平台", "推广目标", "优化方式", "pixel名称"], ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "CPP ($)", "首日付费率 (%)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=total_title), unsafe_allow_html=True)

            st.subheader("🌐 板块 4: 不同语种的账户时区数据透视表")
            if not df_p3_working.empty:
                for lang in df_p3_working["剧目/书籍语言"].unique():
                    st.markdown(generate_pivot_html(df_p3_working[df_p3_working["剧目/书籍语言"] == lang], ["账户时区"], ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=f"{lang} 小计"), unsafe_allow_html=True)

            st.subheader("⏱️ 板块 5: 当日新老广告效果对比")
            if not df_p3_working.empty:
                df_p5 = df_p3_working.copy()
                df_p5["新老广告分类"] = df_p5.apply(lambda r: "当日新建广告" if (0 <= (pd.to_datetime(r["日期_str"]) - pd.to_datetime(r["广告创建时间"])).days <= 1) else "老广告", axis=1) if "广告创建时间" in df_p5.columns else "老广告"
                for lang in ["繁体中文", "韩语"]:
                    group_l = df_p5[df_p5["剧目/书籍语言"] == lang]
                    if not group_l.empty:
                        st.markdown(generate_pivot_html(group_l, ["日期_str", "新老广告分类"], ["花费 ($)", "消耗占比 (%)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=f"{lang} 小计"), unsafe_allow_html=True)

            st.subheader("💻 板块 6: 不同操作系统数据维度情况")
            if not df_p3_working.empty:
                os2_dims = ["剧目/书籍语言", "投放平台", "pixel名称", "操作系统 (广告)", "操作系统 (用户)", "推广目标", "优化方式", "手动出价的金额/系数"]
                st.markdown(generate_pivot_html(df_p3_working, os2_dims, ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title=total_title), unsafe_allow_html=True)

                st.markdown(f'##### 📌 表格 2: 达标 操作系统 & 广告优化方式数据汇总 <span style="color:red; font-weight:bold; font-size:14px;">(💡 最全Pixel+优化策略，可查看此表！)</span> {suffix_title}', unsafe_allow_html=True)
                qualified_os_items = []
                for keys_combo, sub_g in df_p3_working.groupby(os2_dims):
                    m_sub = compute_metrics(sub_g)
                    if m_sub["花费"] > 100 and m_sub["达标率"] >= 0.80:
                        qualified_os_items.append(keys_combo)

                if qualified_os_items:
                    df_qualified_os = df_p3_working[df_p3_working.set_index(os2_dims).index.isin(qualified_os_items)].copy()
                    st.markdown(generate_pivot_html(df_qualified_os, os2_dims, ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"], total_title="达标OS与策略大盘汇总"), unsafe_allow_html=True)

        with part3_tab1: render_part3_content(df_filtered[df_filtered["日期_str"] == latest_date].copy(), is_daily_mode=True)
        with part3_tab2: render_part3_content(df_filtered, is_daily_mode=False)

    # -------------------------------------------------------------------------
    # Sheet 5: 第五部分：单剧目深度透视
    # -------------------------------------------------------------------------
    with sheet5:
        if len(sel_dramas) == 1:
            single_drama_id = sel_dramas[0]
            df_single = df[df["剧目/书籍短ID"].astype(str) == single_drama_id].copy()

            if not df_single.empty:
                s_name = df_single["剧目/书籍名称"].iloc[0]
                s_zh = df_single.get("剧目/书籍中文名称", pd.Series(["-"])).iloc[0]
                s_first_date = df_single["首次投放_str"].min()

                st.header(f"🎯 {s_name}-{s_zh}-{single_drama_id} 深度透视 (首投时间：{s_first_date})")
                p5_metrics = ["花费 ($)", "达标率 (%)", "D0 ROAS (%)", "CPI ($)", "首日付费率 (%)", "CPP ($)", "ARPPU ($)", "新注册用户占比 (%)", "付费用户新注册用户占比 (%)"]

                st.subheader("📊 表格 1: 投放平台数据拆分对比")
                st.markdown(generate_pivot_html(df_single, ["投放平台"], p5_metrics, total_title="剧目平台大盘汇总"), unsafe_allow_html=True)

                st.subheader("👨‍💻 表格 2: 优化师数据对比")
                st.markdown(generate_pivot_html(df_single, ["优化师"], p5_metrics, total_title="剧目优化师大盘汇总"), unsafe_allow_html=True)

                st.subheader("🎯 表格 3: 投放平台 × Pixel名称 成效透视对比")
                st.markdown(generate_pivot_html(df_single, ["投放平台", "pixel名称"], p5_metrics, total_title="剧目Pixel大盘汇总"), unsafe_allow_html=True)

                st.subheader("🚀 表格 4: 投放平台 × 推广目标 拆分对比")
                st.markdown(generate_pivot_html(df_single, ["投放平台", "推广目标"], p5_metrics, total_title="剧目目标大盘汇总"), unsafe_allow_html=True)

                st.subheader("🔬 表格 5: 细分全量深度透视（平台×广告类型×pixel×目标×方式×出价）")
                p5_5_dims = ["投放平台", "广告类型", "pixel名称", "操作系统 (广告)", "推广目标", "优化方式", "手动出价的金额/系数"]
                st.markdown(generate_pivot_html(df_single, p5_5_dims, p5_metrics, total_title="剧目全量细分汇总"), unsafe_allow_html=True)
        else:
            st.info("💡 提示：请在左侧侧边栏【🔍 搜索/筛选 推广剧目/书籍短ID】中**仅选择单个剧目**，即可在此处查阅该剧目的单剧目深度透视分析！")

    # -------------------------------------------------------------------------
    # Sheet 6: 第六部分：素材分析 (百分比格式化直刷修复)
    # -------------------------------------------------------------------------
    with sheet6:
        mat_tab1, mat_tab2 = st.tabs(["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"])

        def render_material_analysis(df_mat_working, is_daily_mode):
            time_suffix = f"(当日明细 - {latest_date})" if is_daily_mode else "(全量汇总)"
            if df_mat_working.empty:
                st.warning("当前没有可分析的素材数据")
                return

            qualified_mat_keys = []
            mat_group_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "素材名称"]

            for keys_combo, sub_mat in df_mat_working.groupby(mat_group_dims):
                m_mat = compute_metrics(sub_mat)
                if m_mat["达标率"] >= 0.80 and m_mat["花费"] > 0:
                    qualified_mat_keys.append(keys_combo)

            if qualified_mat_keys:
                df_mat_qualified = df_mat_working[df_mat_working.set_index(mat_group_dims).index.isin(qualified_mat_keys)].copy()

                copy_text_lines = []
                for (s_lang, s_name, s_id), drama_mat_sub in df_mat_qualified.groupby(["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID"]):
                    zh_title = drama_mat_sub.get("剧目/书籍中文名称", pd.Series(["-"])).iloc[0]
                    copy_text_lines.append(f"{s_name} - {zh_title} - {s_id}\n达标素材清单 (一行一个):")
                    for m_n in sorted(drama_mat_sub["素材名称"].dropna().unique().tolist()):
                        copy_text_lines.append(m_n)
                    copy_text_lines.append("")

                st.markdown('<div class="material-copy-card"><div class="material-header-title">📋 复制达标素材名称列表 (带剧目信息抬头)</div>', unsafe_allow_html=True)
                with st.expander("点击展开全量达标素材清单"):
                    st.text_area("全量达标素材文本框:", value="\n".join(copy_text_lines), height=220)
                st.markdown("</div>", unsafe_allow_html=True)

                st.subheader(f"📊 表格 1: 具体投放剧目素材拆分表 (D0达标率>=80%) {time_suffix}")

                mat_rows = []
                for keys_combo, drama_mat_sub in df_mat_qualified.groupby(mat_group_dims):
                    m_res = compute_metrics(drama_mat_sub)
                    s_zh = drama_mat_sub.get("剧目/书籍中文名称", pd.Series(["-"])).iloc[0]
                    
                    mat_rows.append({
                        "剧目/书籍语言": keys_combo[0],
                        "剧目/书籍名称": keys_combo[1],
                        "中文推广剧目": s_zh,
                        "剧目ID": keys_combo[2],
                        "素材名称": keys_combo[3],
                        "花费 ($)": f"${m_res['花费']:,.2f}",
                        "D0 达标率 (%)": format_pct(m_res["达标率"]),
                        "首日付费用户数": int(m_res["首日付费用户数"]),
                        "D0 ROAS (%)": format_pct(m_res["D0 ROAS"]),
                        "CPI ($)": f"${m_res['CPI']:.2f}",
                        "首日付费率 (%)": format_pct(m_res["首日付费率"]),
                        "首日激活用户数": int(m_res["首日激活用户数"]),
                        "新注册用户占比 (%)": format_pct(m_res["新注册用户占比"]),
                        "付费用户新注册用户占比 (%)": format_pct(m_res["付费用户新注册用户占比"])
                    })

                df_mat_display = pd.DataFrame(mat_rows).sort_values(by="花费 ($)", ascending=False)
                st.dataframe(df_mat_display, use_container_width=True, hide_index=True)
            else:
                st.info("当前时间窗口内，暂无【D0 达标率 >= 80%】的达标素材。")

        with mat_tab1: render_material_analysis(df_filtered[df_filtered["日期_str"] == latest_date].copy(), is_daily_mode=True)
        with mat_tab2: render_material_analysis(df_filtered, is_daily_mode=False)
