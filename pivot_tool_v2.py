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
    h1 {
        font-size: 22px !important;
        margin-bottom: 10px !important;
    }
    h2, .stHeader {
        font-size: 18px !important;
        margin-top: 14px !important;
        margin-bottom: 8px !important;
    }
    h3, .stSubheader {
        font-size: 16px !important;
        margin-top: 10px !important;
        margin-bottom: 6px !important;
    }
    h4, h5 {
        font-size: 14px !important;
    }

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
    .row-total {
        background-color: #FFE066 !important; /* 大盘总计/全局汇总 暖黄色 */
        font-weight: bold;
    }
    .row-subtotal {
        background-color: #FFF9C4 !important; /* 小计行 浅黄色 */
        font-weight: bold;
    }
    .col-highlight-yellow {
        background-color: #FFF9C4 !important; /* 关键维度列 浅黄色 */
    }

    /* 达标率（>80% / >85%）文本与背景样式分类定义 */
    .summary-d0-85 {
        color: #28A745 !important;
        font-weight: bold !important;
        text-decoration: underline !important;
    }
    .summary-d0-80 {
        color: #28A745 !important;
        font-weight: bold !important;
    }

    .detail-d0-85 {
        background-color: #D4EDDA !important; /* 淡绿色背景 */
        color: #28A745 !important;
        font-weight: bold !important;
    }
    .detail-d0-80 {
        color: #28A745 !important;
        font-weight: bold !important;
    }

    /* 侧边栏多选标签 (Tags) 红色卡片视觉重构 */
    [data-testid="stSidebar"] [data-baseweb="select"] {
        max-height: 160px;
        overflow-y: auto;
    }
    [data-testid="stSidebar"] [data-baseweb="tag"] {
        background-color: #FF4D4F !important;
        color: #FFFFFF !important;
        margin: 2px;
        border-radius: 4px;
        font-weight: 500;
        font-size: 12px;
    }
    [data-testid="stSidebar"] [data-baseweb="tag"] span {
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] [data-baseweb="tag"] svg {
        fill: #FFFFFF !important;
    }

    /* 预警标签颜色 */
    .badge-warn-a {
        background-color: #FF4D4F;
        color: white;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 12px;
    }
    .badge-warn-b {
        background-color: #FA8C16;
        color: white;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 12px;
    }
    .badge-warn-c {
        background-color: #FADB14;
        color: #333;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 12px;
    }
    
    /* 策略总结卡片容器美化 */
    .strategy-box {
        background-color: #F6FFED;
        border: 1px solid #B7EB8F;
        border-radius: 6px;
        padding: 12px 16px;
        margin-top: 10px;
        margin-bottom: 20px;
        font-size: 13px;
        line-height: 1.6;
    }
    .strategy-box-title {
        font-weight: bold;
        color: #52C41A;
        font-size: 14px;
        margin-bottom: 6px;
    }

    /* 素材复制卡片UI美化 */
    .material-copy-card {
        background-color: #F0F5FF;
        border: 1px solid #ADC6FF;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 16px;
    }
    .material-header-title {
        font-size: 16px;
        font-weight: bold;
        color: #1D39C4;
        margin-bottom: 8px;
    }
    
    /* 顶部 6 大 Sheet 选项卡样式强化 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: bold;
        padding: 8px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🎬 短剧/漫剧综合数据看板")

# -----------------------------------------------------------------------------
# 核心缓存与帮助函数定义
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)
def get_cached_filter_options(df):
    target_col = "推广目标（属性）" if "推广目标（属性）" in df.columns else "推广目标"
    type_col = "书籍类型" if "书籍类型" in df.columns else "漫剧类型"
    
    return {
        "dramas": sorted(
            df["剧目/书籍短ID"].dropna().astype(str).unique().tolist()
        ),
        "accounts": sorted(
            df["投放账户名称"].dropna().astype(str).unique().tolist()
        ),
        "opts": sorted(
            [o for o in df["优化师"].dropna().unique() if o != "-"]
        ),
        "langs": sorted(
            df["剧目/书籍语言"].dropna().unique().tolist()
        ) if "剧目/书籍语言" in df.columns else [],
        "copyrights": sorted(
            df["版权来源"].dropna().unique().tolist()
        ) if "版权来源" in df.columns else [],
        "comic_types": sorted(
            df["漫剧类型"].dropna().unique().tolist()
        ) if "漫剧类型" in df.columns else [],
        "platforms": sorted(df["投放平台"].dropna().unique().tolist()),
        "targets": sorted(
            df[target_col].dropna().unique().tolist()
        ) if target_col in df.columns else [],
        "opt_modes": sorted(
            df["优化方式"].dropna().unique().tolist()
        ) if "优化方式" in df.columns else [],
        "pixels": sorted(
            df["pixel名称"].dropna().astype(str).unique().tolist()
        ),
        "types": sorted(
            df[type_col].dropna().unique().tolist()
        ) if type_col in df.columns else [],
    }


@st.cache_data
def apply_fast_filter(
    df,
    sel_dramas,
    sel_accounts,
    sel_opts,
    sel_langs,
    sel_copyrights,
    sel_comic_types,
    sel_platforms,
    sel_targets,
    sel_opt_modes,
    sel_pixels,
    sel_types,
    start_date,
    end_date,
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


def clean_num(val):
    try:
        if pd.isna(val) or val == "-" or val == "":
            return 0.0
        return float(val)
    except:
        return 0.0


def format_pct(val):
    """百分比格式化修复函数: 自动转化 0.83 -> 83.00%"""
    if pd.isna(val) or val == "-" or val == "":
        return "0.00%"
    try:
        num = float(val)
        if num <= 1.0 and num >= 0.0:
            return f"{num * 100:.2f}%"
        return f"{num:.2f}%"
    except:
        return f"{val}"


def compute_metrics(sub_df):
    """通用透视与计算规范（精准加权公式）"""
    c_cost = sub_df["花费"].sum()

    sum_d0_target = sub_df["D0目标收入"].sum() if "D0目标收入" in sub_df.columns else 0
    sum_d0_rev = sub_df["d0_revenue"].sum() if "d0_revenue" in sub_df.columns else 0
    if sum_d0_target > 0:
        d0_reach = sum_d0_rev / sum_d0_target
    else:
        d0_reach = (sub_df["D0达标率"] * sub_df["花费"]).sum() / (c_cost or 1)

    d0_roas = (sub_df["d0_roas"] * sub_df["花费"]).sum() / (c_cost or 1)

    sum_active = (
        sub_df["首日激活用户数"].sum()
        if "首日激活用户数" in sub_df.columns
        else 0
    )
    if sum_active > 0:
        cpi = c_cost / sum_active
    else:
        cpi = (sub_df["CPI"] * sub_df["花费"]).sum() / (c_cost or 1)

    sum_pay_users = (
        sub_df["首日付费用户数"].sum()
        if "首日付费用户数" in sub_df.columns
        else 0
    )
    if sum_pay_users > 0:
        cpp = c_cost / sum_pay_users
    else:
        cpp = (sub_df["首日付费用户成本"] * sub_df["花费"]).sum() / (
            c_cost or 1
        )

    if sum_active > 0 and sum_pay_users > 0:
        pay_rate = sum_pay_users / sum_active
    else:
        pay_rate = (sub_df["首日付费率"] * sub_df["花费"]).sum() / (c_cost or 1)

    arppu = (
        sub_df["Arppu- 付费用户人均付费金额"] * sub_df["花费"]
    ).sum() / (c_cost or 1)

    if sum_active > 0:
        new_reg_count = (
            sub_df["首日激活用户数"] * sub_df["新注册用户占比"]
        ).sum()
        reg_ratio = new_reg_count / sum_active
    else:
        reg_ratio = (sub_df["新注册用户占比"] * sub_df["花费"]).sum() / (
            c_cost or 1
        )

    if sum_pay_users > 0:
        new_pay_count = (
            sub_df["首日付费用户数"] * sub_df["付费用户新注册用户占比"]
        ).sum()
        pay_reg_ratio = new_pay_count / sum_pay_users
    else:
        pay_reg_ratio = (
            sub_df["付费用户新注册用户占比"] * sub_df["花费"]
        ).sum() / (c_cost or 1)

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
        sorted_keys = (
            grouped["花费"].sum().sort_values(ascending=False).index.tolist()
        )

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
            prefix = current_row_prefix + [
                {"val": n["val"], "rowspan": n["rowspan"]}
            ]
            if "children" in n:
                flat_rows.extend(flatten_tree_to_rows(n["children"], prefix))
            else:
                flat_rows.append(
                    {"dims": prefix, "metrics": n["leaf_data"]}
                )
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
        if tot_m["达标率"] >= 0.85:
            d0_cell = 'class="summary-d0-85"'
        elif tot_m["达标率"] >= 0.80:
            d0_cell = 'class="summary-d0-80"'
        else:
            d0_cell = ""

        html += f'<tr class="row-total"><td colspan="{len(dims)}">{total_title}</td>'
        for mc in metric_cols:
            if mc == "花费 ($)":
                html += f"<td>${total_cost:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]:
                html += "<td>100.00%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]:
                html += f"<td {d0_cell}>{format_pct(tot_m['达标率'])}</td>"
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

    yellow_dims = [
        "剧目/书籍语言",
        "优化师",
        "剧目/书籍名称",
        "剧目/书籍短ID",
        "投放平台",
        "账户时区",
        "操作系统 (广告)",
        "素材名称",
        "推广目标",
    ]

    for row_idx, r in enumerate(flat_rows):
        html += "<tr>"
        for d_idx, dim_info in enumerate(r["dims"]):
            if rendered_dims_counts[d_idx] == 0:
                rs = dim_info["rowspan"]
                val = dim_info["val"]
                cell_class = (
                    'class="col-highlight-yellow"'
                    if dims[d_idx] in yellow_dims
                    else ""
                )
                html += (
                    f'<td rowspan="{rs}" {cell_class}>{val}</td>'
                    if rs > 1
                    else f"<td {cell_class}>{val}</td>"
                )
                rendered_dims_counts[d_idx] = rs - 1
            else:
                rendered_dims_counts[d_idx] -= 1

        m = r["metrics"]
        d0_val = m["达标率"]

        if d0_val >= 0.85:
            d0_cell = 'class="detail-d0-85"'
        elif d0_val >= 0.80:
            d0_cell = 'class="detail-d0-80"'
        else:
            d0_cell = ""

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
                bid_val = (
                    df_data[
                        (df_data[dims[-1]] == r["dims"][-1]["val"])
                    ]["手动出价的金额/系数"]
                    .iloc[0]
                    if "手动出价的金额/系数" in df_data.columns
                    else "-"
                )
                html += f"<td>{bid_val}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    return html


# -----------------------------------------------------------------------------
# 2. Excel 文件读取与数据预处理
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "📂 请上传数据底表 (.xlsx / .xls)", type=["xlsx", "xls"]
)

if uploaded_file is not None:
    raw_df = pd.read_excel(uploaded_file)

    required_defaults = {
        "优化师": "-",
        "花费": 0.0,
        "剧目/书籍语言": "未知语种",
        "剧目/书籍名称": "未知剧目",
        "剧目/书籍中文名称": "-",
        "剧目/书籍短ID": "-",
        "版权来源": "-",
        "pixel名称": "-",
        "优化方式": "-",
        "操作系统 (广告)": "-",
        "操作系统 (用户)": "-",
        "投放平台": "-",
        "推广目标": "-",
        "推广目标（属性）": "-",
        "广告类型": "-",
        "书籍类型": "未知类型",
        "漫剧类型": "未知类型",
        "首次投放的日期": "2020-01-01",
        "投放账户名称": "-",
        "广告创建时间": "2020-01-01",
        "素材名称": "-",
        "D0达标率": 0.0,
        "d0_roas": 0.0,
        "CPI": 0.0,
        "首日付费率": 0.0,
        "首日付费用户成本": 0.0,
        "Arppu- 付费用户人均付费金额": 0.0,
        "新注册用户占比": 0.0,
        "付费用户新注册用户占比": 0.0,
        "D0目标收入": 0.0,
        "d0_revenue": 0.0,
        "首日激活用户数": 0.0,
        "首日付费用户数": 0.0,
    }
    for col_name, default_val in required_defaults.items():
        if col_name not in raw_df.columns:
            raw_df[col_name] = default_val

    df = raw_df[(raw_df["优化师"] != "-") & (raw_df["花费"] > 0)].copy()

    df["日期_dt"] = pd.to_datetime(df["日期"])
    df["日期_str"] = df["日期_dt"].dt.strftime("%Y-%m-%d")
    df["首次投放_str"] = pd.to_datetime(df["首次投放的日期"]).dt.strftime(
        "%Y-%m-%d"
    )

    all_dates = sorted(df["日期_str"].unique(), reverse=True)
    latest_date = all_dates[0] if len(all_dates) > 0 else ""

    num_cols = [
        "花费",
        "D0达标率",
        "d0_roas",
        "CPI",
        "首日付费率",
        "首日付费用户成本",
        "Arppu- 付费用户人均付费金额",
        "新注册用户占比",
        "付费用户新注册用户占比",
        "D0目标收入",
        "d0_revenue",
        "首日激活用户数",
        "首日付费用户数",
    ]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(clean_num)

    def get_timezone(acc_name):
        acc = str(acc_name)
        if "+8-" in acc:
            return "+8 时区账户"
        elif "+0-" in acc:
            return "0 时区账户"
        return "美西账户"

    df["账户时区"] = df["投放账户名称"].apply(get_timezone)

    # -------------------------------------------------------------------------
    # 侧边栏切片筛选面板
    # -------------------------------------------------------------------------
    options = get_cached_filter_options(df)
    min_date = df["日期_dt"].min().date()
    max_date = df["日期_dt"].max().date()
    first_launch_date = df["首次投放的日期"].min() if "首次投放的日期" in df.columns else latest_date

    with st.sidebar:
        st.header("🎛️ 多维切片筛选")

        with st.form("sidebar_filter_form"):
            submit_btn = st.form_submit_button(
                "🚀 一键应用筛选", use_container_width=True
            )

            st.markdown("---")

            sel_dramas = st.multiselect(
                "🔍 搜索/筛选 推广剧目/书籍短ID",
                options=options["dramas"],
                default=options["dramas"],
                placeholder="输入或选择短ID (留空代表全选)",
            )

            st.markdown("##### 📅 投放时间筛选")
            st.caption(f"💡 首投起始日: {first_launch_date}")
            date_range = st.date_input(
                "选择时间范围 (从首投日开始)",
                value=[min_date, max_date],
            )

            st.markdown("---")

            sel_accounts = st.multiselect(
                "筛选 投放账户名称",
                options=options["accounts"],
                default=options["accounts"],
            )
            sel_opts = st.multiselect(
                "筛选 优化师",
                options=options["opts"],
                default=options["opts"],
            )
            sel_langs = st.multiselect(
                "筛选 剧目/书籍语言",
                options=options["langs"],
                default=options["langs"],
            )
            sel_copyrights = st.multiselect(
                "筛选 版权来源",
                options=options["copyrights"],
                default=options["copyrights"],
            )
            sel_comic_types = st.multiselect(
                "筛选 漫剧类型",
                options=options["comic_types"],
                default=options["comic_types"],
            )
            sel_platforms = st.multiselect(
                "筛选 投放平台",
                options=options["platforms"],
                default=options["platforms"],
            )
            sel_targets = st.multiselect(
                "筛选 推广目标",
                options=options["targets"],
                default=options["targets"],
            )
            sel_opt_modes = st.multiselect(
                "筛选 优化方式",
                options=options["opt_modes"],
                default=options["opt_modes"],
            )
            sel_pixels = st.multiselect(
                "筛选 pixel名称",
                options=options["pixels"],
                default=options["pixels"],
            )
            sel_types = st.multiselect(
                "筛选 书籍类型",
                options=options["types"],
                default=options["types"],
            )

            enable_read_mode = st.toggle(
                "💡 开启/关闭阅读模式 (十字高亮)", value=True
            )

    start_d = date_range[0] if len(date_range) > 0 else min_date
    end_d = date_range[1] if len(date_range) > 1 else max_date

    df_filtered = apply_fast_filter(
        df,
        sel_dramas,
        sel_accounts,
        sel_opts,
        sel_langs,
        sel_copyrights,
        sel_comic_types,
        sel_platforms,
        sel_targets,
        sel_opt_modes,
        sel_pixels,
        sel_types,
        start_d,
        end_d,
    )

    # 十字高亮支持
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
                p1_metrics = [
                    "花费 ($)",
                    "消耗占比 (%)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_p1 = generate_pivot_html(
                    df_filtered,
                    p1_dims,
                    p1_metrics,
                    total_title="大盘总体 (全量汇总)",
                )
                st.markdown(html_p1, unsafe_allow_html=True)

        with part1_tab2:
            st.subheader("📈 每日数据趋势情况表")

            if not df_filtered.empty:
                p2_dims = ["剧目/书籍语言", "优化师", "日期_str"]
                p2_metrics = [
                    "花费 ($)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_p2 = generate_pivot_html(
                    df_filtered,
                    p2_dims,
                    p2_metrics,
                    total_title="全期汇总",
                )
                st.markdown(html_p2, unsafe_allow_html=True)

        # 板块二：达标剧目 & 需重点优化
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
                rank_metrics = [
                    "花费 ($)",
                    "消耗占比 (%)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                st.markdown(
                    '##### 📌 表格 4: 达标优化剧集 <span style="color:red; font-weight:bold;">(优化师重点优化剧目)</span>',
                    unsafe_allow_html=True,
                )
                html_rank = generate_pivot_html(
                    df_qualified, rank_dims, rank_metrics, total_title="重点优化剧集大盘汇总"
                )
                st.markdown(html_rank, unsafe_allow_html=True)

        # 板块三：持续低效/衰退剧目 & 需预警止损
        st.subheader("🚨 板块三：持续低效/衰退剧目 & 需预警止损")
        if not df_filtered.empty:
            warning_rows = []

            for s_id, s_group in df_filtered.groupby("剧目/书籍短ID"):
                s_daily = (
                    s_group.groupby("日期_str")
                    .apply(compute_metrics)
                    .apply(pd.Series)
                    .reset_index()
                    .sort_values(by="日期_str", ascending=True)
                )

                if len(s_daily) >= 2:
                    rates = s_daily["达标率"].tolist()
                    s_lang = s_group["剧目/书籍语言"].iloc[0]
                    s_name = s_group["剧目/书籍名称"].iloc[0]
                    tot_cost = s_group["花费"].sum()
                    tot_m = compute_metrics(s_group)

                    is_warn_a = False
                    is_warn_b = False
                    is_warn_c = False

                    # 条件 A: 连续 2 天 D0 达标率 < 60%
                    for i in range(len(rates) - 1):
                        if rates[i] < 0.60 and rates[i + 1] < 0.60:
                            is_warn_a = True
                            break

                    # 条件 B: 连续 3 天 D0 达标率 < 70%
                    if len(rates) >= 3:
                        for i in range(len(rates) - 2):
                            if rates[i] < 0.70 and rates[i + 1] < 0.70 and rates[i + 2] < 0.70:
                                is_warn_b = True
                                break

                    # 条件 C: 连续 4 天 D0 达标率 < 76%
                    if len(rates) >= 4:
                        for i in range(len(rates) - 3):
                            if (
                                rates[i] < 0.76
                                and rates[i + 1] < 0.76
                                and rates[i + 2] < 0.76
                                and rates[i + 3] < 0.76
                            ):
                                is_warn_c = True
                                break

                    if is_warn_a or is_warn_b or is_warn_c:
                        reasons = []
                        if is_warn_a:
                            reasons.append('<span class="badge-warn-a">🚨 连续2天达标率<60%</span>')
                        if is_warn_b:
                            reasons.append('<span class="badge-warn-b">⚠️ 连续3天达标率<70%</span>')
                        if is_warn_c:
                            reasons.append('<span class="badge-warn-c">🔻 连续4天达标率<76%</span>')

                        warning_rows.append(
                            {
                                "剧目/书籍语言": s_lang,
                                "剧目/书籍名称": s_name,
                                "剧目/书籍短ID": s_id,
                                "花费 ($)": tot_cost,
                                "D0 达标率 (%)": tot_m["达标率"],
                                "预警原因/触发条件": " ".join(reasons),
                            }
                        )

            if warning_rows:
                df_warn_res = (
                    pd.DataFrame(warning_rows)
                    .sort_values(by="花费 ($)", ascending=False)
                )

                html_warn = """<table class="dataframe-table">
                    <thead>
                        <tr>
                            <th>剧目/书籍语言</th>
                            <th>剧目/书籍名称</th>
                            <th>剧目/书籍短ID</th>
                            <th>花费 ($)</th>
                            <th>D0 达标率 (%)</th>
                            <th>预警原因/触发条件</th>
                        </tr>
                    </thead>
                    <tbody>"""

                for _, r in df_warn_res.iterrows():
                    d0_v = r["D0 达标率 (%)"]
                    html_warn += f"""
                    <tr>
                        <td>{r['剧目/书籍语言']}</td>
                        <td>{r['剧目/书籍名称']}</td>
                        <td>{r['剧目/书籍短ID']}</td>
                        <td>${r['花费 ($)']:,.2f}</td>
                        <td style="color:red; font-weight:bold;">{format_pct(d0_v)}</td>
                        <td style="text-align:left; padding-left:15px;">{r['预警原因/触发条件']}</td>
                    </tr>"""

                html_warn += "</tbody></table>"
                st.markdown(html_warn, unsafe_allow_html=True)
            else:
                st.success("🎉 当前全盘表现优异，未发现任何衰退预警止损剧目！")

    # -------------------------------------------------------------------------
    # Sheet 2: 第二部分：多维度分析 (包含追加的新增表格 3)
    # -------------------------------------------------------------------------
    with sheet2:
        # 表格 1: 投放平台 × 账户 成效透视明细表
        st.subheader("📊 表格 1: 投放平台 × 账户 成效透视明细表")
        if not df_filtered.empty:
            acc_dims = ["剧目/书籍语言", "投放平台", "账户时区", "投放账户名称"]
            acc_metrics = [
                "花费 ($)",
                "消耗占比 (%)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
            ]
            html_acc = generate_pivot_html(
                df_filtered, acc_dims, acc_metrics, total_title="大盘总计 / 账户汇总"
            )
            st.markdown(html_acc, unsafe_allow_html=True)

        # 表格 2: 投放平台 × Pixel 名称成效透视明细表
        st.subheader("🎯 表格 2: 投放平台 × Pixel 名称成效透视明细表")
        if not df_filtered.empty:
            pixel_dims = ["剧目/书籍语言", "pixel名称", "投放平台", "推广目标", "优化方式"]
            pixel_metrics = [
                "花费 ($)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "CPP ($)",
                "首日付费率 (%)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_pixel = generate_pivot_html(
                df_filtered, pixel_dims, pixel_metrics, total_title="大盘总计 / Pixel汇总"
            )
            st.markdown(html_pixel, unsafe_allow_html=True)

        # 一、 【本次追加新增】表格 3: 投放平台 × 推广目标 × 账户名称 成效透视明细表
        st.subheader("🚀 表格 3: 投放平台 × 推广目标 × 账户名称 成效透视明细表")
        if not df_filtered.empty:
            target_acc_dims = [
                "剧目/书籍语言",
                "投放平台",
                "推广目标",
                "投放账户名称",
            ]
            target_acc_metrics = [
                "花费 ($)",
                "消耗占比 (%)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
            ]
            html_target_acc = generate_pivot_html(
                df_filtered,
                target_acc_dims,
                target_acc_metrics,
                total_title="大盘总计 / 目标账户汇总",
            )
            st.markdown(html_target_acc, unsafe_allow_html=True)

        # 表格 4: 投放平台 × 剧目类型成效透视明细表
        st.subheader("🎭 表格 4: 投放平台 × 剧目类型成效透视明细表")
        if not df_filtered.empty:
            type_col_name = "漫剧类型" if "漫剧类型" in df_filtered.columns else "书籍类型"
            type_dims = ["剧目/书籍语言", "投放平台", "版权来源", type_col_name]
            type_metrics = [
                "花费 ($)",
                "消耗占比 (%)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_type = generate_pivot_html(
                df_filtered, type_dims, type_metrics, total_title="大盘总计 / 类型汇总"
            )
            st.markdown(html_type, unsafe_allow_html=True)

        # 表格 5: 投放平台 × OS端透视明细表
        st.subheader("💻 表格 5: 投放平台 × OS端透视明细表")
        if not df_filtered.empty:
            os1_dims = [
                "剧目/书籍语言",
                "投放平台",
                "操作系统 (广告)",
                "操作系统 (用户)",
            ]
            os1_metrics = [
                "花费 ($)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_os1 = generate_pivot_html(
                df_filtered, os1_dims, os1_metrics, total_title="大盘总计 / OS汇总"
            )
            st.markdown(html_os1, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # Sheet 3: 第三部分：投放维度拆分
    # -------------------------------------------------------------------------
    with sheet3:
        part2_tab1, part2_tab2 = st.tabs(
            ["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"]
        )

        def render_part2_content(df_p2_working, is_daily_mode):
            time_suffix = f"(当日明细 - {latest_date})" if is_daily_mode else "(全量汇总)"

            if df_p2_working.empty:
                st.warning("无可展示的数据")
                return

            type_col_name = "漫剧类型" if "漫剧类型" in df_p2_working.columns else "书籍类型"

            st.subheader(f"📊 表格 1：具体投放剧目类型情况拆分表 {time_suffix}")
            t1_dims = ["剧目/书籍语言", "版权来源", type_col_name]
            t1_metrics = [
                "花费 ($)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_t1 = generate_pivot_html(
                df_p2_working, t1_dims, t1_metrics, total_title="类型拆分大盘汇总"
            )
            st.markdown(html_t1, unsafe_allow_html=True)

            st.subheader(f"🎬 表格 2：具体投放剧目情况表 {time_suffix}")
            t2_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "版权来源"]
            if "剧目/书籍中文名称" in df_p2_working.columns:
                t2_dims.insert(2, "剧目/书籍中文名称")

            t2_metrics = [
                "花费 ($)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_t2 = generate_pivot_html(
                df_p2_working, t2_dims, t2_metrics, total_title="剧目成效大盘汇总"
            )
            st.markdown(html_t2, unsafe_allow_html=True)

            st.subheader(f"📈 表格 3: 具体投放剧目每日数据情况表 {time_suffix}")
            t3_dims = ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID", "日期_str"]
            if "剧目/书籍中文名称" in df_p2_working.columns:
                t3_dims.insert(2, "剧目/书籍中文名称")

            t3_metrics = [
                "花费 ($)",
                "达标率 (%)",
                "D0 ROAS (%)",
                "CPI ($)",
                "首日付费率 (%)",
                "CPP ($)",
                "ARPPU ($)",
                "新注册用户占比 (%)",
                "付费用户新注册用户占比 (%)",
            ]
            html_t3 = generate_pivot_html(
                df_p2_working, t3_dims, t3_metrics, total_title="剧目每日趋势汇总"
            )
            st.markdown(html_t3, unsafe_allow_html=True)

        with part2_tab1:
            df_p2_daily = df_filtered[
                df_filtered["日期_str"] == latest_date
            ].copy()
            render_part2_content(df_p2_daily, is_daily_mode=True)

        with part2_tab2:
            render_part2_content(df_filtered, is_daily_mode=False)

    # -------------------------------------------------------------------------
    # Sheet 4: 第四部分：优化方式&出价深度拆分
    # -------------------------------------------------------------------------
    with sheet4:
        part3_tab1, part3_tab2 = st.tabs(
            ["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"]
        )

        def render_part3_content(df_p3_working, is_daily_mode):
            suffix_title = "(当日实况)" if is_daily_mode else "(全量总计)"
            total_title = (
                f"大盘总计 / 当日实况汇总"
                if is_daily_mode
                else f"大盘总计 / 全量总计汇总"
            )

            if is_daily_mode:
                st.info(
                    f"📌 当前展示【最新一日 {latest_date}】的数据聚合透视"
                )
            else:
                st.info("🌐 当前展示【底表所有日期】的全量数据聚合透视")

            # 板块 1: 渠道维度情况
            st.subheader("📊 板块 1: 渠道维度情况 (FB / TT)")

            st.markdown(
                f"##### 📌 表格 1: 投放平台指标汇总与消耗占比透视 {suffix_title}"
            )
            if not df_p3_working.empty:
                c1_metrics = [
                    "花费 ($)",
                    "花费占比 (%)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_c1 = generate_pivot_html(
                    df_p3_working,
                    ["剧目/书籍语言", "优化师", "投放平台"],
                    c1_metrics,
                    total_title=total_title,
                )
                st.markdown(html_c1, unsafe_allow_html=True)

            # 板块 2: 优化方式维度情况
            st.subheader("⚙️ 板块 2: 优化方式维度情况")

            st.markdown(
                f"##### 📌 表格 1: 大盘 - 优化方式 & 出价透视表 {suffix_title}"
            )
            if not df_p3_working.empty:
                o1_dims = [
                    "剧目/书籍语言",
                    "投放平台",
                    "推广目标",
                    "优化方式",
                    "手动出价的金额/系数",
                ]
                o1_metrics = [
                    "花费 ($)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_o1 = generate_pivot_html(
                    df_p3_working,
                    o1_dims,
                    o1_metrics,
                    total_title=total_title,
                )
                st.markdown(html_o1, unsafe_allow_html=True)

            # 表格 2: 极效高 ROAS 优化方式 & 出价透视表
            st.markdown(
                f'##### 📌 表格 2: 极效高 ROAS 优化方式 & 出价透视表 (花费>$100 且 达标率>=80%) <span style="color:red; font-weight:bold; font-size:14px;">(💡 广告方向优化策略，可查看此表！)</span> {suffix_title}',
                unsafe_allow_html=True,
            )
            if not df_p3_working.empty:
                tot_p3_cost = df_p3_working["花费"].sum()
                all_strategy_list = []

                for (d_name, l_v, p_v, t_v, m_v, b_v), sub_g in df_p3_working.groupby(
                    [
                        "剧目/书籍名称",
                        "剧目/书籍语言",
                        "投放平台",
                        "推广目标",
                        "优化方式",
                        "手动出价的金额/系数",
                    ]
                ):
                    m_sub = compute_metrics(sub_g)
                    c_ratio = (m_sub["花费"] / tot_p3_cost * 100) if tot_p3_cost > 0 else 0
                    all_strategy_list.append(
                        {
                            "key": (l_v, p_v, t_v, m_v, b_v),
                            "剧目/书籍名称": d_name,
                            "剧目/书籍语言": l_v,
                            "投放平台": p_v,
                            "推广目标": t_v,
                            "优化方式": m_v,
                            "手动出价的金额/系数": b_v,
                            "花费占比": c_ratio,
                            "达标率": m_sub["达标率"],
                            "花费": m_sub["花费"],
                        }
                    )

                high_roas_keys = [
                    item["key"]
                    for item in all_strategy_list
                    if item["花费"] > 100 and item["达标率"] >= 0.80
                ]

                if high_roas_keys:
                    df_high_roas = df_p3_working[
                        df_p3_working.set_index(
                            [
                                "剧目/书籍语言",
                                "投放平台",
                                "推广目标",
                                "优化方式",
                                "手动出价的金额/系数",
                            ]
                        ).index.isin(high_roas_keys)
                    ].copy()

                    o2_dims = [
                        "剧目/书籍语言",
                        "投放平台",
                        "推广目标",
                        "优化方式",
                        "手动出价的金额/系数",
                    ]
                    o2_metrics = [
                        "花费 ($)",
                        "达标率 (%)",
                        "D0 ROAS (%)",
                        "CPI ($)",
                        "首日付费率 (%)",
                        "CPP ($)",
                        "新注册用户占比 (%)",
                        "付费用户新注册用户占比 (%)",
                    ]
                    html_o2 = generate_pivot_html(
                        df_high_roas,
                        o2_dims,
                        o2_metrics,
                        total_title="高ROAS策略大盘汇总",
                    )
                    st.markdown(html_o2, unsafe_allow_html=True)
                else:
                    st.info("当前条件下，暂无同时满足【花费 > $100 且 达标率 >= 80%】的高 ROAS 出价策略。")

                # 💡 策略提效看板 1: 筛选阈值修正 (>5% 不达标止损预警)
                st.markdown("##### 💡 广告方向策略提效看板")
                df_all_summary = pd.DataFrame(all_strategy_list)

                with st.container(border=True):
                    card1_tab1, card1_tab2 = st.tabs(
                        [
                            "✅ 达标方向拓展策略 (达标率>80%, 花费占比>5%)",
                            "🚨 不达标方向预算缩减 (达标率<70%, 花费占比>5%)",
                        ]
                    )

                    with card1_tab1:
                        df_good = df_all_summary[
                            (df_all_summary["达标率"] > 0.80) & (df_all_summary["花费占比"] > 5.0)
                        ].sort_values(by="花费占比", ascending=False)

                        if not df_good.empty:
                            st.dataframe(
                                df_good[
                                    [
                                        "剧目/书籍名称",
                                        "投放平台",
                                        "推广目标",
                                        "优化方式",
                                        "手动出价的金额/系数",
                                        "花费占比",
                                    ]
                                ],
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "花费占比": st.column_config.ProgressColumn(
                                        "花费占比 (%)",
                                        help="拓展消耗占比进度条",
                                        format="%.2f%%",
                                        min_value=0,
                                        max_value=100,
                                    )
                                },
                            )
                        else:
                            st.info("暂无符合【达标率>80% 且 花费占比>5%】的拓展方向策略。")

                    with card1_tab2:
                        df_bad = df_all_summary[
                            (df_all_summary["达标率"] < 0.70) & (df_all_summary["花费占比"] > 5.0)
                        ].sort_values(by="花费占比", ascending=False)

                        if not df_bad.empty:
                            st.markdown(
                                '<span style="color:red; font-weight:bold;">🚨 以下方向达标率低于 70% 且消耗超 5%，建议立即下调预算或止损！</span>',
                                unsafe_allow_html=True,
                            )
                            st.dataframe(
                                df_bad[
                                    [
                                        "剧目/书籍名称",
                                        "投放平台",
                                        "推广目标",
                                        "优化方式",
                                        "手动出价的金额/系数",
                                        "花费占比",
                                    ]
                                ],
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "花费占比": st.column_config.ProgressColumn(
                                        "花费占比 (%)",
                                        help="低效浪费占比进度条",
                                        format="%.2f%%",
                                        min_value=0,
                                        max_value=100,
                                    )
                                },
                            )
                        else:
                            st.success("🎉 未发现【达标率<70% 且 花费占比>5%】的低效浪费方向！")

            # 板块 3: 投放平台 & Pixel 名称数据维度情况
            st.subheader("🎯 板块 3: 投放平台 & Pixel 名称数据维度情况")

            st.markdown(f"##### 📌 表格 1: Pixel 维度明细汇总表 {suffix_title}")
            if not df_p3_working.empty:
                pix1_dims = [
                    "剧目/书籍语言",
                    "投放平台",
                    "推广目标",
                    "优化方式",
                    "pixel名称",
                ]
                pix_metrics = [
                    "花费 ($)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "CPP ($)",
                    "首日付费率 (%)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_pix1 = generate_pivot_html(
                    df_p3_working,
                    pix1_dims,
                    pix_metrics,
                    total_title=total_title,
                )
                st.markdown(html_pix1, unsafe_allow_html=True)

            type_col_name = (
                "漫剧类型" if "漫剧类型" in df_p3_working.columns else "书籍类型"
            )
            st.markdown(f"##### 📌 表格 2: 平台/版权/类型 & Pixel 透视表 {suffix_title}")
            if not df_p3_working.empty:
                pix2_dims = [
                    "剧目/书籍语言",
                    "投放平台",
                    "版权来源",
                    type_col_name,
                    "pixel名称",
                ]
                html_pix2 = generate_pivot_html(
                    df_p3_working,
                    pix2_dims,
                    pix_metrics,
                    total_title=total_title,
                )
                st.markdown(html_pix2, unsafe_allow_html=True)

            # 板块 4: 不同语种的账户时区数据透视表
            st.subheader("🌐 板块 4: 不同语种的账户时区数据透视表")

            if not df_p3_working.empty:
                tz_metrics = [
                    "花费 ($)",
                    "消耗占比 (%)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "CPP ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                langs_in_df = df_p3_working["剧目/书籍语言"].unique().tolist()

                for l_idx, lang in enumerate(langs_in_df, start=1):
                    st.markdown(
                        f"##### 📌 表格 {l_idx}：{lang} - 账户时区数据透视 {suffix_title}"
                    )
                    group_l = df_p3_working[
                        df_p3_working["剧目/书籍语言"] == lang
                    ]
                    html_tz = generate_pivot_html(
                        group_l,
                        ["账户时区"],
                        tz_metrics,
                        total_title=f"{lang} 小计",
                    )
                    st.markdown(html_tz, unsafe_allow_html=True)

            # 板块 5: 当日新老广告效果对比
            st.subheader("⏱️ 板块 5: 当日新老广告效果对比 (按每日拆分展平)")

            if not df_p3_working.empty:
                df_p5 = df_p3_working.copy()

                def get_dynamic_ad_age(row):
                    try:
                        cur_date = pd.to_datetime(row["日期_str"])
                        create_date = pd.to_datetime(row["广告创建时间"])
                        diff_days = (cur_date - create_date).days

                        if 0 <= diff_days <= 1:
                            return "当日新建广告"
                        return "老广告"
                    except:
                        return "老广告"

                if "广告创建时间" in df_p5.columns:
                    df_p5["新老广告分类"] = df_p5.apply(
                        get_dynamic_ad_age, axis=1
                    )
                else:
                    df_p5["新老广告分类"] = "老广告"

                ad_metrics = [
                    "花费 ($)",
                    "消耗占比 (%)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "CPP ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]

                target_langs = ["繁体中文", "韩语"]
                for l_idx, lang in enumerate(target_langs, start=1):
                    st.markdown(
                        f"##### 📌 表格 {l_idx}：{lang} - 当日新老广告效果对比 {suffix_title}"
                    )
                    group_l = df_p5[df_p5["剧目/书籍语言"] == lang]
                    if group_l.empty:
                        st.info(f"暂无 {lang} 相关数据")
                        continue

                    html_ad = generate_pivot_html(
                        group_l,
                        ["日期_str", "新老广告分类"],
                        ad_metrics,
                        total_title=f"{lang} 小计",
                    )
                    st.markdown(html_ad, unsafe_allow_html=True)

            # 板块 6: 结构调整与新增表格 2
            st.subheader("💻 板块 6: 不同操作系统数据维度情况")

            st.markdown(
                f"##### 📌 表格 1: 操作系统 & 广告优化方式数据汇总 {suffix_title}"
            )
            if not df_p3_working.empty:
                os2_dims = [
                    "剧目/书籍语言",
                    "投放平台",
                    "pixel名称",
                    "操作系统 (广告)",
                    "操作系统 (用户)",
                    "推广目标",
                    "优化方式",
                    "手动出价的金额/系数",
                ]
                os2_metrics = [
                    "花费 ($)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]
                html_os1 = generate_pivot_html(
                    df_p3_working,
                    os2_dims,
                    os2_metrics,
                    total_title=total_title,
                )
                st.markdown(html_os1, unsafe_allow_html=True)

            st.markdown(
                f'##### 📌 表格 2: 达标 操作系统 & 广告优化方式数据汇总 <span style="color:red; font-weight:bold; font-size:14px;">(💡 最全Pixel+优化策略，可查看此表！)</span> {suffix_title}',
                unsafe_allow_html=True,
            )
            if not df_p3_working.empty:
                qualified_os_items = []
                for keys_combo, sub_g in df_p3_working.groupby(
                    [
                        "剧目/书籍语言",
                        "投放平台",
                        "pixel名称",
                        "操作系统 (广告)",
                        "操作系统 (用户)",
                        "推广目标",
                        "优化方式",
                        "手动出价的金额/系数",
                    ]
                ):
                    m_sub = compute_metrics(sub_g)
                    if m_sub["花费"] > 100 and m_sub["达标率"] >= 0.80:
                        qualified_os_items.append(keys_combo)

                if qualified_os_items:
                    df_qualified_os = df_p3_working[
                        df_p3_working.set_index(
                            [
                                "剧目/书籍语言",
                                "投放平台",
                                "pixel名称",
                                "操作系统 (广告)",
                                "操作系统 (用户)",
                                "推广目标",
                                "优化方式",
                                "手动出价的金额/系数",
                            ]
                        ).index.isin(qualified_os_items)
                    ].copy()

                    html_os2 = generate_pivot_html(
                        df_qualified_os,
                        os2_dims,
                        os2_metrics,
                        total_title="达标OS与策略大盘汇总",
                    )
                    st.markdown(html_os2, unsafe_allow_html=True)
                else:
                    st.info("当前条件下，暂无同时满足【花费 > $100 且 D0 达标率 >= 80%】的操作系统优化策略。")

            # 💡 最全 Pixel + 优化策略提效看板
            st.markdown("##### 💡 最全 Pixel + 优化策略提效看板")
            tot_p3_cost = df_p3_working["花费"].sum()
            all_pixel_strategy_list = []

            for (d_name, pix_v, plat_v, tgt_v, mode_v, bid_v), sub_g in df_p3_working.groupby(
                [
                    "剧目/书籍名称",
                    "pixel名称",
                    "投放平台",
                    "推广目标",
                    "优化方式",
                    "手动出价的金额/系数",
                ]
            ):
                m_sub = compute_metrics(sub_g)
                c_ratio = (m_sub["花费"] / tot_p3_cost * 100) if tot_p3_cost > 0 else 0
                all_pixel_strategy_list.append(
                    {
                        "剧目/书籍名称": d_name,
                        "pixel名称": pix_v,
                        "投放平台": plat_v,
                        "推广目标": tgt_v,
                        "优化方式": mode_v,
                        "手动出价的金额/系数": bid_v,
                        "花费占比": c_ratio,
                        "达标率": m_sub["达标率"],
                    }
                )

            if all_pixel_strategy_list:
                df_pixel_summary = pd.DataFrame(all_pixel_strategy_list)

                with st.container(border=True):
                    card2_tab1, card2_tab2 = st.tabs(
                        [
                            "✅ 达标 Pixel+策略 (达标率>80%, 花费占比>5%)",
                            "🚨 不达标方向预算缩减 (达标率<70%, 花费占比>5%)",
                        ]
                    )

                    with card2_tab1:
                        df_pix_good = df_pixel_summary[
                            (df_pixel_summary["达标率"] > 0.80)
                            & (df_pixel_summary["花费占比"] > 5.0)
                        ].sort_values(by="花费占比", ascending=False)

                        if not df_pix_good.empty:
                            st.dataframe(
                                df_pix_good[
                                    [
                                        "剧目/书籍名称",
                                        "pixel名称",
                                        "投放平台",
                                        "推广目标",
                                        "优化方式",
                                        "手动出价的金额/系数",
                                        "花费占比",
                                    ]
                                ],
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "花费占比": st.column_config.ProgressColumn(
                                        "花费占比 (%)",
                                        help="Pixel 拓展消耗占比进度条",
                                        format="%.2f%%",
                                        min_value=0,
                                        max_value=100,
                                    )
                                },
                            )
                        else:
                            st.info("暂无符合【达标率>80% 且 花费占比>5%】的优质 Pixel 策略。")

                    with card2_tab2:
                        df_pix_bad = df_pixel_summary[
                            (df_pixel_summary["达标率"] < 0.70)
                            & (df_pixel_summary["花费占比"] > 5.0)
                        ].sort_values(by="花费占比", ascending=False)

                        if not df_pix_bad.empty:
                            st.markdown(
                                '<span style="color:red; font-weight:bold;">🚨 以下 Pixel 组合达标率低于 70% 且消耗超 5%，请立刻更换或关闭 Pixel！</span>',
                                unsafe_allow_html=True,
                            )
                            st.dataframe(
                                df_pix_bad[
                                    [
                                        "剧目/书籍名称",
                                        "pixel名称",
                                        "投放平台",
                                        "推广目标",
                                        "优化方式",
                                        "手动出价的金额/系数",
                                        "花费占比",
                                    ]
                                ],
                                use_container_width=True,
                                hide_index=True,
                                column_config={
                                    "花费占比": st.column_config.ProgressColumn(
                                        "花费占比 (%)",
                                        help="低效 Pixel 占用进度条",
                                        format="%.2f%%",
                                        min_value=0,
                                        max_value=100,
                                    )
                                },
                            )
                        else:
                            st.success("🎉 未发现【达标率<70% 且 花费占比>5%】的低效 Pixel 组合！")

        with part3_tab1:
            df_p3_daily = df_filtered[
                df_filtered["日期_str"] == latest_date
            ].copy()
            render_part3_content(df_p3_daily, is_daily_mode=True)

        with part3_tab2:
            render_part3_content(df_filtered, is_daily_mode=False)

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

                st.header(
                    f"🎯 {s_name}-{s_zh}-{single_drama_id} 深度透视 (首投时间：{s_first_date})"
                )

                p5_metrics = [
                    "花费 ($)",
                    "达标率 (%)",
                    "D0 ROAS (%)",
                    "CPI ($)",
                    "首日付费率 (%)",
                    "CPP ($)",
                    "ARPPU ($)",
                    "新注册用户占比 (%)",
                    "付费用户新注册用户占比 (%)",
                ]

                # 📊 表格 1: 投放平台数据拆分对比
                st.subheader("📊 表格 1: 投放平台数据拆分对比")
                html_p5_1 = generate_pivot_html(
                    df_single, ["投放平台"], p5_metrics, total_title="剧目平台大盘汇总"
                )
                st.markdown(html_p5_1, unsafe_allow_html=True)

                # 👨‍💻 表格 2: 优化师数据对比
                st.subheader("👨‍💻 表格 2: 优化师数据对比")
                html_p5_2 = generate_pivot_html(
                    df_single, ["优化师"], p5_metrics, total_title="剧目优化师大盘汇总"
                )
                st.markdown(html_p5_2, unsafe_allow_html=True)

                # 🎯 表格 3: 投放平台 × Pixel名称 成效透视对比
                st.subheader("🎯 表格 3: 投放平台 × Pixel名称 成效透视对比")
                html_p5_3 = generate_pivot_html(
                    df_single,
                    ["投放平台", "pixel名称"],
                    p5_metrics,
                    total_title="剧目Pixel大盘汇总",
                )
                st.markdown(html_p5_3, unsafe_allow_html=True)

                # 🚀 表格 4: 投放平台 × 推广目标 拆分对比
                st.subheader("🚀 表格 4: 投放平台 × 推广目标 拆分对比")
                html_p5_4 = generate_pivot_html(
                    df_single,
                    ["投放平台", "推广目标"],
                    p5_metrics,
                    total_title="剧目目标大盘汇总",
                )
                st.markdown(html_p5_4, unsafe_allow_html=True)

                # 🔬 表格 5: 细分全量深度透视
                st.subheader(
                    "🔬 表格 5: 细分全量深度透视（平台×广告类型×pixel×目标×方式×出价）"
                )
                p5_5_dims = [
                    "投放平台",
                    "广告类型",
                    "pixel名称",
                    "操作系统 (广告)",
                    "推广目标",
                    "优化方式",
                    "手动出价的金额/系数",
                ]
                html_p5_5 = generate_pivot_html(
                    df_single, p5_5_dims, p5_metrics, total_title="剧目全量细分汇总"
                )
                st.markdown(html_p5_5, unsafe_allow_html=True)

                # 🔥 表格 6: 优化策略方向（达标率>80%）
                st.subheader("🔥 表格 6: 优化策略方向（花费>$100 且 达标率>=80%）")
                p5_6_items = []
                for keys_combo, sub_g in df_single.groupby(p5_5_dims):
                    m_combo = compute_metrics(sub_g)
                    if m_combo["花费"] > 100 and m_combo["达标率"] >= 0.80:
                        p5_6_items.append(keys_combo)

                if p5_6_items:
                    df_p5_qualified = df_single[
                        df_single.set_index(p5_5_dims).index.isin(p5_6_items)
                    ].copy()

                    html_p5_6 = generate_pivot_html(
                        df_p5_qualified,
                        p5_5_dims,
                        p5_metrics,
                        total_title="达标策略方向大盘汇总",
                    )
                    st.markdown(html_p5_6, unsafe_allow_html=True)
                else:
                    st.info("当前剧目暂无满足【花费 > $100 且 达标率 >= 80%】的高高效策略方向。")
        else:
            st.info("💡 提示：请在左侧侧边栏【🔍 搜索/筛选 推广剧目/书籍短ID】中**仅选择单个剧目**，即可在此处查阅该剧目的单剧目深度透视分析！")

    # -------------------------------------------------------------------------
    # Sheet 6: 第六部分：素材分析
    # -------------------------------------------------------------------------
    with sheet6:
        mat_tab1, mat_tab2 = st.tabs(
            ["📅 当日数据 (最新一日)", "🌐 全量总计 (所有时间汇总)"]
        )

        def render_material_analysis(df_mat_working, is_daily_mode):
            time_suffix = f"(当日明细 - {latest_date})" if is_daily_mode else "(全量汇总)"

            if df_mat_working.empty:
                st.warning("当前没有可分析的素材数据")
                return

            qualified_mat_keys = []
            mat_group_dims = [
                "剧目/书籍语言",
                "剧目/书籍名称",
                "剧目/书籍短ID",
                "素材名称",
            ]

            for keys_combo, sub_mat in df_mat_working.groupby(mat_group_dims):
                m_mat = compute_metrics(sub_mat)
                if m_mat["达标率"] >= 0.80 and m_mat["花费"] > 0:
                    qualified_mat_keys.append(keys_combo)

            if qualified_mat_keys:
                df_mat_qualified = df_mat_working[
                    df_mat_working.set_index(mat_group_dims).index.isin(qualified_mat_keys)
                ].copy()

                # 带【剧目信息抬头】的一键复制文本生成
                copy_text_lines = []
                for (s_lang, s_name, s_id), drama_mat_sub in df_mat_qualified.groupby(
                    ["剧目/书籍语言", "剧目/书籍名称", "剧目/书籍短ID"]
                ):
                    zh_title = drama_mat_sub.get("剧目/书籍中文名称", pd.Series(["-"])).iloc[0]
                    header_line = f"{s_name} - {zh_title} - {s_id}"
                    copy_text_lines.append(header_line)
                    copy_text_lines.append("达标素材清单 (一行一个):")

                    m_names = sorted(drama_mat_sub["素材名称"].dropna().unique().tolist())
                    for m_n in m_names:
                        copy_text_lines.append(m_n)
                    copy_text_lines.append("")

                mat_full_block = "\n".join(copy_text_lines)

                # UI 卡片化
                st.markdown(
                    f'<div class="material-copy-card"><div class="material-header-title">📋 复制达标素材名称列表 (带剧目信息抬头)</div>',
                    unsafe_allow_html=True,
                )
                with st.expander("点击展开全量达标素材清单 (可点击右上角图标一键复制)"):
                    st.text_area(
                        "全量达标素材文本框:",
                        value=mat_full_block,
                        height=220,
                        help="直接全选 Ctrl+A 复制，或点击框右上角复制文本",
                    )
                st.markdown("</div>", unsafe_allow_html=True)

                st.subheader(f"📊 表格 1: 具体投放剧目素材拆分表 (D0达标率>=80%) {time_suffix}")

                mat_rows = []
                for keys_combo, drama_mat_sub in df_mat_qualified.groupby(mat_group_dims):
                    m_res = compute_metrics(drama_mat_sub)
                    s_zh = drama_mat_sub.get("剧目/书籍中文名称", pd.Series(["-"])).iloc[0]
                    mat_rows.append(
                        {
                            "剧目/书籍语言": keys_combo[0],
                            "剧目/书籍名称": keys_combo[1],
                            "中文推广剧目": s_zh,
                            "剧目ID": keys_combo[2],
                            "素材名称": keys_combo[3],
                            "花费 ($)": m_res["花费"],
                            "D0 达标率 (%)": m_res["达标率"],
                            "首日付费用户数": int(m_res["首日付费用户数"]),
                            "D0 ROAS (%)": m_res["D0 ROAS"],
                            "CPI ($)": m_res["CPI"],
                            "首日付费率 (%)": m_res["首日付费率"],
                            "首日激活用户数": int(m_res["首日激活用户数"]),
                            "新注册用户占比 (%)": m_res["新注册用户占比"],
                            "付费用户新注册用户占比 (%)": m_res["付费用户新注册用户占比"],
                        }
                    )

                df_mat_display = pd.DataFrame(mat_rows).sort_values(
                    by="花费 ($)", ascending=False
                )

                # 使用交互式 DataFrame，支持鼠标框选与 Ctrl+C 复制
                st.dataframe(
                    df_mat_display,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "花费 ($)": st.column_config.NumberColumn(
                            "花费 ($)", format="$%.2f"
                        ),
                        "D0 达标率 (%)": st.column_config.NumberColumn(
                            "D0 达标率 (%)", format="%.2f%%"
                        ),
                        "首日付费用户数": st.column_config.NumberColumn(
                            "首日付费用户数", format="%d"
                        ),
                        "D0 ROAS (%)": st.column_config.NumberColumn(
                            "D0 ROAS (%)", format="%.2f%%"
                        ),
                        "CPI ($)": st.column_config.NumberColumn(
                            "CPI ($)", format="$%.2f"
                        ),
                        "首日付费率 (%)": st.column_config.NumberColumn(
                            "首日付费率 (%)", format="%.2f%%"
                        ),
                        "首日激活用户数": st.column_config.NumberColumn(
                            "首日激活用户数", format="%d"
                        ),
                        "新注册用户占比 (%)": st.column_config.NumberColumn(
                            "新注册用户占比 (%)", format="%.2f%%"
                        ),
                        "付费用户新注册用户占比 (%)": st.column_config.NumberColumn(
                            "付费用户新注册用户占比 (%)", format="%.2f%%"
                        ),
                    },
                )
            else:
                st.info("当前时间窗口内，暂无【D0 达标率 >= 80%】的达标素材。")

        with mat_tab1:
            df_mat_daily = df_filtered[
                df_filtered["日期_str"] == latest_date
            ].copy()
            render_material_analysis(df_mat_daily, is_daily_mode=True)

        with mat_tab2:
            render_material_analysis(df_filtered, is_daily_mode=False)
