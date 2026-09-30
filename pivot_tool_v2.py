import io
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# 1. 页面基本配置与全局 CSS 样式渲染
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="短剧/漫剧综合数据看板",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
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
    /* 高亮背景规范 */
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
    .cell-highlight-green {
        background-color: #D4EDDA !important; /* 达标率 >= 85% 淡绿色 */
        color: #155724;
        font-weight: bold;
    }
    .cell-highlight-pink {
        background-color: #F8D7DA !important;
        color: #721C24;
    }
    .text-green-bold {
        color: #28A745;
        font-weight: bold;
    }
    .row-muted {
        color: #6C757D;
    }
    .col-trend {
        font-size: 12px;
        color: #555555;
    }
    
    /* 十字高亮 (Crosshair Read Mode) CSS 支持 */
    .crosshair-focus {
        background-color: #FFE8A1 !important;
    }
    .crosshair-row {
        background-color: #F1F3F5;
    }
    .crosshair-col {
        background-color: #F1F3F5;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🎬 短剧/漫剧综合数据看板")

# -----------------------------------------------------------------------------
# 2. Excel 文件读取与数据预处理
# -----------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "📂 请上传数据底表 (.xlsx / .xls)", type=["xlsx", "xls"]
)


def clean_num(val):
    try:
        if pd.isna(val) or val == "-" or val == "":
            return 0.0
        return float(val)
    except:
        return 0.0


def get_trend_str(curr, prev, is_cost=False):
    if pd.isna(prev) or prev == 0:
        return "-"
    diff = curr - prev
    if is_cost:
        if diff > 0:
            return f"+{abs(diff):.2f} ↓"
        return "$0.00"
    else:
        pct = (diff / prev) * 100
        if pct > 0:
            return f"+{pct:.2f}% ↑"
        elif pct < 0:
            return f"{pct:.2f}% ↓"
        return "0.00%"


def generate_pivot_html(
    df_data, dims, metric_cols, total_title="大盘总计 / 全局汇总"
):
    """通用透视表生成算法：支持多级维度跨行合并 (rowspan)，严格按花费降序"""
    if df_data.empty:
        return "<p>无符合条件的数据</p>"
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
                node["leaf_data"] = {
                    "花费": c_cost,
                    "达标率": (sub_df["D0达标率"] * sub_df["花费"]).sum()
                    / (c_cost or 1),
                    "D0 ROAS": (sub_df["d0_roas"] * sub_df["花费"]).sum()
                    / (c_cost or 1),
                    "CPI": (sub_df["CPI"] * sub_df["花费"]).sum()
                    / (c_cost or 1),
                    "首日付费率": (
                        sub_df["首日付费率"] * sub_df["花费"]
                    ).sum()
                    / (c_cost or 1),
                    "CPP": (sub_df["首日付费用户成本"] * sub_df["花费"]).sum()
                    / (c_cost or 1),
                    "ARPPU": (
                        sub_df["Arppu- 付费用户人均付费金额"]
                        * sub_df["花费"]
                    ).sum()
                    / (c_cost or 1),
                    "新注册用户占比": (
                        sub_df["新注册用户占比"] * sub_df["花费"]
                    ).sum()
                    / (c_cost or 1),
                    "付费用户新注册用户占比": (
                        sub_df["付费用户新注册用户占比"] * sub_df["花费"]
                    ).sum()
                    / (c_cost or 1),
                }
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

    # 大盘总计行 (#FFE066)
    if total_cost > 0:
        d0_tot = (df_data["D0达标率"] * df_data["花费"]).sum() / total_cost
        roas_tot = (df_data["d0_roas"] * df_data["花费"]).sum() / total_cost
        cpi_tot = (df_data["CPI"] * df_data["花费"]).sum() / total_cost
        pay_tot = (df_data["首日付费率"] * df_data["花费"]).sum() / total_cost
        cpp_tot = (
            df_data["首日付费用户成本"] * df_data["花费"]
        ).sum() / total_cost
        arppu_tot = (
            df_data["Arppu- 付费用户人均付费金额"] * df_data["花费"]
        ).sum() / total_cost
        reg_tot = (
            df_data["新注册用户占比"] * df_data["花费"]
        ).sum() / total_cost
        preg_tot = (
            df_data["付费用户新注册用户占比"] * df_data["花费"]
        ).sum() / total_cost

        d0_cell = (
            'class="cell-highlight-green"' if d0_tot >= 0.85 else ""
        )

        html += f'<tr class="row-total"><td colspan="{len(dims)}">{total_title}</td>'
        for mc in metric_cols:
            if mc == "花费 ($)":
                html += f"<td>${total_cost:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]:
                html += "<td>100.00%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]:
                html += f"<td {d0_cell}>{d0_tot*100:.2f}%</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]:
                html += f"<td>{roas_tot*100:.2f}%</td>"
            elif mc == "CPI ($)":
                html += f"<td>${cpi_tot:.2f}</td>"
            elif mc == "首日付费率 (%)":
                html += f"<td>{pay_tot*100:.2f}%</td>"
            elif mc == "CPP ($)":
                html += f"<td>${cpp_tot:.2f}</td>"
            elif mc == "ARPPU ($)":
                html += f"<td>${arppu_tot:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]:
                html += f"<td>{reg_tot*100:.2f}%</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]:
                html += f"<td>{preg_tot*100:.2f}%</td>"
        html += "</tr>"

    rendered_dims_counts = [0] * len(dims)

    # 需要高亮的前几个维度名称定义
    yellow_dims = [
        "优化师",
        "剧目/书籍语言",
        "剧目/书籍名称",
        "剧目/书籍短ID",
        "投放平台",
        "操作系统 (广告)",
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
        d0_cell = (
            'class="cell-highlight-green"' if d0_val >= 0.85 else ""
        )
        cost_ratio = (m["花费"] / total_cost * 100) if total_cost > 0 else 0

        for mc in metric_cols:
            if mc == "花费 ($)":
                html += f"<td>${m['花费']:,.2f}</td>"
            elif mc in ["花费占比 (%)", "消耗占比 (%)"]:
                html += f"<td>{cost_ratio:.2f}%</td>"
            elif mc in ["达标率 (%)", "D0 达标率 (%)"]:
                html += f"<td {d0_cell}>{d0_val*100:.2f}%</td>"
            elif mc in ["D0 ROAS (%)", "ROAS (%)"]:
                html += f"<td>{m['D0 ROAS']*100:.2f}%</td>"
            elif mc == "CPI ($)":
                html += f"<td>${m['CPI']:.2f}</td>"
            elif mc == "首日付费率 (%)":
                html += f"<td>{m['首日付费率']*100:.2f}%</td>"
            elif mc == "CPP ($)":
                html += f"<td>${m['CPP']:.2f}</td>"
            elif mc == "ARPPU ($)":
                html += f"<td>${m['ARPPU']:.2f}</td>"
            elif mc in ["新注册用户占比 (%)", "注册拉新占比 (%)"]:
                html += f"<td>{m['新注册用户占比']*100:.2f}%</td>"
            elif mc in ["付费用户新注册用户占比 (%)", "付费拉新占比 (%)"]:
                html += f"<td>{m['付费用户新注册用户占比']*100:.2f}%</td>"
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


if uploaded_file is not None:
    raw_df = pd.read_excel(uploaded_file)

    # 1. 全局过滤零花费数据 (Cost > 0)
    df = raw_df[(raw_df["优化师"] != "-") & (raw_df["花费"] > 0)].copy()

    # 2. 日期转换与处理
    df["日期_str"] = pd.to_datetime(df["日期"]).dt.strftime("%Y-%m-%d")
    df["首次投放_str"] = pd.to_datetime(df["首次投放的日期"]).dt.strftime(
        "%Y-%m-%d"
    )

    all_dates = sorted(df["日期_str"].unique(), reverse=True)
    latest_date = all_dates[0] if len(all_dates) > 0 else ""
    prev_date = all_dates[1] if len(all_dates) > 1 else ""

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

    df["账户时区分类"] = df["投放账户名称"].apply(get_timezone)

    # -------------------------------------------------------------------------
    # 三、 控制栏组件重构 (8大筛选 + 新增【💡 阅读模式开关】)
    # -------------------------------------------------------------------------
    st.markdown("### 🌐 控制栏")

    r1_1, r1_2, r1_3 = st.columns([1, 2, 1])
    with r1_1:
        lang_tab = st.selectbox(
            "快捷语种切换：",
            options=["🌐 全部地区/双语种", "🇹🇼 繁体中文", "🇰🇷 韩语"],
            index=0,
        )

    if lang_tab == "🇹🇼 繁体中文":
        df = df[df["剧目/书籍语言"] == "繁体中文"]
    elif lang_tab == "🇰🇷 韩语":
        df = df[df["剧目/书籍语言"] == "韩语"]

    with r1_2:
        selected_opt = st.multiselect(
            "选择优化师 (`优化师`):",
            options=sorted(df["优化师"].dropna().unique().tolist()),
            default=sorted(df["优化师"].dropna().unique().tolist()),
        )

    # 阅读模式（十字高亮）开关键
    with r1_3:
        enable_read_mode = st.toggle(
            "💡 开启/关闭阅读模式 (十字高亮)", value=True
        )

    r2_1, r2_2, r2_3 = st.columns(3)
    with r2_1:
        selected_shows = st.multiselect(
            "选择剧目名称 (`剧目/书籍名称`):",
            options=sorted(df["剧目/书籍名称"].dropna().unique().tolist()),
            default=sorted(df["剧目/书籍名称"].dropna().unique().tolist()),
        )
        selected_ids = st.multiselect(
            "选择剧目ID (`剧目/书籍短ID`):",
            options=sorted(
                df["剧目/书籍短ID"].dropna().astype(str).unique().tolist()
            ),
            default=sorted(
                df["剧目/书籍短ID"].dropna().astype(str).unique().tolist()
            ),
        )

    with r2_2:
        selected_sources = st.multiselect(
            "选择版权来源 (`版权来源`):",
            options=sorted(df["版权来源"].dropna().unique().tolist()),
            default=sorted(df["版权来源"].dropna().unique().tolist()),
        )
        selected_pixels = st.multiselect(
            "选择Pixel名称 (`pixel名称`):",
            options=sorted(df["pixel名称"].dropna().unique().tolist()),
            default=sorted(df["pixel名称"].dropna().unique().tolist()),
        )

    with r2_3:
        selected_opt_methods = st.multiselect(
            "选择优化方式 (`优化方式`):",
            options=sorted(df["优化方式"].dropna().unique().tolist()),
            default=sorted(df["优化方式"].dropna().unique().tolist()),
        )
        selected_os = st.multiselect(
            "选择OS操作系统 (`操作系统 (广告)`):",
            options=sorted(df["操作系统 (广告)"].dropna().unique().tolist()),
            default=sorted(df["操作系统 (广告)"].dropna().unique().tolist()),
        )

    df["剧目/书籍短ID_str"] = df["剧目/书籍短ID"].astype(str)
    df_filtered = df[
        (df["优化师"].isin(selected_opt))
        & (df["剧目/书籍名称"].isin(selected_shows))
        & (df["剧目/书籍短ID_str"].isin(selected_ids))
        & (df["版权来源"].isin(selected_sources))
        & (df["pixel名称"].isin(selected_pixels))
        & (df["优化方式"].isin(selected_opt_methods))
        & (df["操作系统 (广告)"].isin(selected_os))
    ].copy()

    # -------------------------------------------------------------------------
    # 注入阅读模式 (Crosshair Reading Mode) 前端 JS
    # -------------------------------------------------------------------------
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

                        // 清除旧高亮
                        table.querySelectorAll('.crosshair-row, .crosshair-col, .crosshair-focus')
                             .forEach(el => el.classList.remove('crosshair-row', 'crosshair-col', 'crosshair-focus'));

                        const tr = td.parentElement;
                        const colIndex = td.cellIndex;

                        // 高亮当前焦点单元格
                        td.classList.add('crosshair-focus');

                        // 高亮当前行
                        tr.querySelectorAll('td, th').forEach(c => c.classList.add('crosshair-row'));

                        // 高亮当前列
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

    st.markdown("---")

    # =========================================================================
    # 第一部分：整体数据看板
    # =========================================================================
    st.header("📌 第一部分：整体数据看板")

    # -------------------------------------------------------------------------
    # 板块一：不同优化师、不同语种当日数据实况表
    # -------------------------------------------------------------------------
    st.subheader(
        f"📊 板块一：不同优化师、不同语种当日数据实况表 ({latest_date})"
    )

    df_today = df_filtered[df_filtered["日期_str"] == latest_date].copy()

    if df_today.empty:
        st.warning(f"最新一日 ({latest_date}) 筛选条件下无可展现的数据。")
    else:
        total_cost_all = df_today["花费"].sum()

        html_p1 = """<table class="dataframe-table">
            <thead>
                <tr>
                    <th>日期</th>
                    <th>优化师</th>
                    <th>剧目/书籍语言</th>
                    <th>花费 ($)</th>
                    <th>消耗占比 (%)</th>
                    <th>D0 达标率 (%)</th>
                    <th>D0 ROAS (%)</th>
                    <th>CPI ($)</th>
                    <th>首日付费率 (%)</th>
                    <th>CPP ($)</th>
                    <th>ARPPU ($)</th>
                    <th>新注册用户占比 (%)</th>
                    <th>付费用户新注册用户占比 (%)</th>
                </tr>
            </thead>
            <tbody>"""

        if total_cost_all > 0:
            d0_reach = (
                (df_today["D0达标率"] * df_today["花费"]).sum()
                / total_cost_all
            )
            d0_roas = (
                (df_today["d0_roas"] * df_today["花费"]).sum()
                / total_cost_all
            )
            cpi = (
                df_today["CPI"] * df_today["花费"]
            ).sum() / total_cost_all
            pay_rate = (
                (df_today["首日付费率"] * df_today["花费"]).sum()
                / total_cost_all
            )
            cpp = (
                (df_today["首日付费用户成本"] * df_today["花费"]).sum()
                / total_cost_all
            )
            arppu = (
                (
                    df_today["Arppu- 付费用户人均付费金额"]
                    * df_today["花费"]
                ).sum()
                / total_cost_all
            )
            new_reg = (
                (df_today["新注册用户占比"] * df_today["花费"]).sum()
                / total_cost_all
            )
            pay_new_reg = (
                (
                    df_today["付费用户新注册用户占比"] * df_today["花费"]
                ).sum()
                / total_cost_all
            )

            d0_style = (
                'class="text-green-bold"' if d0_reach >= 0.8 else ""
            )

            html_p1 += f"""
            <tr class="row-total">
                <td>{latest_date}</td>
                <td colspan="2">大盘总体</td>
                <td>${total_cost_all:,.2f}</td>
                <td>100.00%</td>
                <td {d0_style}>{d0_reach*100:.2f}%</td>
                <td>{d0_roas*100:.2f}%</td>
                <td>${cpi:.2f}</td>
                <td>{pay_rate*100:.2f}%</td>
                <td>${cpp:.2f}</td>
                <td>${arppu:.2f}</td>
                <td>{new_reg*100:.2f}%</td>
                <td>{pay_new_reg*100:.2f}%</td>
            </tr>"""

        for lang, group in df_today.groupby("剧目/书籍语言"):
            lang_cost = group["花费"].sum()

            group_agg = (
                group.groupby("优化师")
                .apply(
                    lambda x: pd.Series(
                        {
                            "花费": x["花费"].sum(),
                            "D0达标率": (x["D0达标率"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "d0_roas": (x["d0_roas"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "CPI": (x["CPI"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "首日付费率": (
                                x["首日付费率"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "首日付费用户成本": (
                                x["首日付费用户成本"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "Arppu- 付费用户人均付费金额": (
                                x["Arppu- 付费用户人均付费金额"]
                                * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "新注册用户占比": (
                                x["新注册用户占比"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "付费用户新注册用户占比": (
                                x["付费用户新注册用户占比"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                        }
                    )
                )
                .reset_index()
            )

            group_agg = group_agg.sort_values(by="花费", ascending=False)
            row_count = len(group_agg)

            for idx, (_, row) in enumerate(group_agg.iterrows()):
                cost_ratio = (
                    (row["花费"] / total_cost_all * 100)
                    if total_cost_all > 0
                    else 0
                )
                d0_style = (
                    'class="text-green-bold"'
                    if row["D0达标率"] >= 0.8
                    else ""
                )

                html_p1 += "<tr>"
                if idx == 0:
                    html_p1 += f'<td rowspan="{row_count+1}">{latest_date}</td>'

                html_p1 += f"""
                    <td>{row['优化师']}</td>
                    <td>{lang}</td>
                    <td>${row['花费']:,.2f}</td>
                    <td>{cost_ratio:.2f}%</td>
                    <td {d0_style}>{row['D0达标率']*100:.2f}%</td>
                    <td>{row['d0_roas']*100:.2f}%</td>
                    <td>${row['CPI']:.2f}</td>
                    <td>{row['首日付费率']*100:.2f}%</td>
                    <td>${row['首日付费用户成本']:.2f}</td>
                    <td>${row['Arppu- 付费用户人均付费金额']:.2f}</td>
                    <td>{row['新注册用户占比']*100:.2f}%</td>
                    <td>{row['付费用户新注册用户占比']*100:.2f}%</td>
                </tr>"""

            l_d0 = (group["D0达标率"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_roas = (group["d0_roas"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_cpi = (group["CPI"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_pay = (group["首日付费率"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_cpp = (group["首日付费用户成本"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_arppu = (
                group["Arppu- 付费用户人均付费金额"] * group["花费"]
            ).sum() / (lang_cost or 1)
            l_reg = (group["新注册用户占比"] * group["花费"]).sum() / (
                lang_cost or 1
            )
            l_preg = (
                group["付费用户新注册用户占比"] * group["花费"]
            ).sum() / (lang_cost or 1)
            l_d0_style = (
                'class="text-green-bold"' if l_d0 >= 0.8 else ""
            )

            html_p1 += f"""
            <tr class="row-subtotal">
                <td colspan="2">{lang} 小计</td>
                <td>${lang_cost:,.2f}</td>
                <td>{(lang_cost/total_cost_all*100) if total_cost_all>0 else 0:.2f}%</td>
                <td {l_d0_style}>{l_d0*100:.2f}%</td>
                <td>{l_roas*100:.2f}%</td>
                <td>${l_cpi:.2f}</td>
                <td>{l_pay*100:.2f}%</td>
                <td>${l_cpp:.2f}</td>
                <td>${l_arppu:.2f}</td>
                <td>{l_reg*100:.2f}%</td>
                <td>{l_preg*100:.2f}%</td>
            </tr>"""

        html_p1 += "</tbody></table>"
        st.markdown(html_p1, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # 板块二：不同优化师、不同语种每日数据情况与趋势表
    # -------------------------------------------------------------------------
    st.subheader(
        "📊 2. 不同优化师、不同语种每日数据情况与趋势表"
    )

    if not df_filtered.empty:
        html_p2 = """<table class="dataframe-table">
            <thead>
                <tr>
                    <th>优化师</th>
                    <th>剧目/书籍语言</th>
                    <th>日期</th>
                    <th>当日消耗 ($)</th>
                    <th>D0 达标率 (%)</th>
                    <th>D0 ROAS (%)</th>
                    <th>CPI ($)</th>
                    <th>首日付费率 (%)</th>
                    <th>CPP ($)</th>
                    <th>ARPPU ($)</th>
                    <th>新注册用户占比 (%)</th>
                    <th>付费用户新注册用户占比 (%)</th>
                    <th>消耗趋势</th>
                    <th>达标率趋势</th>
                    <th>ROAS 趋势</th>
                    <th>CPI 趋势</th>
                </tr>
            </thead>
            <tbody>"""

        for lang, lang_df in df_filtered.groupby("剧目/书籍语言"):
            opts = lang_df["优化师"].unique()
            for opt in opts:
                opt_df = lang_df[lang_df["优化师"] == opt].copy()

                opt_daily = (
                    opt_df.groupby("日期_str")
                    .apply(
                        lambda x: pd.Series(
                            {
                                "花费": x["花费"].sum(),
                                "D0达标率": (
                                    x["D0达标率"] * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                                "d0_roas": (x["d0_roas"] * x["花费"]).sum()
                                / (x["花费"].sum() or 1),
                                "CPI": (x["CPI"] * x["花费"]).sum()
                                / (x["花费"].sum() or 1),
                                "首日付费率": (
                                    x["首日付费率"] * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                                "首日付费用户成本": (
                                    x["首日付费用户成本"] * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                                "Arppu- 付费用户人均付费金额": (
                                    x["Arppu- 付费用户人均付费金额"]
                                    * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                                "新注册用户占比": (
                                    x["新注册用户占比"] * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                                "付费用户新注册用户占比": (
                                    x["付费用户新注册用户占比"]
                                    * x["花费"]
                                ).sum()
                                / (x["花费"].sum() or 1),
                            }
                        )
                    )
                    .reset_index()
                )

                opt_daily = opt_daily.sort_values(
                    by="日期_str", ascending=False
                )
                opt_dates_list = opt_daily["日期_str"].tolist()
                num_rows = len(opt_daily) + 1

                for idx, (_, r) in enumerate(opt_daily.iterrows()):
                    cur_date = r["日期_str"]
                    cur_idx = opt_dates_list.index(cur_date)
                    prev_row = None
                    if cur_idx + 1 < len(opt_dates_list):
                        p_date = opt_dates_list[cur_idx + 1]
                        p_match = opt_daily[opt_daily["日期_str"] == p_date]
                        if not p_match.empty:
                            prev_row = p_match.iloc[0]

                    cost_trend = (
                        get_trend_str(r["花费"], prev_row["花费"])
                        if prev_row is not None
                        else "-"
                    )
                    d0_trend = (
                        get_trend_str(r["D0达标率"], prev_row["D0达标率"])
                        if prev_row is not None
                        else "-"
                    )
                    roas_trend = (
                        get_trend_str(r["d0_roas"], prev_row["d0_roas"])
                        if prev_row is not None
                        else "-"
                    )
                    cpi_trend = (
                        get_trend_str(
                            r["CPI"], prev_row["CPI"], is_cost=True
                        )
                        if prev_row is not None
                        else "-"
                    )

                    row_class = (
                        'class="row-muted"' if cur_date != latest_date else ""
                    )
                    d0_cell = (
                        'class="cell-highlight-green"'
                        if r["D0达标率"] >= 0.85
                        else ""
                    )

                    html_p2 += f"<tr {row_class}>"
                    if idx == 0:
                        html_p2 += f'<td rowspan="{num_rows}">{opt}</td>'
                        html_p2 += f'<td rowspan="{num_rows}">{lang}</td>'

                    html_p2 += f"""
                        <td>{cur_date}</td>
                        <td>${r['花费']:,.2f}</td>
                        <td {d0_cell}>{r['D0达标率']*100:.2f}%</td>
                        <td>{r['d0_roas']*100:.2f}%</td>
                        <td>${r['CPI']:.2f}</td>
                        <td>{r['首日付费率']*100:.2f}%</td>
                        <td>${r['首日付费用户成本']:.2f}</td>
                        <td>${r['Arppu- 付费用户人均付费金额']:.2f}</td>
                        <td>{r['新注册用户占比']*100:.2f}%</td>
                        <td>{r['付费用户新注册用户占比']*100:.2f}%</td>
                        <td class="col-trend">{cost_trend}</td>
                        <td class="col-trend">{d0_trend}</td>
                        <td class="col-trend">{roas_trend}</td>
                        <td class="col-trend">{cpi_trend}</td>
                    </tr>"""

                o_cost_tot = opt_df["花费"].sum()
                o_d0_3d = (opt_df["D0达标率"] * opt_df["花费"]).sum() / (
                    o_cost_tot or 1
                )
                o_roas_3d = (opt_df["d0_roas"] * opt_df["花费"]).sum() / (
                    o_cost_tot or 1
                )
                o_cpi_3d = (opt_df["CPI"] * opt_df["花费"]).sum() / (
                    o_cost_tot or 1
                )
                o_pay_3d = (opt_df["首日付费率"] * opt_df["花费"]).sum() / (
                    o_cost_tot or 1
                )
                o_cpp_3d = (
                    opt_df["首日付费用户成本"] * opt_df["花费"]
                ).sum() / (o_cost_tot or 1)
                o_arppu_3d = (
                    opt_df["Arppu- 付费用户人均付费金额"] * opt_df["花费"]
                ).sum() / (o_cost_tot or 1)
                o_reg_3d = (
                    opt_df["新注册用户占比"] * opt_df["花费"]
                ).sum() / (o_cost_tot or 1)
                o_preg_3d = (
                    opt_df["付费用户新注册用户占比"] * opt_df["花费"]
                ).sum() / (o_cost_tot or 1)

                o_d0_style = (
                    'class="cell-highlight-green"'
                    if o_d0_3d >= 0.85
                    else ""
                )

                html_p2 += f"""
                <tr class="row-subtotal">
                    <td>全期汇总</td>
                    <td>${o_cost_tot:,.2f}</td>
                    <td {o_d0_style}>{o_d0_3d*100:.2f}%</td>
                    <td>{o_roas_3d*100:.2f}%</td>
                    <td>${o_cpi_3d:.2f}</td>
                    <td>{o_pay_3d*100:.2f}%</td>
                    <td>${o_cpp_3d:.2f}</td>
                    <td>${o_arppu_3d:.2f}</td>
                    <td>{o_reg_3d*100:.2f}%</td>
                    <td>{o_preg_3d*100:.2f}%</td>
                    <td class="col-trend" colspan="4">-</td>
                </tr>"""

        html_p2 += "</tbody></table>"
        st.markdown(html_p2, unsafe_allow_html=True)

    st.markdown("---")

    # =========================================================================
    # 第二部分：剧集细拆分析 (板块三)
    # =========================================================================
    st.header("📌 第二部分：剧集细拆分析（板块三）")

    df_filtered["latest_date_dt"] = pd.to_datetime(latest_date)
    df_filtered["first_date_dt"] = pd.to_datetime(df_filtered["首次投放的日期"])
    df_filtered["is_new_drama"] = (
        df_filtered["latest_date_dt"] - df_filtered["first_date_dt"]
    ).dt.days <= 3

    # -------------------------------------------------------------------------
    # 表格 1：📊 具体投放剧目类型情况拆分表
    # -------------------------------------------------------------------------
    st.subheader("📊 表格 1：具体投放剧目类型情况拆分表")

    if not df_filtered.empty:
        html_t1 = """<table class="dataframe-table">
            <thead>
                <tr>
                    <th>版权来源</th>
                    <th>漫剧类型</th>
                    <th>花费 ($)</th>
                    <th>在投总数</th>
                    <th>新剧数量</th>
                    <th>达标率 (%)</th>
                    <th>新剧达标率 (%)</th>
                    <th>D0 ROAS (%)</th>
                    <th>CPI ($)</th>
                    <th>首日付费率 (%)</th>
                    <th>CPP ($)</th>
                    <th>ARPPU ($)</th>
                    <th>新注册用户占比 (%)</th>
                    <th>付费用户新注册用户占比 (%)</th>
                </tr>
            </thead>
            <tbody>"""

        for copyright_src, group_c in df_filtered.groupby("版权来源"):
            c_cost = group_c["花费"].sum()

            sub_list = []
            for m_type, group_m in group_c.groupby("漫剧类型"):
                m_cost = group_m["花费"].sum()
                total_shows = group_m["剧目/书籍名称"].nunique()

                new_shows_df = group_m[group_m["is_new_drama"]]
                new_shows_count = new_shows_df["剧目/书籍名称"].nunique()

                d0_reach = (group_m["D0达标率"] * group_m["花费"]).sum() / (
                    m_cost or 1
                )
                new_d0_reach = (
                    (new_shows_df["D0达标率"] * new_shows_df["花费"]).sum()
                    / (new_shows_df["花费"].sum() or 1)
                    if not new_shows_df.empty
                    else 0.0
                )

                sub_list.append(
                    {
                        "漫剧类型": m_type,
                        "花费": m_cost,
                        "在投总数": total_shows,
                        "新剧数量": new_shows_count,
                        "达标率": d0_reach,
                        "新剧达标率": new_d0_reach,
                        "D0 ROAS": (
                            group_m["d0_roas"] * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                        "CPI": (group_m["CPI"] * group_m["花费"]).sum()
                        / (m_cost or 1),
                        "首日付费率": (
                            group_m["首日付费率"] * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                        "CPP": (
                            group_m["首日付费用户成本"] * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                        "ARPPU": (
                            group_m["Arppu- 付费用户人均付费金额"]
                            * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                        "新注册用户占比": (
                            group_m["新注册用户占比"] * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                        "付费用户新注册用户占比": (
                            group_m["付费用户新注册用户占比"]
                            * group_m["花费"]
                        ).sum()
                        / (m_cost or 1),
                    }
                )

            sub_df = pd.DataFrame(sub_list).sort_values(
                by="花费", ascending=False
            )
            row_count = len(sub_df)

            for idx, (_, r) in enumerate(sub_df.iterrows()):
                d0_cell = (
                    'class="cell-highlight-green"'
                    if r["达标率"] >= 0.85
                    else ""
                )
                new_d0_cell = (
                    'class="cell-highlight-green"'
                    if r["新剧达标率"] >= 0.85
                    else ""
                )

                html_t1 += "<tr>"
                if idx == 0:
                    html_t1 += f'<td rowspan="{row_count+1}">{copyright_src}</td>'

                html_t1 += f"""
                    <td>{r['漫剧类型']}</td>
                    <td>${r['花费']:,.2f}</td>
                    <td>{r['在投总数']}</td>
                    <td>{r['新剧数量']}</td>
                    <td {d0_cell}>{r['达标率']*100:.2f}%</td>
                    <td {new_d0_cell}>{r['新剧达标率']*100:.2f}%</td>
                    <td>{r['D0 ROAS']*100:.2f}%</td>
                    <td>${r['CPI']:.2f}</td>
                    <td>{r['首日付费率']*100:.2f}%</td>
                    <td>${r['CPP']:.2f}</td>
                    <td>${r['ARPPU']:.2f}</td>
                    <td>{r['新注册用户占比']*100:.2f}%</td>
                    <td>{r['付费用户新注册用户占比']*100:.2f}%</td>
                </tr>"""

            c_d0 = (group_c["D0达标率"] * group_c["花费"]).sum() / (c_cost or 1)
            c_d0_cell = (
                'class="cell-highlight-green"' if c_d0 >= 0.85 else ""
            )
            c_new = group_c[group_c["is_new_drama"]]
            c_new_d0 = (
                (c_new["D0达标率"] * c_new["花费"]).sum() / (c_new["花费"].sum() or 1)
                if not c_new.empty
                else 0.0
            )
            c_new_d0_cell = (
                'class="cell-highlight-green"' if c_new_d0 >= 0.85 else ""
            )

            html_t1 += f"""
            <tr class="row-subtotal">
                <td>{copyright_src} 小计</td>
                <td>${c_cost:,.2f}</td>
                <td>{group_c['剧目/书籍名称'].nunique()}</td>
                <td>{c_new['剧目/书籍名称'].nunique()}</td>
                <td {c_d0_cell}>{c_d0*100:.2f}%</td>
                <td {c_new_d0_cell}>{c_new_d0*100:.2f}%</td>
                <td>{(group_c['d0_roas']*group_c['花费']).sum()/(c_cost or 1)*100:.2f}%</td>
                <td>${(group_c['CPI']*group_c['花费']).sum()/(c_cost or 1):.2f}</td>
                <td>{(group_c['首日付费率']*group_c['花费']).sum()/(c_cost or 1)*100:.2f}%</td>
                <td>${(group_c['首日付费用户成本']*group_c['花费']).sum()/(c_cost or 1):.2f}</td>
                <td>${(group_c['Arppu- 付费用户人均付费金额']*group_c['花费']).sum()/(c_cost or 1):.2f}</td>
                <td>{(group_c['新注册用户占比']*group_c['花费']).sum()/(c_cost or 1)*100:.2f}%</td>
                <td>{(group_c['付费用户新注册用户占比']*group_c['花费']).sum()/(c_cost or 1)*100:.2f}%</td>
            </tr>"""

        html_t1 += "</tbody></table>"
        st.markdown(html_t1, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # 表格 2：🎬 具体投放剧目情况表
    # -------------------------------------------------------------------------
    st.subheader(f"🎬 表格 2：具体投放剧目情况表 (当日明细 - {latest_date})")

    df_today_shows = df_filtered[df_filtered["日期_str"] == latest_date].copy()

    if not df_today_shows.empty:
        d_all_reach_dict = {}
        for s_name, s_group in df_filtered.groupby("剧目/书籍名称"):
            s_cost_all = s_group["花费"].sum()
            if s_cost_all > 0:
                d_all_reach_dict[s_name] = (
                    s_group["D0达标率"] * s_group["花费"]
                ).sum() / s_cost_all
            else:
                d_all_reach_dict[s_name] = 0.0

        show_agg = (
            df_today_shows.groupby(
                [
                    "剧目/书籍名称",
                    "剧目/书籍中文名称",
                    "剧目/书籍短ID",
                    "版权来源",
                ]
            )
            .apply(
                lambda x: pd.Series(
                    {
                        "花费": x["花费"].sum(),
                        "达标率": (x["D0达标率"] * x["花费"]).sum()
                        / (x["花费"].sum() or 1),
                        "D0 ROAS": (x["d0_roas"] * x["花费"]).sum()
                        / (x["花费"].sum() or 1),
                        "CPI": (x["CPI"] * x["花费"]).sum()
                        / (x["花费"].sum() or 1),
                        "首日付费率": (x["首日付费率"] * x["花费"]).sum()
                        / (x["花费"].sum() or 1),
                        "CPP": (x["首日付费用户成本"] * x["花费"]).sum()
                        / (x["花费"].sum() or 1),
                        "ARPPU": (
                            x["Arppu- 付费用户人均付费金额"] * x["花费"]
                        ).sum()
                        / (x["花费"].sum() or 1),
                        "新注册占比": (
                            x["新注册用户占比"] * x["花费"]
                        ).sum()
                        / (x["花费"].sum() or 1),
                        "付费新注册占比": (
                            x["付费用户新注册用户占比"] * x["花费"]
                        ).sum()
                        / (x["花费"].sum() or 1),
                    }
                )
            )
            .reset_index()
            .sort_values(by="花费", ascending=False)
        )

        html_t2 = """<table class="dataframe-table">
            <thead>
                <tr>
                    <th>序号</th>
                    <th>剧目名称</th>
                    <th>中文推广剧目</th>
                    <th>剧目ID</th>
                    <th>版权来源</th>
                    <th>花费 ($)</th>
                    <th>达标率 (%)</th>
                    <th>全期汇总达标率 (%)</th>
                    <th>D0 ROAS (%)</th>
                    <th>CPI ($)</th>
                    <th>首日付费率 (%)</th>
                    <th>CPP ($)</th>
                    <th>ARPPU ($)</th>
                    <th>新注册占比 (%)</th>
                    <th>付费新注册占比 (%)</th>
                </tr>
            </thead>
            <tbody>"""

        for idx, (_, r) in enumerate(show_agg.iterrows(), start=1):
            d0_cell = (
                'class="cell-highlight-green"'
                if r["达标率"] >= 0.85
                else ""
            )
            d_all_val = d_all_reach_dict.get(r["剧目/书籍名称"], 0.0)
            d_all_cell = 'class="cell-highlight-pink"'

            html_t2 += f"""
            <tr>
                <td>{idx}</td>
                <td>{r['剧目/书籍名称']}</td>
                <td>{r['剧目/书籍中文名称']}</td>
                <td>{r['剧目/书籍短ID']}</td>
                <td>{r['版权来源']}</td>
                <td>${r['花费']:,.2f}</td>
                <td {d0_cell}>{r['达标率']*100:.2f}%</td>
                <td {d_all_cell}>{d_all_val*100:.2f}%</td>
                <td>{r['D0 ROAS']*100:.2f}%</td>
                <td>${r['CPI']:.2f}</td>
                <td>{r['首日付费率']*100:.2f}%</td>
                <td>${r['CPP']:.2f}</td>
                <td>${r['ARPPU']:.2f}</td>
                <td>{r['新注册占比']*100:.2f}%</td>
                <td>{r['付费新注册占比']*100:.2f}%</td>
            </tr>"""

        html_t2 += "</tbody></table>"
        st.markdown(html_t2, unsafe_allow_html=True)

    # -------------------------------------------------------------------------
    # 表格 3：📈 具体投放剧目每日数据情况表
    # -------------------------------------------------------------------------
    st.subheader(
        "📈 表格 3: 具体投放剧目每日数据情况表"
    )

    if not df_filtered.empty:
        show_rank = (
            df_filtered.groupby(
                [
                    "剧目/书籍名称",
                    "剧目/书籍中文名称",
                    "剧目/书籍短ID",
                ]
            )["花费"]
            .sum()
            .reset_index()
            .sort_values(by="花费", ascending=False)
        )

        html_t3 = """<table class="dataframe-table">
            <thead>
                <tr>
                    <th>序号</th>
                    <th>剧目名称</th>
                    <th>中文推广剧目</th>
                    <th>剧目ID</th>
                    <th>日期</th>
                    <th>花费 ($)</th>
                    <th>达标率 (%)</th>
                    <th>D0 ROAS (%)</th>
                    <th>CPI ($)</th>
                    <th>首日付费率 (%)</th>
                    <th>CPP ($)</th>
                    <th>ARPPU ($)</th>
                    <th>新注册占比 (%)</th>
                    <th>付费新注册占比 (%)</th>
                    <th>消耗趋势</th>
                    <th>达标率趋势</th>
                    <th>ROAS 趋势</th>
                    <th>CPI 趋势</th>
                </tr>
            </thead>
            <tbody>"""

        for rank_idx, (_, s_row) in enumerate(show_rank.iterrows(), start=1):
            s_name = s_row["剧目/书籍名称"]
            s_zh = s_row["剧目/书籍中文名称"]
            s_id = s_row["剧目/书籍短ID"]

            s_df = df_filtered[
                df_filtered["剧目/书籍名称"] == s_name
            ].copy()

            s_daily = (
                s_df.groupby("日期_str")
                .apply(
                    lambda x: pd.Series(
                        {
                            "花费": x["花费"].sum(),
                            "达标率": (x["D0达标率"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "D0 ROAS": (x["d0_roas"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "CPI": (x["CPI"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "首日付费率": (x["首日付费率"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "CPP": (x["首日付费用户成本"] * x["花费"]).sum()
                            / (x["花费"].sum() or 1),
                            "ARPPU": (
                                x["Arppu- 付费用户人均付费金额"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "新注册占比": (
                                x["新注册用户占比"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                            "付费新注册占比": (
                                x["付费用户新注册用户占比"] * x["花费"]
                            ).sum()
                            / (x["花费"].sum() or 1),
                        }
                    )
                )
                .reset_index()
            )

            s_daily = s_daily.sort_values(by="日期_str", ascending=False)
            s_dates_list = s_daily["日期_str"].tolist()
            num_rows = len(s_daily) + 1

            for idx, (_, r) in enumerate(s_daily.iterrows()):
                cur_date = r["日期_str"]
                cur_idx = s_dates_list.index(cur_date)
                prev_row = None
                if cur_idx + 1 < len(s_dates_list):
                    p_date = s_dates_list[cur_idx + 1]
                    p_match = s_daily[s_daily["日期_str"] == p_date]
                    if not p_match.empty:
                        prev_row = p_match.iloc[0]

                cost_trend = (
                    get_trend_str(r["花费"], prev_row["花费"])
                    if prev_row is not None
                    else "-"
                )
                d0_trend = (
                    get_trend_str(r["达标率"], prev_row["达标率"])
                    if prev_row is not None
                    else "-"
                )
                roas_trend = (
                    get_trend_str(r["D0 ROAS"], prev_row["D0 ROAS"])
                    if prev_row is not None
                    else "-"
                )
                cpi_trend = (
                    get_trend_str(r["CPI"], prev_row["CPI"], is_cost=True)
                    if prev_row is not None
                    else "-"
                )

                row_class = (
                    'class="row-muted"' if cur_date != latest_date else ""
                )
                d0_cell = (
                    'class="cell-highlight-green"'
                    if r["达标率"] >= 0.85
                    else ""
                )

                html_t3 += f"<tr {row_class}>"
                if idx == 0:
                    html_t3 += f'<td rowspan="{num_rows}">{rank_idx}</td>'
                    html_t3 += f'<td rowspan="{num_rows}">{s_name}</td>'
                    html_t3 += f'<td rowspan="{num_rows}">{s_zh}</td>'
                    html_t3 += f'<td rowspan="{num_rows}">{s_id}</td>'

                html_t3 += f"""
                    <td>{cur_date}</td>
                    <td>${r['花费']:,.2f}</td>
                    <td {d0_cell}>{r['达标率']*100:.2f}%</td>
                    <td>{r['D0 ROAS']*100:.2f}%</td>
                    <td>${r['CPI']:.2f}</td>
                    <td>{r['首日付费率']*100:.2f}%</td>
                    <td>${r['CPP']:.2f}</td>
                    <td>${r['ARPPU']:.2f}</td>
                    <td>{r['新注册占比']*100:.2f}%</td>
                    <td>{r['付费新注册占比']*100:.2f}%</td>
                    <td class="col-trend">{cost_trend}</td>
                    <td class="col-trend">{d0_trend}</td>
                    <td class="col-trend">{roas_trend}</td>
                    <td class="col-trend">{cpi_trend}</td>
                </tr>"""

            s_cost_total = s_df["花费"].sum()
            s_d0_3d = (s_df["D0达标率"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_roas_3d = (s_df["d0_roas"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_cpi_3d = (s_df["CPI"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_pay_3d = (s_df["首日付费率"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_cpp_3d = (s_df["首日付费用户成本"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_arppu_3d = (
                s_df["Arppu- 付费用户人均付费金额"] * s_df["花费"]
            ).sum() / (s_cost_total or 1)
            s_reg_3d = (s_df["新注册用户占比"] * s_df["花费"]).sum() / (
                s_cost_total or 1
            )
            s_preg_3d = (
                s_df["付费用户新注册用户占比"] * s_df["花费"]
            ).sum() / (s_cost_total or 1)

            tot_d0_cell = (
                'class="cell-highlight-green"' if s_d0_3d >= 0.85 else ""
            )

            html_t3 += f"""
            <tr class="row-subtotal">
                <td>剧目全期汇总</td>
                <td>${s_cost_total:,.2f}</td>
                <td {tot_d0_cell}>{s_d0_3d*100:.2f}%</td>
                <td>{s_roas_3d*100:.2f}%</td>
                <td>${s_cpi_3d:.2f}</td>
                <td>{s_pay_3d*100:.2f}%</td>
                <td>${s_cpp_3d:.2f}</td>
                <td>${s_arppu_3d:.2f}</td>
                <td>{s_reg_3d*100:.2f}%</td>
                <td>{s_preg_3d*100:.2f}%</td>
                <td class="col-trend" colspan="4">-</td>
            </tr>"""

        html_t3 += "</tbody></table>"
        st.markdown(html_t3, unsafe_allow_html=True)

    st.markdown("---")

    # =========================================================================
    # 第三部分：投放维度拆分
    # =========================================================================
    st.header("📌 第三部分：投放维度拆分")

    # -------------------------------------------------------------------------
    # 第三部分双时间窗 (Time Window Tabs) 数据隔离机制
    # -------------------------------------------------------------------------
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

        # ---------------------------------------------------------------------
        # 板块 1: 渠道维度情况 (FB / TT)
        # ---------------------------------------------------------------------
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
                ["优化师", "剧目/书籍语言", "投放平台"],
                c1_metrics,
                total_title=total_title,
            )
            st.markdown(html_c1, unsafe_allow_html=True)

        st.markdown(
            f"##### 📌 表格 2: 投放平台每日数据趋势表 {suffix_title}"
        )
        if not df_p3_working.empty:
            html_c2 = """<table class="dataframe-table">
                <thead>
                    <tr>
                        <th>优化师</th>
                        <th>语种</th>
                        <th>投放平台</th>
                        <th>日期</th>
                        <th>花费 ($)</th>
                        <th>达标率 (%)</th>
                        <th>D0 ROAS (%)</th>
                        <th>CPI ($)</th>
                        <th>CPP ($)</th>
                        <th>新注册用户占比 (%)</th>
                        <th>付费用户新注册用户占比 (%)</th>
                        <th>消耗趋势</th>
                        <th>达标率趋势</th>
                        <th>ROAS 趋势</th>
                        <th>CPI 趋势</th>
                    </tr>
                </thead>
                <tbody>"""

            total_cost_p2 = df_p3_working["花费"].sum()
            if total_cost_p2 > 0:
                d0_all = (
                    df_p3_working["D0达标率"] * df_p3_working["花费"]
                ).sum() / total_cost_p2
                roas_all = (
                    df_p3_working["d0_roas"] * df_p3_working["花费"]
                ).sum() / total_cost_p2
                cpi_all = (
                    df_p3_working["CPI"] * df_p3_working["花费"]
                ).sum() / total_cost_p2
                cpp_all = (
                    df_p3_working["首日付费用户成本"]
                    * df_p3_working["花费"]
                ).sum() / total_cost_p2
                reg_all = (
                    df_p3_working["新注册用户占比"] * df_p3_working["花费"]
                ).sum() / total_cost_p2
                preg_all = (
                    df_p3_working["付费用户新注册用户占比"]
                    * df_p3_working["花费"]
                ).sum() / total_cost_p2
                d0_cell = (
                    'class="cell-highlight-green"' if d0_all >= 0.85 else ""
                )

                html_c2 += f"""
                <tr class="row-total">
                    <td colspan="4">{total_title}</td>
                    <td>${total_cost_p2:,.2f}</td>
                    <td {d0_cell}>{d0_all*100:.2f}%</td>
                    <td>{roas_all*100:.2f}%</td>
                    <td>${cpi_all:.2f}</td>
                    <td>${cpp_all:.2f}</td>
                    <td>{reg_all*100:.2f}%</td>
                    <td>{preg_all*100:.2f}%</td>
                    <td class="col-trend" colspan="4">-</td>
                </tr>"""

            opt_sorted = (
                df_p3_working.groupby("优化师")["花费"]
                .sum()
                .reset_index()
                .sort_values(by="花费", ascending=False)
            )

            for _, opt_r in opt_sorted.iterrows():
                opt = opt_r["优化师"]
                g_opt = df_p3_working[df_p3_working["优化师"] == opt]

                for lang, g_lang in g_opt.groupby("剧目/书籍语言"):
                    for plat, g_plat in g_lang.groupby("投放平台"):
                        p_daily = (
                            g_plat.groupby("日期_str")
                            .apply(
                                lambda x: pd.Series(
                                    {
                                        "花费": x["花费"].sum(),
                                        "达标率": (
                                            x["D0达标率"] * x["花费"]
                                        ).sum()
                                        / (x["花费"].sum() or 1),
                                        "D0 ROAS": (
                                            x["d0_roas"] * x["花费"]
                                        ).sum()
                                        / (x["花费"].sum() or 1),
                                        "CPI": (x["CPI"] * x["花费"]).sum()
                                        / (x["花费"].sum() or 1),
                                        "CPP": (
                                            x["首日付费用户成本"] * x["花费"]
                                        ).sum()
                                        / (x["花费"].sum() or 1),
                                        "新注册用户占比": (
                                            x["新注册用户占比"] * x["花费"]
                                        ).sum()
                                        / (x["花费"].sum() or 1),
                                        "付费用户新注册用户占比": (
                                            x["付费用户新注册用户占比"]
                                            * x["花费"]
                                        ).sum()
                                        / (x["花费"].sum() or 1),
                                    }
                                )
                            )
                            .reset_index()
                            .sort_values(by="日期_str", ascending=False)
                        )

                        p_dates_list = p_daily["日期_str"].tolist()
                        num_rows = len(p_daily)
                        for idx, (_, r) in enumerate(p_daily.iterrows()):
                            cur_date = r["日期_str"]
                            cur_idx = p_dates_list.index(cur_date)
                            prev_row = None
                            if cur_idx + 1 < len(p_dates_list):
                                p_date = p_dates_list[cur_idx + 1]
                                p_match = p_daily[
                                    p_daily["日期_str"] == p_date
                                ]
                                if not p_match.empty:
                                    prev_row = p_match.iloc[0]

                            cost_trend = (
                                get_trend_str(r["花费"], prev_row["花费"])
                                if prev_row is not None
                                else "-"
                            )
                            d0_trend = (
                                get_trend_str(
                                    r["达标率"], prev_row["达标率"]
                                )
                                if prev_row is not None
                                else "-"
                            )
                            roas_trend = (
                                get_trend_str(
                                    r["D0 ROAS"], prev_row["D0 ROAS"]
                                )
                                if prev_row is not None
                                else "-"
                            )
                            cpi_trend = (
                                get_trend_str(
                                    r["CPI"], prev_row["CPI"], is_cost=True
                                )
                                if prev_row is not None
                                else "-"
                            )

                            row_class = (
                                'class="row-muted"'
                                if cur_date != latest_date
                                else ""
                            )
                            d0_cell = (
                                'class="cell-highlight-green"'
                                if r["达标率"] >= 0.85
                                else ""
                            )

                            html_c2 += f"<tr {row_class}>"
                            if idx == 0:
                                html_c2 += (
                                    f'<td rowspan="{num_rows}">{opt}</td>'
                                )
                                html_c2 += (
                                    f'<td rowspan="{num_rows}">{lang}</td>'
                                )
                                html_c2 += (
                                    f'<td rowspan="{num_rows}">{plat}</td>'
                                )

                            html_c2 += f"""
                                <td>{cur_date}</td>
                                <td>${r['花费']:,.2f}</td>
                                <td {d0_cell}>{r['达标率']*100:.2f}%</td>
                                <td>{r['D0 ROAS']*100:.2f}%</td>
                                <td>${r['CPI']:.2f}</td>
                                <td>${r['CPP']:.2f}</td>
                                <td>{r['新注册用户占比']*100:.2f}%</td>
                                <td>{r['付费用户新注册用户占比']*100:.2f}%</td>
                                <td class="col-trend">{cost_trend}</td>
                                <td class="col-trend">{d0_trend}</td>
                                <td class="col-trend">{roas_trend}</td>
                                <td class="col-trend">{cpi_trend}</td>
                            </tr>"""

            html_c2 += "</tbody></table>"
            st.markdown(html_c2, unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 板块 2: 优化方式维度情况
        # ---------------------------------------------------------------------
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

        # 第一张图指示：表格 2 字段与透视维度重构修正
        st.markdown(
            f"##### 📌 表格 2: 优化师-剧目 - 优化方式 & 出价透视表 {suffix_title}"
        )
        if not df_p3_working.empty:
            # 修正后透视维度顺序：优化师 -> 剧目/书籍语言 -> 剧目/书籍名称 -> 剧目/书籍短ID -> 投放平台 -> 操作系统 (广告) -> 推广目标 -> 优化方式 -> 手动出价的金额/系数
            o2_dims = [
                "优化师",
                "剧目/书籍语言",
                "剧目/书籍名称",
                "剧目/书籍短ID",
                "投放平台",
                "操作系统 (广告)",
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
                df_p3_working,
                o2_dims,
                o2_metrics,
                total_title=total_title,
            )
            st.markdown(html_o2, unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 板块 3: 投放平台 & Pixel 名称数据维度情况
        # ---------------------------------------------------------------------
        st.subheader("🎯 板块 3: 投放平台 & Pixel 名称数据维度情况")

        st.markdown(f"##### 📌 Pixel 维度明细汇总表 {suffix_title}")
        if not df_p3_working.empty:
            pix_dims = [
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
            html_pix = generate_pivot_html(
                df_p3_working,
                pix_dims,
                pix_metrics,
                total_title=total_title,
            )
            st.markdown(html_pix, unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 板块 4: 不同语种的账户时区数据透视表
        # ---------------------------------------------------------------------
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
                    ["账户时区分类"],
                    tz_metrics,
                    total_title=f"{lang} 小计",
                )
                st.markdown(html_tz, unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # 板块 5: 当日新老广告效果对比 (精确动态日期计算定义)
        # ---------------------------------------------------------------------
        st.subheader(
            "⏱️ 板块 5: 当日新老广告效果对比 (按每日拆分展平)"
        )

        if not df_p3_working.empty:
            df_p5 = df_p3_working.copy()

            # 精准新老广告定义：根据【当前计算日期 (Date)】与【广告创建时间】动态判定
            def get_dynamic_ad_age(row):
                try:
                    cur_date = pd.to_datetime(row["日期_str"])
                    create_date = pd.to_datetime(row["广告创建时间"])
                    diff_days = (cur_date - create_date).days

                    # 当日新建广告：创建时间为 Date 或 Date - 1日
                    if 0 <= diff_days <= 1:
                        return "当日新建广告"
                    # 老广告：创建时间为 Date - 2日 及以前
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

        # ---------------------------------------------------------------------
        # 板块 6: 不同操作系统数据维度情况
        # ---------------------------------------------------------------------
        st.subheader("💻 板块 6: 不同操作系统数据维度情况")

        st.markdown(
            f"##### 📌 表格 1: 操作系统数据汇总 {suffix_title}"
        )
        if not df_p3_working.empty:
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
                df_p3_working,
                os1_dims,
                os1_metrics,
                total_title=total_title,
            )
            st.markdown(html_os1, unsafe_allow_html=True)

        st.markdown(
            f"##### 📌 表格 2: 操作系统 & 广告优化方式数据汇总 {suffix_title}"
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
            html_os2 = generate_pivot_html(
                df_p3_working,
                os2_dims,
                os2_metrics,
                total_title=total_title,
            )
            st.markdown(html_os2, unsafe_allow_html=True)

    # 选项卡 1：📅 当日数据 (最新一日) -> 强行过滤为最新一天数据
    with part3_tab1:
        df_p3_daily = df_filtered[
            df_filtered["日期_str"] == latest_date
        ].copy()
        render_part3_content(df_p3_daily, is_daily_mode=True)

    # 选项卡 2：🌐 全量总计 (所有时间汇总) -> 底表全量数据
    with part3_tab2:
        render_part3_content(df_filtered, is_daily_mode=False)