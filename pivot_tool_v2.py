import streamlit as st
import pandas as pd
import numpy as np
import datetime

# ==========================================
# 0. 页面基础配置与 CSS 跨行合并/高亮样式
# ==========================================
st.set_page_config(
    page_title="短剧/漫剧数据分析工具 v2",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 注入自定义 CSS：样式、高亮、阅读模式交叉高亮
st.markdown("""
<style>
    /* 全局表格基础样式 */
    .pivot-table {
        width: 100%;
        border-collapse: collapse;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        font-size: 13px;
        color: #333333;
        margin-bottom: 20px;
    }
    .pivot-table th {
        background-color: #f8f9fa;
        color: #212529;
        font-weight: 600;
        text-align: center;
        padding: 8px 10px;
        border: 1px solid #dee2e6;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    .pivot-table td {
        padding: 6px 10px;
        border: 1px solid #dee2e6;
        text-align: center;
        white-space: nowrap;
    }
    
    /* 汇总行样式 */
    .subtotal-row {
        background-color: #FFF9C4 !important; /* 浅黄色小计 */
        font-weight: bold;
    }
    .grandtotal-row {
        background-color: #FFE066 !important; /* 暖黄色大盘总计 */
        font-weight: bold;
    }
    
    /* 单元格特定高亮 */
    .highlight-green {
        background-color: #D4EDDA !important; /* 达标率 >= 85% 淡绿 */
        color: #155724;
        font-weight: bold;
    }
    .highlight-pink {
        background-color: #FFE6E6 !important; /* 达标率列标题/汇总粉色 */
        color: #721c24;
    }
    .text-grey {
        color: #8c8c8c !important; /* 历史日期文本中灰视觉降权 */
    }
    .bg-trend {
        background-color: #F2F2F2 !important; /* 趋势列浅灰背景 */
    }

    /* 阅读模式：鼠标悬停行/列交叉高亮 */
    .pivot-table tbody tr:hover {
        background-color: #E8F0FE !important;
    }
    .pivot-table td:hover {
        background-color: #D2E3FC !important;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# 1. 辅助函数： HTML 跨行合并 (Rowspan) 渲染器
# ==========================================
def render_html_pivot_table(df, dimension_cols, metric_cols, 
                            subtotal_level=None, 
                            highlight_reach_rate=True,
                            grey_historical_dates=False,
                            date_col="日期",
                            today_date=None):
    """
    将 Pandas DataFrame 转换为支持 HTML rowspan 跨行合并的 HTML 表格。
    支持小计行、大盘总计行、达标率单元格独立高亮、历史日期降权。
    """
    if df.empty:
        return "<p>⚠️ 暂无满足条件的数据</p>"

    # 1. 计算维度列的 Rowspan 结构
    n_rows = len(df)
    n_dims = len(dimension_cols)
    rowspan_matrix = np.ones((n_rows, n_dims), dtype=int)
    show_matrix = np.ones((n_rows, n_dims), dtype=bool)

    for j in range(n_dims):
        i = 0
        while i < n_rows:
            if df.iloc[i].get('is_subtotal', False) or df.iloc[i].get('is_grandtotal', False):
                i += 1
                continue
                
            span = 1
            while i + span < n_rows:
                same_parent = True
                if j > 0:
                    same_parent = (df.iloc[i:i+span+1, :j].nunique().max() == 1)
                
                current_same = (df.iloc[i, j] == df.iloc[i+span, j])
                not_summary = not (df.iloc[i+span].get('is_subtotal', False) or df.iloc[i+span].get('is_grandtotal', False))
                
                if same_parent and current_same and not_summary:
                    span += 1
                else:
                    break
            
            rowspan_matrix[i, j] = span
            for k in range(1, span):
                show_matrix[i+k, j] = False
            i += span

    # 2. 构建 HTML 字符串
    html = ['<table class="pivot-table"><thead><tr>']
    
    all_headers = dimension_cols + metric_cols
    for header in all_headers:
        if "达标率" in header:
            html.append(f'<th class="highlight-pink">{header}</th>')
        elif "趋势" in header:
            html.append(f'<th class="bg-trend">{header}</th>')
        else:
            html.append(f'<th>{header}</th>')
    html.append('</tr></thead><tbody>')

    for i in range(n_rows):
        row = df.iloc[i]
        is_sub = row.get('is_subtotal', False)
        is_grand = row.get('is_grandtotal', False)
        
        row_class = ""
        if is_grand:
            row_class = ' class="grandtotal-row"'
        elif is_sub:
            row_class = ' class="subtotal-row"'
            
        html.append(f'<tr{row_class}>')

        if is_grand:
            html.append(f'<td colspan="{n_dims}"><b>全盘大盘总计</b></td>')
        elif is_sub:
            sub_name = row.get('subtotal_name', '小计')
            html.append(f'<td colspan="{n_dims}"><b>{sub_name}</b></td>')
        else:
            for j in range(n_dims):
                if show_matrix[i, j]:
                    r_span = rowspan_matrix[i, j]
                    span_attr = f' rowspan="{r_span}"' if r_span > 1 else ''
                    val = row[dimension_cols[j]]
                    val_str = "" if pd.isna(val) else str(val)
                    
                    cell_class = ""
                    if grey_historical_dates and dimension_cols[j] == date_col and today_date:
                        if str(val) != str(today_date):
                            cell_class = ' class="text-grey"'
                            
                    html.append(f'<td{span_attr}{cell_class}>{val_str}</td>')

        for m_col in metric_cols:
            val = row.get(m_col, np.nan)
            val_str = "-"
            cell_class = ""

            if pd.notna(val) and val != "-":
                if any(k in m_col for k in ["%", "率", "ROAS", "占比"]):
                    val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                elif any(k in m_col for k in ["花费", "CPI", "CPP", "金额", "消耗"]):
                    val_str = f"${val:,.2f}" if isinstance(val, (int, float)) else str(val)
                elif isinstance(val, (int, float)):
                    val_str = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
                else:
                    val_str = str(val)

            if highlight_reach_rate and "达标率" in m_col and not is_sub and not is_grand:
                try:
                    num_val = float(val)
                    if num_val >= 85.0:
                        cell_class = ' class="highlight-green"'
                except (ValueError, TypeError):
                    pass
                    
            if "趋势" in m_col:
                cell_class = ' class="bg-trend"'

            html.append(f'<td{cell_class}>{val_str}</td>')

        html.append('</tr>')

    html.append('tbody></table>')
    return "".join(html)


# ==========================================
# 2. 核心数据计算引擎（衍生指标与数据清洗）
# ==========================================
def process_raw_data(df):
    """基础数据清洗与衍生指标计算"""
    if '花费' in df.columns:
        df = df[df['花费'] > 0].copy()
    elif '消耗' in df.columns:
        df = df[df['消耗'] > 0].copy()
        df['花费'] = df['消耗']

    if '日期' in df.columns:
        df['日期'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
    if '广告创建时间' in df.columns:
        df['广告创建时间'] = pd.to_datetime(df['广告创建时间']).dt.strftime('%Y-%m-%d')
    elif '系列创建时间' in df.columns:
        df['广告创建时间'] = pd.to_datetime(df['系列创建时间']).dt.strftime('%Y-%m-%d')

    numeric_cols = ['花费', '回收', '展现量', '点击量', '安装量', '首日付费人数', '首日付费金额', '注册用户数', '付费用户数', '新注册用户数']
    for c in numeric_cols:
        if c not in df.columns:
            df[c] = 0.0
        else:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

    df['D0 ROAS (%)'] = np.where(df['花费'] > 0, (df['回收'] / df['花费']) * 100, 0)
    df['CPI ($)'] = np.where(df['安装量'] > 0, df['花费'] / df['安装量'], 0)
    df['CPP ($)'] = np.where(df['首日付费人数'] > 0, df['花费'] / df['首日付费人数'], 0)
    df['首日付费率 (%)'] = np.where(df['安装量'] > 0, (df['首日付费人数'] / df['安装量']) * 100, 0)
    
    if '目标ROAS' not in df.columns:
        df['目标ROAS'] = 100.0
    df['达标率 (%)'] = np.where(df['目标ROAS'] > 0, (df['D0 ROAS (%)'] / df['目标ROAS']) * 100, 0)

    return df


def calculate_summary_metrics(df, group_cols):
    """按分组聚合汇总计算"""
    grouped = df.groupby(group_cols, as_index=False).agg({
        '花费': 'sum',
        '回收': 'sum',
        '安装量': 'sum',
        '首日付费人数': 'sum',
        '首日付费金额': 'sum',
        '注册用户数': 'sum',
        '付费用户数': 'sum',
        '新注册用户数': 'sum'
    })
    
    grouped['D0 ROAS (%)'] = np.where(grouped['花费'] > 0, (grouped['回收'] / grouped['花费']) * 100, 0)
    grouped['CPI ($)'] = np.where(grouped['安装量'] > 0, grouped['花费'] / grouped['安装量'], 0)
    grouped['CPP ($)'] = np.where(grouped['首日付费人数'] > 0, grouped['花费'] / grouped['首日付费人数'], 0)
    grouped['首日付费率 (%)'] = np.where(grouped['安装量'] > 0, (grouped['首日付费人数'] / grouped['安装量']) * 100, 0)
    grouped['新注册用户占比 (%)'] = np.where(grouped['注册用户数'] > 0, (grouped['新注册用户数'] / grouped['注册用户数']) * 100, 0)
    grouped['付费用户新注册用户占比 (%)'] = np.where(grouped['付费用户数'] > 0, (grouped['新注册用户数'] / grouped['付费用户数']) * 100, 0)
    
    target_roas = df.groupby(group_cols)['目标ROAS'].mean().reset_index()['目标ROAS']
    grouped['达标率 (%)'] = np.where(target_roas > 0, (grouped['D0 ROAS (%)'] / target_roas) * 100, 0)

    grouped = grouped.sort_values(by='花费', ascending=False).reset_index(drop=True)
    return grouped


# ==========================================
# 3. Streamlit 主页面应用入口与数据加载
# ==========================================
st.title("📊 短剧/漫剧数据透视与投放分析工具 (新板重构)")
st.markdown("— *全量每日明细透视、维度重构与单元格深度高亮*")

st.sidebar.header("📁 数据导入与过滤")
uploaded_file = st.sidebar.file_uploader("上传投放数据 Excel / CSV", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            raw_df = pd.read_csv(uploaded_file)
        else:
            raw_df = pd.read_excel(uploaded_file)
            
        df_clean = process_raw_data(raw_df)
        
        st.sidebar.subheader("🎯 维度筛选")
        all_languages = df_clean['语种'].unique().tolist() if '语种' in df_clean.columns else []
        sel_languages = st.sidebar.multiselect("选择语种", options=all_languages, default=all_languages)
        
        all_opt_methods = df_clean['优化方式'].unique().tolist() if '优化方式' in df_clean.columns else []
        sel_opt_methods = st.sidebar.multiselect("选择优化方式", options=all_opt_methods, default=all_opt_methods)
        
        df_filtered = df_clean.copy()
        if sel_languages:
            df_filtered = df_filtered[df_filtered['语种'].isin(sel_languages)]
        if sel_opt_methods:
            df_filtered = df_filtered[df_filtered['优化方式'].isin(sel_opt_methods)]

        all_dates = sorted(df_filtered['日期'].unique().tolist(), reverse=True)
        today_date = all_dates[0] if all_dates else None

        # ==========================================
        # 4. 第一部分：综合概述与优化师趋势
        # ==========================================
        st.header("第一部分：综合概况与趋势透视")
        
        with st.expander("📊 2. 不同优化师、不同语种每日数据情况与趋势表", expanded=True):
            st.markdown("📌 *全量日期纵向展平，跨行合并优化师与语种，含前一日（T-1）趋势对比及优化师全期汇总行*")
            
            dim_cols = ['优化师', '语种', '日期']
            opt_df = calculate_summary_metrics(df_filtered, dim_cols)
            opt_df = opt_df.sort_values(by=['优化师', '语种', '日期'], ascending=[True, True, False]).reset_index(drop=True)
            
            final_rows = []
            for (opt, lang), group in opt_df.groupby(['优化师', '语种']):
                group = group.sort_values(by='日期', ascending=True)
                group['花费趋势'] = group['花费'].diff().apply(lambda x: f"{x:+.2f}" if pd.notna(x) else "-")
                group['ROAS趋势'] = group['D0 ROAS (%)'].diff().apply(lambda x: f"{x:+.2f}%" if pd.notna(x) else "-")
                group = group.sort_values(by='日期', ascending=False)
                
                for _, r in group.iterrows():
                    final_rows.append(r.to_dict())
                
                sub_row = {
                    '优化师': opt,
                    '语种': lang,
                    '日期': '全期小计',
                    '花费': group['花费'].sum(),
                    '回收': group['回收'].sum(),
                    '安装量': group['安装量'].sum(),
                    '首日付费人数': group['首日付费人数'].sum(),
                    'D0 ROAS (%)': (group['回收'].sum() / group['花费'].sum() * 100) if group['花费'].sum() > 0 else 0,
                    'CPI ($)': (group['花费'].sum() / group['安装量'].sum()) if group['安装量'].sum() > 0 else 0,
                    '达标率 (%)': group['达标率 (%)'].mean(),
                    'is_subtotal': True,
                    'subtotal_name': f'{opt} ({lang}) - 全期汇总'
                }
                final_rows.append(sub_row)
                
            df_opt_final = pd.DataFrame(final_rows)
            metric_list = ['花费', '回收', 'D0 ROAS (%)', '达标率 (%)', 'CPI ($)', '首日付费人数', '花费趋势', 'ROAS趋势']
            
            html_code = render_html_pivot_table(
                df_opt_final, dim_cols, metric_list, 
                highlight_reach_rate=True, 
                grey_historical_dates=True, 
                today_date=today_date
            )
            st.markdown(html_code, unsafe_allow_html=True)


        # ==========================================
        # 5. 第二部分：剧目明细透视表
        # ==========================================
        st.header("第二部分：剧目维度深度透视")
        
        with st.expander("📈 表格 3: 具体投放剧目每日数据情况表", expanded=True):
            st.markdown("📌 *按全量日期纵向展开，合并剧目信息；仅 Today 数据常规高亮，历史日期文字颜色调浅 `#8c8c8c` 视觉降权*")
            
            drama_dims = ['序号', '剧目名称', '中文推广剧目', '剧目ID', '日期']
            if '剧目名称' in df_filtered.columns:
                unique_dramas = df_filtered[['剧目名称']].drop_duplicates().reset_index(drop=True)
                unique_dramas['序号'] = unique_dramas.index + 1
                df_drama_base = pd.merge(df_filtered, unique_dramas, on='剧目名称', how='left')
            else:
                df_drama_base = df_filtered.copy()
                df_drama_base['序号'] = 1
                df_drama_base['剧目名称'] = '未知剧目'
                df_drama_base['中文推广剧目'] = '-'
                df_drama_base['剧目ID'] = '-'

            drama_df = calculate_summary_metrics(df_drama_base, drama_dims)
            drama_df = drama_df.sort_values(by=['序号', '日期'], ascending=[True, False]).reset_index(drop=True)
            
            drama_rows = []
            for drama_name, group in drama_df.groupby('剧目名称'):
                for _, r in group.iterrows():
                    drama_rows.append(r.to_dict())
                
                sub_row = {
                    '序号': group['序号'].iloc[0],
                    '剧目名称': drama_name,
                    '中文推广剧目': group['中文推广剧目'].iloc[0],
                    '剧目ID': group['剧目ID'].iloc[0],
                    '日期': '剧目全期汇总',
                    '花费': group['花费'].sum(),
                    '回收': group['回收'].sum(),
                    '安装量': group['安装量'].sum(),
                    'D0 ROAS (%)': (group['回收'].sum() / group['花费'].sum() * 100) if group['花费'].sum() > 0 else 0,
                    '达标率 (%)': group['达标率 (%)'].mean(),
                    '全期汇总达标率 (%)': group['达标率 (%)'].mean(),
                    'is_subtotal': True,
                    'subtotal_name': f'🎬 {drama_name} - 全期汇总'
                }
                drama_rows.append(sub_row)

            df_drama_final = pd.DataFrame(drama_rows)
            df_drama_final['全期汇总达标率 (%)'] = df_drama_final['达标率 (%)']
            
            drama_metrics = ['花费', '回收', 'D0 ROAS (%)', '达标率 (%)', '全期汇总达标率 (%)', 'CPI ($)', 'CPP ($)']
            html_drama = render_html_pivot_table(
                df_drama_final, drama_dims, drama_metrics,
                highlight_reach_rate=True,
                grey_historical_dates=True,
                today_date=today_date
            )
            st.markdown(html_drama, unsafe_allow_html=True)


        # ==========================================
        # 6. 第三部分：投放维度拆分深度重构
        # ==========================================
        st.header("第三部分：投放维度拆分")

        # 板块 1: 投放平台每日趋势表 (表格 2)
        with st.expander("📊 板块 1: 渠道维度情况 & 投放平台每日数据趋势表", expanded=True):
            st.subheader("表格 1: 渠道维度情况 (全量总计)")
            p1_t1_dims = ['渠道', '投放平台']
            p1_t1_df = calculate_summary_metrics(df_filtered, p1_t1_dims)
            p1_metrics = ['花费', '回收', 'D0 ROAS (%)', '达标率 (%)', 'CPI ($)', 'CPP ($)']
            st.markdown(render_html_pivot_table(p1_t1_df, p1_t1_dims, p1_metrics), unsafe_allow_html=True)

            st.subheader("表格 2: 投放平台每日数据趋势表")
            p1_t2_dims = ['优化师', '语种', '投放平台', '日期']
            p1_t2_df = calculate_summary_metrics(df_filtered, p1_t2_dims)
            p1_t2_df = p1_t2_df.sort_values(by=['优化师', '语种', '投放平台', '日期'], ascending=[True, True, True, False])
            st.markdown(render_html_pivot_table(
                p1_t2_df, p1_t2_dims, p1_metrics, 
                grey_historical_dates=True, 
                today_date=today_date
            ), unsafe_allow_html=True)

        # 板块 2: 优化方式维度情况 (全量总计)
        with st.expander("⚙️ 板块 2: 优化方式维度情况 (全量总计)", expanded=False):
            p2_dims = ['语种', '投放平台', '推广目标', '优化方式']
            p2_df = calculate_summary_metrics(df_filtered, p2_dims)
            st.markdown(render_html_pivot_table(p2_df, p2_dims, p1_metrics), unsafe_allow_html=True)

        # 板块 3: 投放平台 & Pixel 名称维度 (全量总计)
        with st.expander("🎯 板块 3: 投放平台 & Pixel 名称数据维度情况 (全量总计)", expanded=False):
            p3_dims = ['语种', '投放平台', 'Pixel名称'] if 'Pixel名称' in df_filtered.columns else ['语种', '投放平台']
            p3_df = calculate_summary_metrics(df_filtered, p3_dims)
            st.markdown(render_html_pivot_table(p3_df, p3_dims, p1_metrics), unsafe_allow_html=True)

        # 板块 4: 不同语种的账户时区数据透视表 (全量总计)
        with st.expander("🌐 板块 4: 不同语种的账户时区数据透视表 (全量总计)", expanded=False):
            p4_dims = ['语种', '账户时区'] if '账户时区' in df_filtered.columns else ['语种']
            p4_df = calculate_summary_metrics(df_filtered, p4_dims)
            st.markdown(render_html_pivot_table(p4_df, p4_dims, p1_metrics), unsafe_allow_html=True)

        # 板块 5: 当日新老广告效果对比 (按每日 Date 展平)
        with st.expander("📌 板块 5: 当日新老广告效果对比 (按每日拆分展平)", expanded=True):
            st.markdown("""
            📌 **新老广告精准定义（基于广告/系列创建时间）：**
            - **当日新建广告**：在该日期【当日（Date）】或【前一日（Date - 1日）】创建的广告。
            - **老广告**：在【前 2 日及更早（Date - 2日 及以前）】创建的广告。
            """)
            
            df_p5 = df_filtered.copy()
            if '广告创建时间' in df_p5.columns:
                df_p5['日期_dt'] = pd.to_datetime(df_p5['日期'])
                df_p5['创建_dt'] = pd.to_datetime(df_p5['广告创建时间'])
                df_p5['diff_days'] = (df_p5['日期_dt'] - df_p5['创建_dt']).dt.days
                df_p5['广告类型'] = np.where(df_p5['diff_days'] <= 1, '当日新建广告', '老广告')
            else:
                df_p5['广告类型'] = '未知分类'

            for target_lang in ['繁体中文', '韩语']:
                st.subheader(f"🌐 广告效果对比表 - {target_lang}")
                df_lang = df_p5[df_p5['语种'] == target_lang]
                
                if df_lang.empty:
                    st.info(f"暂无 {target_lang} 相关数据")
                    continue

                p5_dims = ['日期', '广告类型']
                p5_df = calculate_summary_metrics(df_lang, p5_dims)
                p5_df = p5_df.sort_values(by=['日期', '广告类型'], ascending=[False, True]).reset_index(drop=True)
                
                p5_rows = []
                for d_val, group in p5_df.groupby('日期'):
                    for _, r in group.iterrows():
                        p5_rows.append(r.to_dict())
                    sub_row = {
                        '日期': d_val,
                        '广告类型': '当日汇总',
                        '花费': group['花费'].sum(),
                        '回收': group['回收'].sum(),
                        '安装量': group['安装量'].sum(),
                        'D0 ROAS (%)': (group['回收'].sum() / group['花费'].sum() * 100) if group['花费'].sum() > 0 else 0,
                        '达标率 (%)': group['达标率 (%)'].mean(),
                        'CPI ($)': (group['花费'].sum() / group['安装量'].sum()) if group['安装量'].sum() > 0 else 0,
                        'is_subtotal': True,
                        'subtotal_name': f'📅 {d_val} - {target_lang} 小计'
                    }
                    p5_rows.append(sub_row)

                df_p5_final = pd.DataFrame(p5_rows)
                p5_metrics = ['花费', '回收', 'D0 ROAS (%)', '达标率 (%)', 'CPI ($)']
                st.markdown(render_html_pivot_table(df_p5_final, p5_dims, p5_metrics), unsafe_allow_html=True)

        # 板块 6: 不同操作系统数据维度情况 (重构)
        with st.expander("💻 板块 6: 不同操作系统数据维度情况 (维度顺序与单元格高亮重构)", expanded=True):
            st.subheader("表格 1: 操作系统维度概览 (全量总计)")
            p6_t1_dims = ['操作系统 (广告)', '操作系统 (用户)'] if all(k in df_filtered.columns for k in ['操作系统 (广告)', '操作系统 (用户)']) else ['语种']
            p6_t1_df = calculate_summary_metrics(df_filtered, p6_t1_dims)
            st.markdown(render_html_pivot_table(p6_t1_df, p6_t1_dims, p1_metrics), unsafe_allow_html=True)

            st.subheader("表格 2: 操作系统 & 广告优化方式数据汇总")
            required_p6_dims = [
                '语种', '投放平台', 'Pixel名称', '操作系统 (广告)', 
                '操作系统 (用户)', '推广目标', '优化方式', '手动出价的金额/系数'
            ]
            for col in required_p6_dims:
                if col not in df_filtered.columns:
                    df_filtered[col] = '-'

            p6_t2_df = calculate_summary_metrics(df_filtered, required_p6_dims)
            p6_t2_metrics = [
                '花费 ($)', '达标率 (%)', 'D0 ROAS (%)', 'CPI ($)', 
                '首日付费率 (%)', 'CPP ($)', '新注册用户占比 (%)', '付费用户新注册用户占比 (%)'
            ]
            p6_t2_df['花费 ($)'] = p6_t2_df['花费']
            
            html_p6_t2 = render_html_pivot_table(
                p6_t2_df, required_p6_dims, p6_t2_metrics, 
                highlight_reach_rate=True
            )
            st.markdown(html_p6_t2, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ 数据解析或代码运行出错: {str(e)}")
        st.exception(e)
else:
    st.info("💡 请在左侧边栏上传投放数据 Excel / CSV 文件以开启自动化透视分析看板。")