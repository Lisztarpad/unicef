import streamlit as st
import pandas as pd
import re
import io


@st.cache_data
def load_data(file):
    return pd.read_excel(file)


# Set page configuration
st.set_page_config(page_title="RAS Monitoring Report Analysis", layout="wide")


# ==========================================
# 🌟 新增：全局数据下钻弹窗组件
# ==========================================
@st.dialog("📋 Drill-down Details", width="large")
def show_drilldown_modal(msg, df):
    st.info(msg)
    st.dataframe(df, use_container_width=True, hide_index=True)


# ==========================================
# 辅助函数：终极正则解析器与列名映射器
# ==========================================
def parse_proc_col(raw_col):
    if isinstance(raw_col, (list, tuple)) and len(raw_col) == 2:
        status, days = str(raw_col[0]).strip(), str(raw_col[1]).strip()
        if status in ['No', 'Yes']: return status, days
        return 'Grand Total', ''
    s = str(raw_col)
    if 'Name of Agent' in s or 'Grand Total' in s: return 'Grand Total', ''
    status = 'No' if 'No' in s else ('Yes' if 'Yes' in s else 'Grand Total')
    if status in ['No', 'Yes']:
        nums = re.findall(r'\d+\.?\d*', s)
        if nums: return status, nums[-1]
    return status, ''


def parse_pub_col(raw_col, cols_order):
    s = str(raw_col)
    if 'Name of Agent' in s or 'Total' in s: return 'Total'
    for c in cols_order:
        if c in s: return c
    return 'Total'


def robust_column_matcher(pivot_df, raw_col_name):
    if raw_col_name in pivot_df.columns: return raw_col_name
    for c in pivot_df.columns:
        if str(c) == str(raw_col_name): return c
    return raw_col_name


# ==========================================
# 🌟 动态时间分组核心引擎 (主表)
# ==========================================
def apply_dynamic_grouping(df_valid, date_col, period_type, num_periods=4):
    max_date = df_valid[date_col].max()
    if period_type == 'Weekly':
        p1_start = max_date.floor('D') - pd.Timedelta(days=max_date.weekday())
        offsets = [pd.Timedelta(weeks=i) for i in range(num_periods)]

        def get_end(s):
            return s + pd.Timedelta(days=6)
    elif period_type == 'Monthly':
        p1_start = max_date.floor('D').replace(day=1)
        offsets = [pd.DateOffset(months=i) for i in range(num_periods)]

        def get_end(s):
            return s + pd.DateOffset(months=1) - pd.Timedelta(days=1)
    elif period_type == 'Quarterly':
        q_month = ((max_date.month - 1) // 3) * 3 + 1
        p1_start = max_date.floor('D').replace(month=q_month, day=1)
        offsets = [pd.DateOffset(months=3 * i) for i in range(num_periods)]

        def get_end(s):
            return s + pd.DateOffset(months=3) - pd.Timedelta(days=1)
    elif period_type == 'Yearly':
        p1_start = max_date.floor('D').replace(month=1, day=1)
        offsets = [pd.DateOffset(years=i) for i in range(num_periods)]

        def get_end(s):
            return s + pd.DateOffset(years=1) - pd.Timedelta(days=1)

    starts = [p1_start - offsets[i] for i in range(num_periods)]
    ranges = [f"{s.strftime('%Y-%m-%d')} ~ {get_end(s).strftime('%Y-%m-%d')}" for s in starts]
    range_older = f"Before {starts[-1].strftime('%Y-%m-%d')}"

    def categorize(date_val):
        for i in range(num_periods):
            if date_val >= starts[i]: return ranges[i]
        return range_older

    time_group = df_valid[date_col].apply(categorize)
    cols_order = [range_older] + ranges[::-1] + ['Total']
    return time_group, cols_order, ranges[0]


# ==========================================
# 🌟 独立的时间过滤引擎 (SLA & Proc 小表专用)
# ==========================================
def filter_current_period(df_valid, date_col, period_type):
    if df_valid.empty: return df_valid, ""
    max_date = df_valid[date_col].max()

    if period_type == 'Last 4 Weeks':
        current_week_monday = max_date.floor('D') - pd.Timedelta(days=max_date.weekday())
        start_date = current_week_monday - pd.Timedelta(weeks=3)
    elif period_type == 'Current Quarter':
        q_month = ((max_date.month - 1) // 3) * 3 + 1
        start_date = max_date.floor('D').replace(month=q_month, day=1)
    elif period_type == 'Current Year':
        start_date = max_date.floor('D').replace(month=1, day=1)
    else:
        current_week_monday = max_date.floor('D') - pd.Timedelta(days=max_date.weekday())
        start_date = current_week_monday - pd.Timedelta(weeks=3)

    df_filtered = df_valid[df_valid[date_col] >= start_date].copy()
    date_range_str = f"from {start_date.strftime('%d-%b-%Y')} to {max_date.strftime('%d-%b-%Y')}"
    return df_filtered, date_range_str


# ==========================================
# 各模块数据处理函数... (保持不变)
# ==========================================
def process_va_email_sent(df, period_type='Weekly'):
    required_columns = ["Name of Agent (VA)", "VA E-mail sent", "JPR"]
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols: return None, None, None
    df['VA E-mail sent'] = pd.to_datetime(df['VA E-mail sent'], errors='coerce')
    df_valid = df.dropna(subset=['VA E-mail sent']).copy()
    if df_valid.empty: return None, None, None
    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'VA E-mail sent', period_type)
    pivot_df = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (VA)', columns='Time Group', aggfunc='count',
                              fill_value=0, margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in pivot_df.columns]
    pivot_df = pivot_df[cols_order].reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (VA)' in c.columns: c['Name of Agent (VA)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        if (x['Name of Agent (VA)'] == 'Total').any(): c.loc[x[x['Name of Agent (VA)'] == 'Total'].index[
            0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    return pivot_df.astype(object).replace(0, "").style.apply(highlight_cells, axis=None), pivot_df, df_valid


def generate_sla_meeting_rate_table(df_valid, agent_col='Name of Agent (VA)'):
    sla_col = 'MET VA-SLA' if 'VA' in agent_col else (
        'MET LL-SLA' if 'LL' in agent_col else ('MET OC-SLA' if 'OC' in agent_col else 'MET FC-SLA'))
    if sla_col not in df_valid.columns: return None, None
    pivot = pd.pivot_table(df_valid, values='JPR', index=agent_col, columns=sla_col, aggfunc='count', fill_value=0)
    for col in ['No', 'Yes']:
        if col not in pivot.columns: pivot[col] = 0
    pivot = pivot[['No', 'Yes']]
    pivot['Grand Total'] = pivot['No'] + pivot['Yes']
    grand_no, grand_yes, grand_tot = pivot['No'].sum(), pivot['Yes'].sum(), pivot['Grand Total'].sum()
    rate_col = f"{sla_col.replace('MET ', '')} Meeting Rate"
    pivot[rate_col] = pivot['Yes'] / pivot['Grand Total']
    pivot.loc['Grand Total'] = [grand_no, grand_yes, grand_tot, grand_yes / grand_tot if grand_tot > 0 else 0]
    pivot = pivot.reset_index()

    def highlight_sla(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        c[agent_col] = 'font-weight: bold;'
        mask_no = pd.to_numeric(x['No'], errors='coerce') > 0
        c.loc[mask_no, 'No'] = 'background-color: #FFF2CC; color: #000000;'
        tot_mask = x[agent_col] == 'Grand Total'
        if tot_mask.any(): c.loc[
            x[tot_mask].index[0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    return pivot.style.apply(highlight_sla, axis=None).format(
        {'No': lambda v: "" if v == 0 else f"{int(v)}", 'Yes': lambda v: "" if v == 0 else f"{int(v)}",
         'Grand Total': lambda v: "" if v == 0 else f"{int(v)}", rate_col: lambda v: f"{v:.0%}"}), pivot


def generate_processing_time_table(df_valid, agent_col='Name of Agent (VA)', proc_col='VA-SLA'):
    sla_col = f"MET {proc_col}"
    if sla_col not in df_valid.columns or proc_col not in df_valid.columns: return None, None
    df_sub = df_valid.dropna(subset=[proc_col]).copy()
    pivot = pd.pivot_table(df_sub, values='JPR', index=agent_col, columns=[sla_col, proc_col], aggfunc='count',
                           fill_value=0)
    if 'No' in pivot.columns.levels[0] and 'Yes' in pivot.columns.levels[0]: pivot = pivot[['No', 'Yes']]
    pivot[('Grand Total', '')] = pivot.sum(axis=1)
    pivot.loc['Grand Total'] = pivot.sum()
    pivot = pivot.reset_index()
    new_cols = []
    for col in pivot.columns:
        if isinstance(col, tuple):
            status, days = str(col[0]).strip(), col[1]
            if pd.isna(days) or str(days) == '':
                new_cols.append(status)
            else:
                try:
                    new_cols.append(
                        f"{status} ({str(int(float(days))) if float(days).is_integer() else str(float(days))})")
                except:
                    new_cols.append(f"{status} ({str(days).strip()})")
        else:
            new_cols.append(str(col).strip())
    pivot.columns = new_cols

    def highlight_proc(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        c[agent_col] = 'font-weight: bold;'
        for col in x.columns:
            if col.startswith('No'):
                mask = pd.to_numeric(x[col], errors='coerce') > 0
                c.loc[mask, col] = 'background-color: #FFF2CC; color: #000000;'
        tot_mask = x[agent_col] == 'Grand Total'
        if tot_mask.any(): c.loc[
            x[tot_mask].index[0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    return pivot.style.apply(highlight_proc, axis=None).format(
        {col: lambda v: "" if v == 0 else f"{int(v)}" for col in pivot.columns if col != agent_col}), pivot


def process_va_published(df, period_type='Weekly', small_period='Last 4 Weeks'):
    required_columns = ["Name of Agent (VA)", "VA Addition Date", "JPR", "VA-SLA"]
    if [col for col in required_columns if
        col not in df.columns]: return None, None, None, None, None, None, None, None, None
    df['VA Addition Date'] = pd.to_datetime(df['VA Addition Date'], errors='coerce')
    df_valid = df.dropna(subset=['VA Addition Date']).copy()
    df_valid['VA-SLA'] = pd.to_numeric(df_valid['VA-SLA'], errors='coerce')
    if 'MET VA-SLA' in df_valid.columns: df_valid['MET VA-SLA'] = df_valid['MET VA-SLA'].astype(str).str.strip()
    if df_valid.empty: return None, None, None, None, None, None, None, None, None

    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'VA Addition Date', period_type)
    count_pivot = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (VA)', columns='Time Group',
                                 aggfunc='count', fill_value=0, margins=True, margins_name='Total')
    mean_pivot = pd.pivot_table(df_valid, values='VA-SLA', index='Name of Agent (VA)', columns='Time Group',
                                aggfunc='mean', margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in count_pivot.columns]
    count_pivot, mean_pivot = count_pivot[cols_order], mean_pivot.reindex(index=count_pivot.index, columns=cols_order)
    combined_pivot = pd.DataFrame(index=count_pivot.index, columns=cols_order)
    for col in cols_order:
        for row in count_pivot.index:
            c, m = count_pivot.at[row, col], mean_pivot.at[row, col]
            combined_pivot.at[row, col] = "" if pd.isna(c) or c == 0 else (
                f"{int(c)}" if pd.isna(m) else f"{int(c)} (Sla: {m:.1f})")

    combined_pivot = combined_pivot.reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (VA)' in c.columns: c['Name of Agent (VA)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        if (x['Name of Agent (VA)'] == 'Total').any(): c.loc[x[x['Name of Agent (VA)'] == 'Total'].index[
            0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    df_recent, date_range_str = filter_current_period(df_valid, 'VA Addition Date', small_period)
    sla_styler, sla_pivot = generate_sla_meeting_rate_table(df_recent, 'Name of Agent (VA)')
    proc_styler, proc_pivot = generate_processing_time_table(df_recent, 'Name of Agent (VA)', 'VA-SLA')
    return combined_pivot.astype(object).style.apply(highlight_cells,
                                                     axis=None), combined_pivot, df_valid, sla_styler, sla_pivot, proc_styler, proc_pivot, date_range_str, df_recent


def process_awaiting_publish(df):
    required_columns = ["Name of Agent (VA)", "Current State", "VA E-mail sent", "VA Request received", "JPR"]
    if [col for col in required_columns if col not in df.columns]: return None
    df_base = df[df['Current State'].isin(['Awaiting Approval', 'Pending RAS Validation'])].copy()
    df_base['VA Request received'] = pd.to_datetime(df_base['VA Request received'], errors='coerce')
    df_base = df_base.dropna(subset=['VA Request received'])
    if df_base.empty: return None

    df_base['Date_Only'] = df_base['VA Request received'].dt.date
    df_base['Formatted_Date'] = df_base['VA Request received'].dt.strftime('%Y-%m-%d')
    df_base['VA E-mail sent'] = pd.to_datetime(df_base['VA E-mail sent'], errors='coerce')

    def build_flat_jpr_table(df_sub, col_name):
        if df_sub.empty: return None, None
        agent_totals = df_sub.groupby('Name of Agent (VA)')['JPR'].count()
        detail_jprs = df_sub.groupby(['Name of Agent (VA)', 'Date_Only', 'Formatted_Date'])['JPR'].apply(
            lambda x: ', '.join(x.dropna().astype(str))).reset_index()
        display_rows = []
        for agent in agent_totals.index:
            display_rows.append({'Row Labels': agent, col_name: f"{agent_totals[agent]} (Total)"})
            agent_dates = detail_jprs[detail_jprs['Name of Agent (VA)'] == agent].sort_values('Date_Only')
            for _, row in agent_dates.iterrows():
                display_rows.append({'Row Labels': f"　　{row['Formatted_Date']}", col_name: row['JPR']})
        display_rows.append({'Row Labels': 'Grand Total', col_name: f"{df_sub['JPR'].count()} (Total)"})
        display_df = pd.DataFrame(display_rows)

        def style_flat(x):
            c = pd.DataFrame('', index=x.index, columns=x.columns)
            agent_mask = ~x['Row Labels'].str.startswith('　　')
            c.loc[agent_mask, 'Row Labels'] = 'font-weight: bold; background-color: #f8f9fa;'
            c.loc[agent_mask, col_name] = 'font-weight: bold; background-color: #f8f9fa;'
            gt_mask = x['Row Labels'] == 'Grand Total'
            c.loc[gt_mask, 'Row Labels'] = 'font-weight: bold; background-color: #EAEAEA;'
            c.loc[gt_mask, col_name] = 'font-weight: bold; background-color: #EAEAEA;'
            return c

        return display_df.style.apply(style_flat, axis=None), display_df

    return {
        'not_sent': build_flat_jpr_table(df_base[df_base['VA E-mail sent'].isna()].copy(), 'Pending JPRs'),
        'sent': build_flat_jpr_table(df_base[df_base['VA E-mail sent'].notna()].copy(), 'Pending JPRs'),
        'total': build_flat_jpr_table(df_base, 'Pending JPRs')
    }


def process_ll_email_sent(df, period_type='Weekly', small_period='Last 4 Weeks'):
    required_columns = ["Name of Agent (LL)", "LL E-mail sent", "JPR", "LL-SLA", "MET LL-SLA", "Post Title",
                        "Applications Reviewed"]
    if [col for col in required_columns if col not in df.columns]: return [None] * 9
    df_filtered = df[~df['Post Title'].astype(str).str.contains('Deputy', case=False, na=False)].copy()
    df_filtered['LL E-mail sent'] = pd.to_datetime(df_filtered['LL E-mail sent'], errors='coerce')
    df_valid = df_filtered.dropna(subset=['LL E-mail sent']).copy()
    df_valid['LL-SLA'] = pd.to_numeric(df_valid['LL-SLA'], errors='coerce')
    df_valid['Applications Reviewed'] = pd.to_numeric(df_valid['Applications Reviewed'], errors='coerce')
    if 'MET LL-SLA' in df_valid.columns: df_valid['MET LL-SLA'] = df_valid['MET LL-SLA'].astype(str).str.strip()
    if df_valid.empty: return [None] * 9

    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'LL E-mail sent', period_type)
    count_pivot = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (LL)', columns='Time Group',
                                 aggfunc='count', fill_value=0, margins=True, margins_name='Total')
    mean_pivot = pd.pivot_table(df_valid, values='LL-SLA', index='Name of Agent (LL)', columns='Time Group',
                                aggfunc='mean', margins=True, margins_name='Total')
    sum_pivot = pd.pivot_table(df_valid, values='Applications Reviewed', index='Name of Agent (LL)',
                               columns='Time Group', aggfunc='sum', fill_value=0, margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in count_pivot.columns]
    count_pivot, mean_pivot, sum_pivot = count_pivot[cols_order], mean_pivot.reindex(index=count_pivot.index,
                                                                                     columns=cols_order), sum_pivot.reindex(
        index=count_pivot.index, columns=cols_order)

    combined_pivot = pd.DataFrame(index=count_pivot.index, columns=cols_order)
    for col in cols_order:
        for row in count_pivot.index:
            c, m, s = count_pivot.at[row, col], mean_pivot.at[row, col], sum_pivot.at[row, col]
            if pd.isna(c) or c == 0:
                combined_pivot.at[row, col] = ""
            else:
                m_str = f"Sla: {m:.1f}" if pd.notna(m) else ""
                s_str = f"Apps: {int(s) if pd.notna(s) else 0}"
                combined_pivot.at[row, col] = f"{int(c)} ({m_str} | {s_str})" if m_str else f"{int(c)} ({s_str})"

    combined_pivot = combined_pivot.reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (LL)' in c.columns: c['Name of Agent (LL)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        if (x['Name of Agent (LL)'] == 'Total').any(): c.loc[x[x['Name of Agent (LL)'] == 'Total'].index[
            0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    df_recent, date_range_str = filter_current_period(df_valid, 'LL E-mail sent', small_period)
    sla_styler, sla_pivot = generate_sla_meeting_rate_table(df_recent, 'Name of Agent (LL)')
    proc_styler, proc_pivot = generate_processing_time_table(df_recent, 'Name of Agent (LL)', 'LL-SLA')
    return combined_pivot.astype(object).style.apply(highlight_cells,
                                                     axis=None), combined_pivot, df_valid, sla_styler, sla_pivot, proc_styler, proc_pivot, date_range_str, df_recent


def process_ll_released(df, period_type='Weekly'):
    required_columns = ["Name of Agent (LL)", "LL Triggered", "JPR", "Applications Reviewed"]
    if [col for col in required_columns if col not in df.columns]: return None, None, None
    df_valid = df.copy()
    df_valid['LL Triggered'] = pd.to_datetime(df_valid['LL Triggered'], errors='coerce')
    df_valid = df_valid.dropna(subset=['LL Triggered']).copy()
    df_valid['Applications Reviewed'] = pd.to_numeric(df_valid['Applications Reviewed'], errors='coerce')
    if df_valid.empty: return None, None, None

    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'LL Triggered', period_type)
    count_pivot = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (LL)', columns='Time Group',
                                 aggfunc='count', fill_value=0, margins=True, margins_name='Total')
    sum_pivot = pd.pivot_table(df_valid, values='Applications Reviewed', index='Name of Agent (LL)',
                               columns='Time Group', aggfunc='sum', fill_value=0, margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in count_pivot.columns]
    count_pivot, sum_pivot = count_pivot[cols_order], sum_pivot.reindex(index=count_pivot.index, columns=cols_order)

    combined_pivot = pd.DataFrame(index=count_pivot.index, columns=cols_order)
    for col in cols_order:
        for row in count_pivot.index:
            c, s = count_pivot.at[row, col], sum_pivot.at[row, col]
            combined_pivot.at[row, col] = "" if pd.isna(
                c) or c == 0 else f"{int(c)} (Apps: {int(s) if pd.notna(s) else 0})"

    combined_pivot = combined_pivot.reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (LL)' in c.columns: c['Name of Agent (LL)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        if (x['Name of Agent (LL)'] == 'Total').any(): c.loc[x[x['Name of Agent (LL)'] == 'Total'].index[
            0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    return combined_pivot.astype(object).style.apply(highlight_cells, axis=None), combined_pivot, df_valid


def process_awaiting_ll(df):
    required_columns = ["Name of Agent (LL)", "LL-Agent-Done?", "Post Title", "LL SLA Due Date", "JPR"]
    if [col for col in required_columns if col not in df.columns]: return None, None, None
    df_base = df[df['LL-Agent-Done?'].astype(str).str.strip().str.upper() == 'NO'].copy()
    df_base = df_base[~df_base['Post Title'].astype(str).str.contains('Deputy', case=False, na=False)]
    df_base['LL SLA Due Date'] = pd.to_datetime(df_base['LL SLA Due Date'], errors='coerce')
    df_valid = df_base.dropna(subset=['LL SLA Due Date']).copy()
    if df_valid.empty: return None, None, None

    today = pd.Timestamp.today().normalize()

    def format_due_date(d):
        d_norm = d.normalize()
        delta = (d_norm - today).days
        delta_str = f"in {delta} days" if delta > 1 else "in 1 day" if delta == 1 else "Today" if delta == 0 else "1 day ago" if delta == -1 else f"{-delta} days ago"
        return f"{d_norm.strftime('%Y-%m-%d')} ({d_norm.strftime('%A')}, {delta_str})"

    unique_dates = df_valid['LL SLA Due Date'].dt.normalize().drop_duplicates().sort_values()
    date_mapping = {d: format_due_date(d) for d in unique_dates}
    df_valid['Formatted Due Date'] = df_valid['LL SLA Due Date'].dt.normalize().map(date_mapping)
    pivot_df = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (LL)', columns='Formatted Due Date',
                              aggfunc='count', fill_value=0, margins=True, margins_name='Grand Total')
    cols_order = [c for c in [date_mapping[d] for d in unique_dates] + ['Grand Total'] if c in pivot_df.columns]
    pivot_df = pivot_df[cols_order].reset_index()

    def highlight_awaiting(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (LL)' in c.columns: c['Name of Agent (LL)'] = 'font-weight: bold;'
        for col in x.columns:
            if 'ago' in col:
                for row_idx in x.index:
                    if x.at[row_idx, col] != "": c.at[row_idx, col] = 'background-color: #F4CCCC; color: #000000;'
            elif 'Today' in col:
                for row_idx in x.index:
                    if x.at[row_idx, col] != "": c.at[row_idx, col] = 'background-color: #FFF2CC; color: #000000;'
            elif 'in ' in col:
                for row_idx in x.index:
                    if x.at[row_idx, col] != "": c.at[row_idx, col] = 'background-color: #CFE2F3; color: #000000;'
        if 'Grand Total' in c.columns: c[
            'Grand Total'] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        if (x['Name of Agent (LL)'] == 'Grand Total').any(): c.loc[x[x['Name of Agent (LL)'] == 'Grand Total'].index[
            0], :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    return pivot_df.astype(object).replace(0, "").style.apply(highlight_awaiting, axis=None), pivot_df, df_valid


def process_oc_creation(df, period_type='Weekly', small_period='Last 4 Weeks'):
    required_columns = ["Name of Agent (OC)", "Offer Creation Date", "JPR", "OC-SLA"]
    if [col for col in required_columns if col not in df.columns]: return [None] * 9
    df['Offer Creation Date'] = pd.to_datetime(df['Offer Creation Date'], errors='coerce')
    df_valid = df.dropna(subset=['Offer Creation Date']).copy()
    df_valid['OC-SLA'] = pd.to_numeric(df_valid['OC-SLA'], errors='coerce')
    if 'MET OC-SLA' in df_valid.columns: df_valid['MET OC-SLA'] = df_valid['MET OC-SLA'].astype(str).str.strip()
    if df_valid.empty: return [None] * 9

    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'Offer Creation Date', period_type)
    count_pivot = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (OC)', columns='Time Group',
                                 aggfunc='count', fill_value=0, margins=True, margins_name='Total')
    mean_pivot = pd.pivot_table(df_valid, values='OC-SLA', index='Name of Agent (OC)', columns='Time Group',
                                aggfunc='mean', margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in count_pivot.columns]
    count_pivot, mean_pivot = count_pivot[cols_order], mean_pivot.reindex(index=count_pivot.index, columns=cols_order)

    combined_pivot = pd.DataFrame(index=count_pivot.index, columns=cols_order)
    for col in cols_order:
        for row in count_pivot.index:
            c, m = count_pivot.at[row, col], mean_pivot.at[row, col]
            combined_pivot.at[row, col] = "" if pd.isna(c) or c == 0 else (
                f"{int(c)}" if pd.isna(m) else f"{int(c)} (Sla: {m:.1f})")

    combined_pivot = combined_pivot.reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (OC)' in c.columns: c['Name of Agent (OC)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #F2F2F2; color: #000000; font-weight: bold;'
        if (x['Name of Agent (OC)'] == 'Total').any(): c.loc[x[x['Name of Agent (OC)'] == 'Total'].index[
            0], :] = 'background-color: #F2F2F2; color: #000000; font-weight: bold;'
        return c

    df_recent, date_range_str = filter_current_period(df_valid, 'Offer Creation Date', small_period)
    sla_styler, sla_pivot = generate_sla_meeting_rate_table(df_recent, 'Name of Agent (OC)')
    proc_styler, proc_pivot = generate_processing_time_table(df_recent, 'Name of Agent (OC)', 'OC-SLA')
    return combined_pivot.astype(object).style.apply(highlight_cells,
                                                     axis=None), combined_pivot, df_valid, sla_styler, sla_pivot, proc_styler, proc_pivot, date_range_str, df_recent


def process_fc_request(df, period_type='Weekly', small_period='Last 4 Weeks'):
    required_columns = ["Name of Agent (OC)", "Request Funding Check", "JPR", "FC-SLA"]
    if [col for col in required_columns if col not in df.columns]: return [None] * 9
    df['Request Funding Check'] = pd.to_datetime(df['Request Funding Check'], errors='coerce')
    df_valid = df.dropna(subset=['Request Funding Check']).copy()
    df_valid['FC-SLA'] = pd.to_numeric(df_valid['FC-SLA'], errors='coerce')
    if 'MET FC-SLA' in df_valid.columns: df_valid['MET FC-SLA'] = df_valid['MET FC-SLA'].astype(str).str.strip()
    if df_valid.empty: return [None] * 9

    df_valid['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_valid, 'Request Funding Check',
                                                                          period_type)
    count_pivot = pd.pivot_table(df_valid, values='JPR', index='Name of Agent (OC)', columns='Time Group',
                                 aggfunc='count', fill_value=0, margins=True, margins_name='Total')
    mean_pivot = pd.pivot_table(df_valid, values='FC-SLA', index='Name of Agent (OC)', columns='Time Group',
                                aggfunc='mean', margins=True, margins_name='Total')
    cols_order = [col for col in cols_order if col in count_pivot.columns]
    count_pivot, mean_pivot = count_pivot[cols_order], mean_pivot.reindex(index=count_pivot.index, columns=cols_order)

    combined_pivot = pd.DataFrame(index=count_pivot.index, columns=cols_order)
    for col in cols_order:
        for row in count_pivot.index:
            c, m = count_pivot.at[row, col], mean_pivot.at[row, col]
            combined_pivot.at[row, col] = "" if pd.isna(c) or c == 0 else (
                f"{int(c)}" if pd.isna(m) else f"{int(c)} (Avg: {m:.1f})")

    combined_pivot = combined_pivot.reset_index()

    def highlight_cells(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        if 'Name of Agent (OC)' in c.columns: c['Name of Agent (OC)'] = 'font-weight: bold;'
        if p1_range in c.columns: c[p1_range] = 'background-color: #FFF2CC; color: #000000;'
        if 'Total' in c.columns: c['Total'] = 'background-color: #F2F2F2; color: #000000; font-weight: bold;'
        if (x['Name of Agent (OC)'] == 'Total').any(): c.loc[x[x['Name of Agent (OC)'] == 'Total'].index[
            0], :] = 'background-color: #F2F2F2; color: #000000; font-weight: bold;'
        return c

    df_recent, date_range_str = filter_current_period(df_valid, 'Request Funding Check', small_period)
    sla_styler, sla_pivot = generate_sla_meeting_rate_table(df_recent, 'Name of Agent (OC)')
    proc_styler, proc_pivot = generate_processing_time_table(df_recent, 'Name of Agent (OC)', 'FC-SLA')
    return combined_pivot.astype(object).style.apply(highlight_cells,
                                                     axis=None), combined_pivot, df_valid, sla_styler, sla_pivot, proc_styler, proc_pivot, date_range_str, df_recent


def process_ras_kpi(df, period_type='Weekly'):
    required_columns = ["Request Funding Check", "NWD of VA expiry to FC submitted", "LL-SLA", "OC-SLA", "FC-SLA",
                        "JPR"]
    if [col for col in required_columns if col not in df.columns]: return None, None, None
    df_kpi = df.dropna(subset=['NWD of VA expiry to FC submitted']).copy()
    df_kpi['Request Funding Check'] = pd.to_datetime(df_kpi['Request Funding Check'], errors='coerce')
    df_kpi = df_kpi.dropna(subset=['Request Funding Check'])
    if df_kpi.empty: return None, None, None

    num_cols = ["LL-SLA", "OC-SLA", "FC-SLA", "NWD of VA expiry to FC submitted"]
    for col in num_cols: df_kpi[col] = pd.to_numeric(df_kpi[col], errors='coerce')

    df_kpi['Time Group'], cols_order, p1_range = apply_dynamic_grouping(df_kpi, 'Request Funding Check', period_type,
                                                                        num_periods=12)
    grouped = df_kpi.groupby('Time Group')[num_cols].mean()
    grouped['Total JPR Completed'] = df_kpi.groupby('Time Group')['JPR'].count()

    total_mean = df_kpi[num_cols].mean()
    total_mean['Total JPR Completed'] = df_kpi['JPR'].count()
    total_mean.name = 'Total'
    grouped = pd.concat([grouped, pd.DataFrame(total_mean).T])

    valid_rows = [r for r in cols_order if r in grouped.index]
    grouped = grouped.loc[valid_rows]

    grouped['RAS KPI'] = grouped['LL-SLA'].fillna(0) + grouped['OC-SLA'].fillna(0) + grouped['FC-SLA'].fillna(0)
    grouped['Outside RAS'] = grouped['NWD of VA expiry to FC submitted'].fillna(0) - grouped['RAS KPI']

    grouped = grouped.reset_index().rename(columns={
        'index': 'Time Group', 'LL-SLA': 'Avg LL-SLA', 'OC-SLA': 'Avg OC-SLA', 'FC-SLA': 'Avg FC-SLA',
        'NWD of VA expiry to FC submitted': 'Avg NWD (VA to FC)'
    })

    cols_order_display = ['Time Group', 'Total JPR Completed', 'Avg LL-SLA', 'Avg OC-SLA', 'Avg FC-SLA',
                          'Avg NWD (VA to FC)', 'RAS KPI', 'Outside RAS']
    grouped = grouped[cols_order_display]

    def highlight_kpi(x):
        c = pd.DataFrame('', index=x.index, columns=x.columns)
        for idx in x.index:
            if x.at[idx, 'Time Group'] == p1_range:
                c.loc[idx, :] = 'background-color: #FFF2CC; color: #000000;'
                c.at[idx, 'Time Group'] = 'background-color: #FFF2CC; color: #000000; font-weight: bold;'
            elif x.at[idx, 'Time Group'] == 'Total':
                c.loc[idx, :] = 'background-color: #EAEAEA; color: #000000; font-weight: bold;'
        return c

    format_dict = {'Total JPR Completed': "{:.0f}", 'Avg LL-SLA': "{:.1f}", 'Avg OC-SLA': "{:.1f}",
                   'Avg FC-SLA': "{:.1f}", 'Avg NWD (VA to FC)': "{:.1f}", 'RAS KPI': "{:.1f}", 'Outside RAS': "{:.1f}"}
    return grouped.style.apply(highlight_kpi, axis=None).format(format_dict, na_rep="-"), grouped, df_kpi


# ==========================================
# 🌟 新增：生成带格式的多 Sheet Excel 报告核心引擎
# ==========================================
@st.cache_data(show_spinner=False)
def create_excel_report(df, period_options, small_table_options):
    output = io.BytesIO()
    # 强制使用 xlsxwriter 引擎以支持 DataFrame Styler 的 CSS 解析
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        title_format = workbook.add_format({'bold': True, 'font_size': 13, 'bg_color': '#D9E1F2', 'border': 1})

        def write_section(sheet_name, sections):
            if sheet_name not in writer.sheets:
                worksheet = workbook.add_worksheet(sheet_name)
                writer.sheets[sheet_name] = worksheet
            worksheet = writer.sheets[sheet_name]

            startrow = 0
            col_widths = {}

            for title, item in sections:
                if item is None: continue

                if isinstance(item, tuple):
                    styler, data_df = item[0], item[1]
                elif hasattr(item, 'data'):  # Is Styler
                    styler, data_df = item, item.data
                else:  # Is DataFrame
                    styler, data_df = item.style, item

                if styler is None or data_df is None or data_df.empty:
                    continue

                # 写标题
                worksheet.write_string(startrow, 0, title, title_format)
                startrow += 2

                # 写入带有 CSS 格式的 Styler (Pandas 会自动将其转为 Excel 颜色)
                styler.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)

                # 计算并记录此 Sheet 每一列的最大字符宽度
                for i, col in enumerate(data_df.columns):
                    # 取列头和该列内容的最大长度 (防止遇到空数据报错)
                    max_len = max(data_df[col].astype(str).map(len).max() if not data_df.empty else 0, len(str(col)))
                    col_widths[i] = max(col_widths.get(i, 0), max_len)

                startrow += len(data_df.index) + 3  # 加 3 行空隙给下一个表格

            # 统一应用列宽调节 (加 3 个字符作为留白)
            for i, width in col_widths.items():
                worksheet.set_column(i, i, width + 3)

        # ====== 生成 VA 标签页 ======
        styled_va_email, _, _ = process_va_email_sent(df, period_options)
        result_pub = process_va_published(df, period_options, small_table_options)
        awaiting_va = process_awaiting_publish(df)
        va_sections = [
            ("1. VA E-mail Sent Volume", styled_va_email),
            ("2. VA SLA Meeting Rate", result_pub[3] if result_pub else None),
            ("3. VA Published Processing Time", result_pub[5] if result_pub else None),
            ("4. VA Published Volume & Avg Processing Time", result_pub[0] if result_pub else None),
        ]
        if awaiting_va:
            va_sections.extend([
                ("5. Awaiting VA Publish - Not Sent", awaiting_va['not_sent'][0] if awaiting_va['not_sent'] else None),
                ("6. Awaiting VA Publish - Sent & Waiting", awaiting_va['sent'][0] if awaiting_va['sent'] else None),
                ("7. Awaiting VA Publish - Total", awaiting_va['total'][0] if awaiting_va['total'] else None),
            ])
        write_section("VA", va_sections)

        # ====== 生成 LL 标签页 ======
        result_ll_email = process_ll_email_sent(df, period_options, small_table_options)
        styled_ll_rel, _, _ = process_ll_released(df, period_options)
        result_awaiting_ll = process_awaiting_ll(df)
        ll_sections = [
            ("1. LL SLA Meeting Rate", result_ll_email[3] if result_ll_email else None),
            ("2. LL Processing Time", result_ll_email[5] if result_ll_email else None),
            ("3. LL E-mail Sent Volume", result_ll_email[0] if result_ll_email else None),
            ("4. LL Released Volume", styled_ll_rel),
            ("5. Awaiting LL", result_awaiting_ll[0] if result_awaiting_ll else None)
        ]
        write_section("LL", ll_sections)

        # ====== 生成 OC 标签页 ======
        result_oc = process_oc_creation(df, period_options, small_table_options)
        oc_sections = [
            ("1. OC SLA Meeting Rate", result_oc[3] if result_oc else None),
            ("2. OC Processing Time", result_oc[5] if result_oc else None),
            ("3. Offer Creation Volume & Avg Processing Time", result_oc[0] if result_oc else None)
        ]
        write_section("OC", oc_sections)

        # ====== 生成 FC 标签页 ======
        result_fc = process_fc_request(df, period_options, small_table_options)
        fc_sections = [
            ("1. FC SLA Meeting Rate", result_fc[3] if result_fc else None),
            ("2. FC Processing Time", result_fc[5] if result_fc else None),
            ("3. Funding Check Volume & Avg Processing Time", result_fc[0] if result_fc else None)
        ]
        write_section("FC", fc_sections)

        # ====== 生成 KPI 标签页 ======
        result_kpi = process_ras_kpi(df, period_options)
        kpi_sections = [
            ("1. End-to-End Processing Time Analysis", result_kpi[0] if result_kpi else None)
        ]
        write_section("RAS-KPI", kpi_sections)

    return output.getvalue()


# ==========================================
# 主应用入口
# ==========================================
def main():
    st.markdown(
        """
        <style>
            [data-testid="stSidebar"] { min-width: 200px !important; max-width: 200px !important; }
            [data-testid="stSidebarContent"] { display: flex !important; flex-direction: column !important; justify-content: center !important; }
            [data-testid="stSidebarHeader"] { display: none !important; }
            [data-testid="stSidebarContent"] div[role="radiogroup"] { gap: 1.5rem !important; padding-left: 20px !important; }
            [data-testid="stSidebarContent"] div[role="radiogroup"] label { font-size: 1.3rem !important; font-weight: bold !important; }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("📊 RAS Monitoring Report Analysis")
    st.markdown("Upload your data source (Excel file) to generate reports for VA, LL, OC, FC, and RAS-KPI.")

    with st.sidebar:
        selected_nav = st.radio("Navigation", ["VA", "LL", "OC", "FC", "RAS-KPI"], label_visibility="collapsed")

    uploaded_file = st.file_uploader("Upload Data Source (.xlsx)", type=["xlsx"])
    period_options = ["Weekly", "Monthly", "Quarterly", "Yearly"]
    small_table_options = ["Last 4 Weeks", "Current Quarter", "Current Year"]

    if uploaded_file is not None:
        try:
            with st.spinner("Reading data..."):
                df = load_data(uploaded_file)

            # 🌟 新增：Excel 导出控制区
            st.sidebar.markdown("---")
            st.sidebar.markdown("### 📥 Export Report")
            export_period = st.sidebar.selectbox("Main Table Time:", period_options, index=0)
            export_small = st.sidebar.selectbox("SLA Table Time:", small_table_options, index=0)

            # 静默生成带有样式的 Excel 二进制数据
            excel_data = create_excel_report(df, export_period, export_small)

            st.sidebar.download_button(
                label="⬇️ Download Excel",
                data=excel_data,
                file_name=f"RAS_Report_{export_period}_{export_small.replace(' ', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

            # --------------------------
            # 模块 1: VA
            # --------------------------
            if selected_nav == "VA":
                st.header("Vacancy Announcement (VA) Reports")
                st.subheader("1. VA E-mail Sent")
                period_va_email = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                           key="va_email_period")
                styled_va_email, raw_pivot_email, df_valid_email = process_va_email_sent(df, period_va_email)

                if styled_va_email is not None:
                    col_table1, _ = st.columns([8, 2])
                    with col_table1:
                        event_email = st.dataframe(styled_va_email, use_container_width=True, hide_index=True,
                                                   height=(len(raw_pivot_email) * 35) + 38, on_select="rerun",
                                                   selection_mode="single-cell", key="table_email")

                    if 'va_email_state' not in st.session_state: st.session_state.va_email_state = []
                    cells_email = getattr(event_email.selection, 'cells', [])

                    if cells_email != st.session_state.va_email_state:
                        st.session_state.va_email_state = cells_email
                        if cells_email:
                            cell = cells_email[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_email, raw_col_name)
                            selected_agent = raw_pivot_email['Name of Agent (VA)'].iloc[row_idx]

                            drill_df = df_valid_email.copy().dropna(subset=['JPR'])
                            if selected_agent != 'Total': drill_df = drill_df[
                                drill_df['Name of Agent (VA)'] == selected_agent]
                            if col_name not in ['Total', 'Name of Agent (VA)']: drill_df = drill_df[
                                drill_df['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing raw data for Agent: **{selected_agent}** | Time Group: **{col_name}**",
                                drill_df)

                st.markdown("---")
                st.subheader("2. VA Published")

                col_ctrl_va_left, col_ctrl_va_right = st.columns([3.5, 6.5])
                with col_ctrl_va_left:
                    small_va_pub = st.radio("SLA & Proc. Time Grouping:", small_table_options, horizontal=True,
                                            key="va_pub_small")
                with col_ctrl_va_right:
                    period_va_pub = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                             key="va_pub_period")

                result_pub = process_va_published(df, period_va_pub, small_va_pub)
                if result_pub[0] is not None:
                    styled_va_pub, raw_pivot_pub, df_valid_pub, sla_styler, sla_pivot, proc_styler, proc_pivot, date_range_str, df_recent_4w = result_pub

                    col_left, col_right = st.columns([3.5, 6.5])
                    with col_left:
                        st.markdown(f"**VA SLA Meeting Rate ({date_range_str})**")
                        event_sla = st.dataframe(sla_styler, use_container_width=False, hide_index=True,
                                                 height=(len(sla_pivot) * 35) + 38, on_select="rerun",
                                                 selection_mode="single-cell",
                                                 key="table_sla") if sla_styler is not None else None
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(f"**VA Published Processing Time ({date_range_str})**")
                        event_proc = st.dataframe(proc_styler, use_container_width=False, hide_index=True,
                                                  height=(len(proc_pivot) * 35) + 38, on_select="rerun",
                                                  selection_mode="single-cell",
                                                  key="table_proc") if proc_styler is not None else None
                    with col_right:
                        st.markdown("**VA Published Volume & Avg Processing Time**")
                        event_pub = st.dataframe(styled_va_pub, use_container_width=False, hide_index=True,
                                                 height=(len(raw_pivot_pub) * 35) + 38, on_select="rerun",
                                                 selection_mode="single-cell", key="table_pub_main")

                    if 'va_pub_state' not in st.session_state: st.session_state.va_pub_state = {'sla': [], 'proc': [],
                                                                                                'pub': []}
                    cells_sla = getattr(event_sla.selection, 'cells', []) if event_sla else []
                    cells_proc = getattr(event_proc.selection, 'cells', []) if event_proc else []
                    cells_pub = getattr(event_pub.selection, 'cells', []) if event_pub else []

                    if cells_sla != st.session_state.va_pub_state['sla']:
                        st.session_state.va_pub_state['sla'] = cells_sla
                        if cells_sla:
                            cell = cells_sla[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(sla_pivot, raw_col_name)
                            agent = sla_pivot['Name of Agent (VA)'].iloc[row_idx]

                            drill_df = df_recent_4w.copy().dropna(subset=['JPR'])
                            drill_df = drill_df[drill_df['MET VA-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df = drill_df[drill_df['Name of Agent (VA)'] == agent]
                            if col_name in ['No', 'Yes']: drill_df = drill_df[drill_df['MET VA-SLA'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **SLA Meeting Rate** data for: **{agent}** | Column: **{col_name}**",
                                drill_df)

                    elif cells_proc != st.session_state.va_pub_state['proc']:
                        st.session_state.va_pub_state['proc'] = cells_proc
                        if cells_proc:
                            cell = cells_proc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(proc_pivot, raw_col_name)
                            agent = proc_pivot['Name of Agent (VA)'].iloc[row_idx]

                            drill_df = df_recent_4w.copy().dropna(subset=['JPR', 'VA-SLA'])
                            drill_df = drill_df[drill_df['MET VA-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df = drill_df[drill_df['Name of Agent (VA)'] == agent]
                            if col_name not in ['Grand Total', 'Name of Agent (VA)']:
                                if col_name.startswith('No'):
                                    drill_df = drill_df[drill_df['MET VA-SLA'] == 'No']
                                elif col_name.startswith('Yes'):
                                    drill_df = drill_df[drill_df['MET VA-SLA'] == 'Yes']
                                nums = re.findall(r'\((\d+\.?\d*)\)', col_name)
                                if nums: drill_df = drill_df[drill_df['VA-SLA'].astype(float) == float(nums[0])]
                            show_drilldown_modal(
                                f"💡 Showing **Processing Time** data for: **{agent}** | Column: **{col_name}**",
                                drill_df)

                    elif cells_pub != st.session_state.va_pub_state['pub']:
                        st.session_state.va_pub_state['pub'] = cells_pub
                        if cells_pub:
                            cell = cells_pub[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_pub, raw_col_name)
                            agent = raw_pivot_pub['Name of Agent (VA)'].iloc[row_idx]

                            drill_df = df_valid_pub.copy().dropna(subset=['JPR'])
                            if agent != 'Total': drill_df = drill_df[drill_df['Name of Agent (VA)'] == agent]
                            if col_name not in ['Total', 'Name of Agent (VA)']: drill_df = drill_df[
                                drill_df['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **VA Published** data for: **{agent}** | Time Group: **{col_name}**",
                                drill_df)

                st.markdown("---")
                st.subheader("3. Awaiting VA Publish")
                awaiting_results = process_awaiting_publish(df)
                if awaiting_results:
                    col_aw1, col_aw2, col_aw3 = st.columns(3)
                    with col_aw1:
                        st.markdown("**1. VA E-mail not sent to HRBP by Request Received Date**")
                        if awaiting_results['not_sent'][0] is not None: st.dataframe(awaiting_results['not_sent'][0],
                                                                                     use_container_width=True,
                                                                                     hide_index=True, height=(
                                                                                                                         len(
                                                                                                                             awaiting_results[
                                                                                                                                 'not_sent'][
                                                                                                                                 1]) * 35) + 40)
                    with col_aw2:
                        st.markdown("**2. VA E-mail sent & Pending HRBP's Validation by Request Received Date**")
                        if awaiting_results['sent'][0] is not None: st.dataframe(awaiting_results['sent'][0],
                                                                                 use_container_width=True,
                                                                                 hide_index=True, height=(
                                                                                                                     len(
                                                                                                                         awaiting_results[
                                                                                                                             'sent'][
                                                                                                                             1]) * 35) + 40)
                    with col_aw3:
                        st.markdown("**3. Total Pending Advertisement by Request Received Date**")
                        if awaiting_results['total'][0] is not None: st.dataframe(awaiting_results['total'][0],
                                                                                  use_container_width=True,
                                                                                  hide_index=True, height=(
                                                                                                                      len(
                                                                                                                          awaiting_results[
                                                                                                                              'total'][
                                                                                                                              1]) * 35) + 40)

            # --------------------------
            # 模块 2: LL
            # --------------------------
            elif selected_nav == "LL":
                st.header("Longlist QA Mannual Check (LL) Reports")
                st.subheader("1. LL E-mail Sent (Excluding 'Deputy')")

                col_ctrl_ll_left, col_ctrl_ll_right = st.columns([3, 7])
                with col_ctrl_ll_left:
                    small_ll_email = st.radio("SLA & Proc. Time Grouping:", small_table_options, horizontal=True,
                                              key="ll_email_small")
                with col_ctrl_ll_right:
                    period_ll_email = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                               key="ll_email_period")

                result_ll_email = process_ll_email_sent(df, period_ll_email, small_ll_email)
                if result_ll_email[0] is not None:
                    styled_ll_email, raw_pivot_ll_email, df_valid_ll_email, sla_styler_ll, sla_pivot_ll, proc_styler_ll, proc_pivot_ll, date_range_str_ll, df_recent_4w_ll = result_ll_email

                    col_left_ll, col_right_ll = st.columns([3, 7])
                    with col_left_ll:
                        st.markdown(f"**LL SLA Meeting Rate ({date_range_str_ll})**")
                        event_sla_ll = st.dataframe(sla_styler_ll, use_container_width=True, hide_index=True,
                                                    height=(len(sla_pivot_ll) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell",
                                                    key="table_ll_email_sla") if sla_styler_ll is not None else None
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(f"**LL Processing Time ({date_range_str_ll})**")
                        event_proc_ll = st.dataframe(proc_styler_ll, use_container_width=False, hide_index=True,
                                                     height=(len(proc_pivot_ll) * 35) + 38, on_select="rerun",
                                                     selection_mode="single-cell",
                                                     key="table_ll_email_proc") if proc_styler_ll is not None else None
                    with col_right_ll:
                        st.markdown("**LL E-mail Sent Volume, Avg Processing Time & Apps Reviewed**")
                        event_ll_email = st.dataframe(styled_ll_email, use_container_width=False, hide_index=True,
                                                      height=(len(raw_pivot_ll_email) * 35) + 38, on_select="rerun",
                                                      selection_mode="single-cell", key="table_ll_email_main")

                    if 'll_email_state' not in st.session_state: st.session_state.ll_email_state = {'sla': [],
                                                                                                    'proc': [],
                                                                                                    'email': []}
                    cells_sla_ll = getattr(event_sla_ll.selection, 'cells', []) if event_sla_ll else []
                    cells_proc_ll = getattr(event_proc_ll.selection, 'cells', []) if event_proc_ll else []
                    cells_email_ll = getattr(event_ll_email.selection, 'cells', []) if event_ll_email else []

                    if cells_sla_ll != st.session_state.ll_email_state['sla']:
                        st.session_state.ll_email_state['sla'] = cells_sla_ll
                        if cells_sla_ll:
                            cell = cells_sla_ll[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(sla_pivot_ll, raw_col_name)
                            agent = sla_pivot_ll['Name of Agent (LL)'].iloc[row_idx]

                            drill_df_ll = df_recent_4w_ll.copy().dropna(subset=['JPR'])
                            drill_df_ll = drill_df_ll[drill_df_ll['MET LL-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_ll = drill_df_ll[
                                drill_df_ll['Name of Agent (LL)'] == agent]
                            if col_name in ['No', 'Yes']: drill_df_ll = drill_df_ll[
                                drill_df_ll['MET LL-SLA'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **SLA Meeting Rate** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_ll)

                    elif cells_proc_ll != st.session_state.ll_email_state['proc']:
                        st.session_state.ll_email_state['proc'] = cells_proc_ll
                        if cells_proc_ll:
                            cell = cells_proc_ll[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(proc_pivot_ll, raw_col_name)
                            agent = proc_pivot_ll['Name of Agent (LL)'].iloc[row_idx]

                            drill_df_ll = df_recent_4w_ll.copy().dropna(subset=['JPR', 'LL-SLA'])
                            drill_df_ll = drill_df_ll[drill_df_ll['MET LL-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_ll = drill_df_ll[
                                drill_df_ll['Name of Agent (LL)'] == agent]
                            if col_name not in ['Grand Total', 'Name of Agent (LL)']:
                                if col_name.startswith('No'):
                                    drill_df_ll = drill_df_ll[drill_df_ll['MET LL-SLA'] == 'No']
                                elif col_name.startswith('Yes'):
                                    drill_df_ll = drill_df_ll[drill_df_ll['MET LL-SLA'] == 'Yes']
                                nums = re.findall(r'\((\d+\.?\d*)\)', col_name)
                                if nums: drill_df_ll = drill_df_ll[
                                    drill_df_ll['LL-SLA'].astype(float) == float(nums[0])]
                            show_drilldown_modal(
                                f"💡 Showing **Processing Time** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_ll)

                    elif cells_email_ll != st.session_state.ll_email_state['email']:
                        st.session_state.ll_email_state['email'] = cells_email_ll
                        if cells_email_ll:
                            cell = cells_email_ll[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_ll_email, raw_col_name)
                            agent = raw_pivot_ll_email['Name of Agent (LL)'].iloc[row_idx]

                            drill_df_ll = df_valid_ll_email.copy().dropna(subset=['JPR'])
                            if agent != 'Total': drill_df_ll = drill_df_ll[drill_df_ll['Name of Agent (LL)'] == agent]
                            if col_name not in ['Total', 'Name of Agent (LL)']: drill_df_ll = drill_df_ll[
                                drill_df_ll['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **LL E-mail Sent** data for: **{agent}** | Time Group: **{col_name}**",
                                drill_df_ll)

                st.markdown("---")
                st.subheader("2. LL Released")
                period_ll_rel = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                         key="ll_rel_period")
                styled_ll_rel, raw_pivot_ll_rel, df_valid_ll_rel = process_ll_released(df, period_ll_rel)

                if styled_ll_rel is not None:
                    col_table_ll2, _ = st.columns([8, 2])
                    with col_table_ll2:
                        event_ll_rel = st.dataframe(styled_ll_rel, use_container_width=True, hide_index=True,
                                                    height=(len(raw_pivot_ll_rel) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell", key="table_ll_rel_simple")

                    if 'll_rel_state' not in st.session_state: st.session_state.ll_rel_state = []
                    cells_ll_rel = getattr(event_ll_rel.selection, 'cells', [])
                    if cells_ll_rel != st.session_state.ll_rel_state:
                        st.session_state.ll_rel_state = cells_ll_rel
                        if cells_ll_rel:
                            cell = cells_ll_rel[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_ll_rel, raw_col_name)
                            selected_agent = raw_pivot_ll_rel['Name of Agent (LL)'].iloc[row_idx]

                            drill_df_ll_rel = df_valid_ll_rel.copy().dropna(subset=['JPR'])
                            if selected_agent != 'Total': drill_df_ll_rel = drill_df_ll_rel[
                                drill_df_ll_rel['Name of Agent (LL)'] == selected_agent]
                            if col_name not in ['Total', 'Name of Agent (LL)']: drill_df_ll_rel = drill_df_ll_rel[
                                drill_df_ll_rel['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing raw data for Agent: **{selected_agent}** | Time Group: **{col_name}**",
                                drill_df_ll_rel)

                st.markdown("---")
                st.subheader("3. Awaiting LL")
                st.markdown(
                    """
                    <div style="display: flex; gap: 20px; margin-bottom: 10px; font-size: 14px;">
                        <div style="display: flex; align-items: center; gap: 6px;"><div style="width: 16px; height: 16px; background-color: #F4CCCC; border: 1px solid #d9d9d9; border-radius: 3px;"></div><span><b>Overdue</b> (ago)</span></div>
                        <div style="display: flex; align-items: center; gap: 6px;"><div style="width: 16px; height: 16px; background-color: #FFF2CC; border: 1px solid #d9d9d9; border-radius: 3px;"></div><span><b>Due Today</b> (Today)</span></div>
                        <div style="display: flex; align-items: center; gap: 6px;"><div style="width: 16px; height: 16px; background-color: #CFE2F3; border: 1px solid #d9d9d9; border-radius: 3px;"></div><span><b>Future Due</b> (in days)</span></div>
                    </div>
                    """, unsafe_allow_html=True
                )
                result_awaiting_ll = process_awaiting_ll(df)
                if result_awaiting_ll and result_awaiting_ll[0] is not None:
                    styled_awaiting_ll, raw_pivot_awaiting_ll, df_valid_awaiting_ll = result_awaiting_ll
                    event_awaiting_ll = st.dataframe(styled_awaiting_ll, use_container_width=True, hide_index=True,
                                                     height=(len(raw_pivot_awaiting_ll) * 35) + 38, on_select="rerun",
                                                     selection_mode="single-cell", key="table_awaiting_ll")

                    if 'awaiting_ll_state' not in st.session_state: st.session_state.awaiting_ll_state = []
                    cells_awaiting_ll = getattr(event_awaiting_ll.selection, 'cells', [])
                    if cells_awaiting_ll != st.session_state.awaiting_ll_state:
                        st.session_state.awaiting_ll_state = cells_awaiting_ll
                        if cells_awaiting_ll:
                            cell = cells_awaiting_ll[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_awaiting_ll, raw_col_name)
                            selected_agent = raw_pivot_awaiting_ll['Name of Agent (LL)'].iloc[row_idx]

                            drill_df_awaiting = df_valid_awaiting_ll.copy().dropna(subset=['JPR'])
                            if selected_agent != 'Grand Total': drill_df_awaiting = drill_df_awaiting[
                                drill_df_awaiting['Name of Agent (LL)'] == selected_agent]
                            if col_name not in ['Grand Total', 'Name of Agent (LL)']: drill_df_awaiting = \
                            drill_df_awaiting[drill_df_awaiting['Formatted Due Date'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing raw data for Agent: **{selected_agent}** | Due Date: **{col_name}**",
                                drill_df_awaiting)

            # --------------------------
            # 模块 3: OC
            # --------------------------
            elif selected_nav == "OC":
                st.header("Offer Creation (OC) Reports")
                st.subheader("1. Offer Creation Volume & Processing Time")

                col_ctrl_oc_left, col_ctrl_oc_right = st.columns([4, 6])
                with col_ctrl_oc_left:
                    small_oc_pub = st.radio("SLA & Proc. Time Grouping:", small_table_options, horizontal=True,
                                            key="oc_pub_small")
                with col_ctrl_oc_right:
                    period_oc_pub = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                             key="oc_pub_period")

                result_oc_pub = process_oc_creation(df, period_oc_pub, small_oc_pub)
                if result_oc_pub[0] is not None:
                    styled_oc_pub, raw_pivot_oc_pub, df_valid_oc_pub, sla_styler_oc, sla_pivot_oc, proc_styler_oc, proc_pivot_oc, date_range_str_oc, df_recent_4w_oc = result_oc_pub

                    col_left_oc, col_right_oc = st.columns([4, 6])
                    with col_left_oc:
                        st.markdown(f"**OC SLA Meeting Rate ({date_range_str_oc})**")
                        event_sla_oc = st.dataframe(sla_styler_oc, use_container_width=True, hide_index=True,
                                                    height=(len(sla_pivot_oc) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell",
                                                    key="table_oc_sla") if sla_styler_oc is not None else None
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(f"**OC Processing Time ({date_range_str_oc})**")
                        event_proc_oc = st.dataframe(proc_styler_oc, use_container_width=True, hide_index=True,
                                                     height=(len(proc_pivot_oc) * 35) + 38, on_select="rerun",
                                                     selection_mode="single-cell",
                                                     key="table_oc_proc") if proc_styler_oc is not None else None
                    with col_right_oc:
                        st.markdown("**Offer Creation Volume & Avg Processing Time**")
                        event_oc_pub = st.dataframe(styled_oc_pub, use_container_width=True, hide_index=True,
                                                    height=(len(raw_pivot_oc_pub) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell", key="table_oc_pub_main")

                    if 'oc_pub_state' not in st.session_state: st.session_state.oc_pub_state = {'sla': [], 'proc': [],
                                                                                                'pub': []}
                    cells_sla_oc = getattr(event_sla_oc.selection, 'cells', []) if event_sla_oc else []
                    cells_proc_oc = getattr(event_proc_oc.selection, 'cells', []) if event_proc_oc else []
                    cells_pub_oc = getattr(event_oc_pub.selection, 'cells', []) if event_oc_pub else []

                    if cells_sla_oc != st.session_state.oc_pub_state['sla']:
                        st.session_state.oc_pub_state['sla'] = cells_sla_oc
                        if cells_sla_oc:
                            cell = cells_sla_oc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(sla_pivot_oc, raw_col_name)
                            agent = sla_pivot_oc['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_oc = df_recent_4w_oc.copy().dropna(subset=['JPR'])
                            drill_df_oc = drill_df_oc[drill_df_oc['MET OC-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_oc = drill_df_oc[
                                drill_df_oc['Name of Agent (OC)'] == agent]
                            if col_name in ['No', 'Yes']: drill_df_oc = drill_df_oc[
                                drill_df_oc['MET OC-SLA'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **SLA Meeting Rate** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_oc)

                    elif cells_proc_oc != st.session_state.oc_pub_state['proc']:
                        st.session_state.oc_pub_state['proc'] = cells_proc_oc
                        if cells_proc_oc:
                            cell = cells_proc_oc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(proc_pivot_oc, raw_col_name)
                            agent = proc_pivot_oc['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_oc = df_recent_4w_oc.copy().dropna(subset=['JPR', 'OC-SLA'])
                            drill_df_oc = drill_df_oc[drill_df_oc['MET OC-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_oc = drill_df_oc[
                                drill_df_oc['Name of Agent (OC)'] == agent]
                            if col_name not in ['Grand Total', 'Name of Agent (OC)']:
                                if col_name.startswith('No'):
                                    drill_df_oc = drill_df_oc[drill_df_oc['MET OC-SLA'] == 'No']
                                elif col_name.startswith('Yes'):
                                    drill_df_oc = drill_df_oc[drill_df_oc['MET OC-SLA'] == 'Yes']
                                nums = re.findall(r'\((\d+\.?\d*)\)', col_name)
                                if nums: drill_df_oc = drill_df_oc[
                                    drill_df_oc['OC-SLA'].astype(float) == float(nums[0])]
                            show_drilldown_modal(
                                f"💡 Showing **Processing Time** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_oc)

                    elif cells_pub_oc != st.session_state.oc_pub_state['pub']:
                        st.session_state.oc_pub_state['pub'] = cells_pub_oc
                        if cells_pub_oc:
                            cell = cells_pub_oc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_oc_pub, raw_col_name)
                            agent = raw_pivot_oc_pub['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_oc = df_valid_oc_pub.copy().dropna(subset=['JPR'])
                            if agent != 'Total': drill_df_oc = drill_df_oc[drill_df_oc['Name of Agent (OC)'] == agent]
                            if col_name not in ['Total', 'Name of Agent (OC)']: drill_df_oc = drill_df_oc[
                                drill_df_oc['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **Offer Creation** data for: **{agent}** | Time Group: **{col_name}**",
                                drill_df_oc)

            # --------------------------
            # 模块 4: FC
            # --------------------------
            elif selected_nav == "FC":
                st.header("Funding Check (FC) Reports")
                st.subheader("1. Funding Check Volume & Processing Time")

                col_ctrl_fc_left, col_ctrl_fc_right = st.columns([4, 6])
                with col_ctrl_fc_left:
                    small_fc_req = st.radio("SLA & Proc. Time Grouping:", small_table_options, horizontal=True,
                                            key="fc_req_small")
                with col_ctrl_fc_right:
                    period_fc_req = st.radio("Main Table Time Grouping:", period_options, horizontal=True,
                                             key="fc_req_period")

                result_fc_req = process_fc_request(df, period_fc_req, small_fc_req)
                if result_fc_req[0] is not None:
                    styled_fc_req, raw_pivot_fc_req, df_valid_fc_req, sla_styler_fc, sla_pivot_fc, proc_styler_fc, proc_pivot_fc, date_range_str_fc, df_recent_4w_fc = result_fc_req

                    col_left_fc, col_right_fc = st.columns([4, 6])
                    with col_left_fc:
                        st.markdown(f"**FC SLA Meeting Rate ({date_range_str_fc})**")
                        event_sla_fc = st.dataframe(sla_styler_fc, use_container_width=True, hide_index=True,
                                                    height=(len(sla_pivot_fc) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell",
                                                    key="table_fc_sla") if sla_styler_fc is not None else None
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown(f"**FC Processing Time ({date_range_str_fc})**")
                        event_proc_fc = st.dataframe(proc_styler_fc, use_container_width=True, hide_index=True,
                                                     height=(len(proc_pivot_fc) * 35) + 38, on_select="rerun",
                                                     selection_mode="single-cell",
                                                     key="table_fc_proc") if proc_styler_fc is not None else None
                    with col_right_fc:
                        st.markdown("**Funding Check Volume & Avg Processing Time**")
                        event_fc_req = st.dataframe(styled_fc_req, use_container_width=True, hide_index=True,
                                                    height=(len(raw_pivot_fc_req) * 35) + 38, on_select="rerun",
                                                    selection_mode="single-cell", key="table_fc_req_main")

                    if 'fc_req_state' not in st.session_state: st.session_state.fc_req_state = {'sla': [], 'proc': [],
                                                                                                'req': []}
                    cells_sla_fc = getattr(event_sla_fc.selection, 'cells', []) if event_sla_fc else []
                    cells_proc_fc = getattr(event_proc_fc.selection, 'cells', []) if event_proc_fc else []
                    cells_req_fc = getattr(event_fc_req.selection, 'cells', []) if event_fc_req else []

                    if cells_sla_fc != st.session_state.fc_req_state['sla']:
                        st.session_state.fc_req_state['sla'] = cells_sla_fc
                        if cells_sla_fc:
                            cell = cells_sla_fc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(sla_pivot_fc, raw_col_name)
                            agent = sla_pivot_fc['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_fc = df_recent_4w_fc.copy().dropna(subset=['JPR'])
                            drill_df_fc = drill_df_fc[drill_df_fc['MET FC-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_fc = drill_df_fc[
                                drill_df_fc['Name of Agent (OC)'] == agent]
                            if col_name in ['No', 'Yes']: drill_df_fc = drill_df_fc[
                                drill_df_fc['MET FC-SLA'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **SLA Meeting Rate** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_fc)

                    elif cells_proc_fc != st.session_state.fc_req_state['proc']:
                        st.session_state.fc_req_state['proc'] = cells_proc_fc
                        if cells_proc_fc:
                            cell = cells_proc_fc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(proc_pivot_fc, raw_col_name)
                            agent = proc_pivot_fc['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_fc = df_recent_4w_fc.copy().dropna(subset=['JPR', 'FC-SLA'])
                            drill_df_fc = drill_df_fc[drill_df_fc['MET FC-SLA'].isin(['No', 'Yes'])]
                            if agent != 'Grand Total': drill_df_fc = drill_df_fc[
                                drill_df_fc['Name of Agent (OC)'] == agent]
                            if col_name not in ['Grand Total', 'Name of Agent (OC)']:
                                if col_name.startswith('No'):
                                    drill_df_fc = drill_df_fc[drill_df_fc['MET FC-SLA'] == 'No']
                                elif col_name.startswith('Yes'):
                                    drill_df_fc = drill_df_fc[drill_df_fc['MET FC-SLA'] == 'Yes']
                                nums = re.findall(r'\((\d+\.?\d*)\)', col_name)
                                if nums: drill_df_fc = drill_df_fc[
                                    drill_df_fc['FC-SLA'].astype(float) == float(nums[0])]
                            show_drilldown_modal(
                                f"💡 Showing **Processing Time** data for: **{agent}** | Column: **{col_name}**",
                                drill_df_fc)

                    elif cells_req_fc != st.session_state.fc_req_state['req']:
                        st.session_state.fc_req_state['req'] = cells_req_fc
                        if cells_req_fc:
                            cell = cells_req_fc[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]
                            raw_col_name = cell.get("column") if isinstance(cell, dict) else cell[1]
                            col_name = robust_column_matcher(raw_pivot_fc_req, raw_col_name)
                            agent = raw_pivot_fc_req['Name of Agent (OC)'].iloc[row_idx]

                            drill_df_fc = df_valid_fc_req.copy().dropna(subset=['JPR'])
                            if agent != 'Total': drill_df_fc = drill_df_fc[drill_df_fc['Name of Agent (OC)'] == agent]
                            if col_name not in ['Total', 'Name of Agent (OC)']: drill_df_fc = drill_df_fc[
                                drill_df_fc['Time Group'] == col_name]
                            show_drilldown_modal(
                                f"💡 Showing **Funding Check** data for: **{agent}** | Time Group: **{col_name}**",
                                drill_df_fc)

            # --------------------------
            # 模块 5: KPI
            # --------------------------
            elif selected_nav == "RAS-KPI":
                st.header("RAS Key Performance Indicators (KPI)")
                st.subheader("1. End-to-End Processing Time Analysis")

                period_kpi = st.radio("Time Grouping (KPI Analysis):", period_options, horizontal=True,
                                      key="kpi_period")
                result_kpi = process_ras_kpi(df, period_kpi)

                if result_kpi and result_kpi[0] is not None:
                    styled_kpi, raw_kpi, df_valid_kpi = result_kpi

                    col_kpi, _ = st.columns([8, 2])
                    with col_kpi:
                        st.markdown(
                            """
                            <div style="background-color: #F8F9FA; border-left: 4px solid #4285F4; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                                <div style="font-size: 14px; color: #333; margin-bottom: 12px;">
                                    <b>💡 Note:</b><br>
                                    <span style="color: #555;">• <b>RAS KPI</b> = Avg LL-SLA + Avg OC-SLA + Avg FC-SLA</span><br>
                                    <span style="color: #555;">• <b>Outside RAS</b> = Avg NWD (VA to FC) - RAS KPI</span>
                                </div>
                                <div style="font-size: 14px; color: #333; margin-bottom: 8px;"><b>🎯 Target SLA Reference:</b></div>
                                <div style="display: flex; flex-wrap: wrap; gap: 10px;">
                                    <div style="background-color: #E8F0FE; color: #1967D2; padding: 4px 10px; border-radius: 4px; font-size: 13px; border: 1px solid #D2E3FC;"><b>LL-SLA:</b> 3-5 Work Days</div>
                                    <div style="background-color: #E8F0FE; color: #1967D2; padding: 4px 10px; border-radius: 4px; font-size: 13px; border: 1px solid #D2E3FC;"><b>OC-SLA:</b> 2 Work Days</div>
                                    <div style="background-color: #E8F0FE; color: #1967D2; padding: 4px 10px; border-radius: 4px; font-size: 13px; border: 1px solid #D2E3FC;"><b>FC-SLA:</b> 2 Work Days</div>
                                    <div style="background-color: #FFF3CD; color: #856404; padding: 4px 10px; border-radius: 4px; font-size: 13px; border: 1px solid #FFEEBA;"><b>RAS KPI:</b> 9 Work Days</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True
                        )
                        event_kpi = st.dataframe(styled_kpi, use_container_width=True, hide_index=True,
                                                 height=(len(raw_kpi) * 35) + 38, on_select="rerun",
                                                 selection_mode="single-cell", key="table_ras_kpi")

                    if 'kpi_state' not in st.session_state: st.session_state.kpi_state = []
                    cells_kpi = getattr(event_kpi.selection, 'cells', [])
                    if cells_kpi != st.session_state.kpi_state:
                        st.session_state.kpi_state = cells_kpi
                        if cells_kpi:
                            cell = cells_kpi[0]
                            row_idx = cell.get("row") if isinstance(cell, dict) else cell[0]

                            selected_time_group = raw_kpi['Time Group'].iloc[row_idx]
                            drill_df_kpi = df_valid_kpi.copy().dropna(subset=['JPR'])
                            if selected_time_group != 'Total': drill_df_kpi = drill_df_kpi[
                                drill_df_kpi['Time Group'] == selected_time_group]
                            show_drilldown_modal(f"💡 Showing raw data for Time Group: **{selected_time_group}**",
                                                 drill_df_kpi)

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray; font-size: 14px; margin-top: 20px;'>&copy; 2026 Li Zihan. All rights reserved. <br><span style='font-size: 12px;'>RAS Monitoring Report Analysis Dashboard v2.0</span></div>",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()