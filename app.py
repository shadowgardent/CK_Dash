import os
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
from datetime import date
import sys

# --- Configuration & Constants ---
CHART_OPTIONS = {
    '📈 Production Line Performance': 'production_line',
    '📊 Pareto Chart (Top Defects)': 'pareto',
    '📋 Summary Table': 'summary',
    '📌 Line Performance Table': 'line_perf',
    '⏰ Hourly Trend': 'hourly_trend',
    '🔍 Line & QC Analysis': 'line_qc',
    '🔥 Heatmap & Top Defects': 'heatmap'
}

# --- 1. Thai Font Setup ---
def setup_thai_font():
    """Configures Matplotlib to support Thai characters."""
    try:
        local_font_path = os.path.join(os.path.dirname(__file__), 'fonts', 'Sarabun-Regular.ttf')
        if os.path.exists(local_font_path):
            fm.fontManager.addfont(local_font_path)
            plt.rcParams['font.family'] = 'Sarabun'
        elif sys.platform == 'win32':
            plt.rcParams['font.family'] = 'Tahoma'
        else:
            font_path = '/usr/share/fonts/truetype/tlwg/Loma.ttf'
            if os.path.exists(font_path):
                fm.fontManager.addfont(font_path)
                plt.rcParams['font.family'] = 'Loma'
            else:
                plt.rcParams['font.family'] = 'DejaVu Sans'
        plt.rcParams['axes.unicode_minus'] = False
    except Exception as e:
        plt.rcParams['font.family'] = 'DejaVu Sans'
        st.sidebar.warning(f"Font configuration warning: {e}")

# --- 2. Data Helper Functions ---
@st.cache_data
def load_data(uploaded_file):
    """Loads and cleans the Excel data."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name='Data', skiprows=3, engine='openpyxl')
        df.columns = df.columns.str.strip()
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        return df.dropna(subset=['date'])
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None

def extract_hour(t):
    """Extracts hour from various time formats."""
    try:
        if isinstance(t, str):
            return int(t.split(':')[0]) if ':' in t else None
        if pd.notna(t) and hasattr(t, 'hour'):
            return t.hour
        return pd.to_datetime(str(t)).hour
    except:
        return None

def get_location_col(df):
    """Safely identifies the location description column due to potential typos."""
    possible_names = ['location_description', 'location_desc๐ription']
    for name in possible_names:
        if name in df.columns:
            return name
    return next((col for col in df.columns if 'location' in col.lower()), None)

# --- 3. Analysis Logic ---
def get_line_perf_data(df, start_d, end_d, site_name):
    mask = (df['date'].dt.date >= start_d) & (df['date'].dt.date <= end_d) & (df['site'] == site_name)
    df_plot = df.loc[mask].copy()
    if df_plot.empty:
        return None, None, None

    # Performance grouping
    line_perf = df_plot.groupby(['line', df_plot['severity_desc'].apply(lambda x: 'Pass' if x=='ผ่าน' else 'Defect')]).size().unstack(fill_value=0)
    for col in ['Pass', 'Defect']:
        if col not in line_perf.columns: line_perf[col] = 0

    line_perf['Total_Units'] = line_perf['Pass'] + line_perf['Defect']
    line_perf['Pass Rate (%)'] = (line_perf['Pass'] / line_perf['Total_Units'] * 100).fillna(0)
    line_perf_sorted = line_perf.sort_values('Total_Units', ascending=False)

    # Defect breakdown grouping
    df_defects_only = df_plot[df_plot['severity_desc'] != 'ผ่าน'].copy()
    if not df_defects_only.empty:
        defect_breakdown = df_defects_only.groupby(['line', 'severity_desc']).size().unstack(fill_value=0)
        all_defect_types = [s for s in df['severity_desc'].dropna().unique() if s != 'ผ่าน']
        for col in all_defect_types:
            if col not in defect_breakdown.columns: defect_breakdown[col] = 0
        defect_breakdown = defect_breakdown[all_defect_types].reindex(line_perf_sorted.index, fill_value=0)
    else:
        defect_breakdown = pd.DataFrame(index=line_perf_sorted.index)
    
    return line_perf, line_perf_sorted, defect_breakdown

# --- 4. Plotting Functions ---

def plot_production_line_performance(df, periods, site_name):
    st.subheader(f"Production Line Performance - {site_name}")
    n_periods = len(periods)
    fig, axes = plt.subplots(2, n_periods, figsize=(12 * n_periods, 12))
    if n_periods == 1: axes = axes.reshape(2, 1)

    for i, (start, end, label) in enumerate(periods):
        _, lps, db = get_line_perf_data(df, start, end, site_name)
        
        # Row 1: Total Units
        ax_top = axes[0, i]
        if lps is not None and not lps.empty:
            sns.barplot(x=lps.index, y='Total_Units', data=lps, palette='deep', ax=ax_top, hue=lps.index, legend=False)
            for p in ax_top.patches:
                ax_top.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()), 
                                ha='center', va='bottom', xytext=(0, 5), textcoords='offset points')
            ax_top.set_title(f'จำนวนเครื่องที่ผลิตได้ ({label})', fontsize=14)
            ax_top.tick_params(axis='x', rotation=45)
        else:
            ax_top.text(0.5, 0.5, 'No Data', ha='center', va='center', transform=ax_top.transAxes)

        # Row 2: Defect Breakdown
        ax_bot = axes[1, i]
        if db is not None and not db.empty and not db.columns.empty:
            db.plot(kind='bar', stacked=True, ax=ax_bot, cmap='Paired')
            ax_bot.set_title(f'ประเภทปัญหาที่พบ ({label})', fontsize=14)
            ax_bot.tick_params(axis='x', rotation=45)
            ax_bot.legend(title='ประเภทปัญหา', bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=8)
            for container in ax_bot.containers:
                for p in container.patches:
                    if p.get_height() > 0:
                        ax_bot.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width()/2., p.get_y() + p.get_height()/2.), 
                                        ha='center', va='center', fontsize=8, fontweight='bold')
        else:
            ax_bot.text(0.5, 0.5, 'No Defect Data', ha='center', va='center', transform=ax_bot.transAxes)

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

def plot_pareto_chart(df, periods, site_name):
    st.subheader(f"Pareto Chart: Top Defects - {site_name}")
    n_periods = len(periods)
    fig, axes = plt.subplots(1, n_periods, figsize=(12 * n_periods, 7))
    if n_periods == 1: axes = [axes]

    for i, (start, end, label) in enumerate(periods):
        mask = (df['date'].dt.date >= start) & (df['date'].dt.date <= end) & (df['site'] == site_name)
        df_site = df.loc[mask]
        df_defects = df_site[~df_site['severity_desc'].isin(['ผ่าน'])]
        df_defects = df_defects[df_defects['defect_description'] != 'ไม่พบปัญหา']
        ax = axes[i]

        if df_defects.empty:
            ax.text(0.5, 0.5, 'No Data', ha='center', va='center', transform=ax.transAxes)
            ax.set_title(f'Pareto Chart ({label})')
            continue

        counts = df_defects['defect_description'].value_counts().head(15).reset_index()
        counts.columns = ['Defect', 'Count']
        total = counts['Count'].sum() if not counts.empty else 1
        counts['percent'] = (counts['Count'] / len(df_defects)) * 100
        counts['cumpercent'] = counts['Count'].cumsum() / len(df_defects) * 100

        sns.barplot(x='Defect', y='Count', data=counts, ax=ax, palette='magma', hue='Defect', legend=False)
        ax.set_title(f'Pareto Chart ({label})', fontsize=14)
        ax.tick_params(axis='x', rotation=45, labelsize=9)
        
        for idx, p in enumerate(ax.patches):
            ax.annotate(f"{counts['percent'].iloc[idx]:.1f}%", (p.get_x() + p.get_width()/2., p.get_height()),
                        ha='center', va='bottom', xytext=(0, 5), textcoords='offset points', fontsize=8, fontweight='bold')

        ax2 = ax.twinx()
        ax2.plot(counts['Defect'], counts['cumpercent'], color='red', marker='D', ms=5)
        ax2.axhline(80, color='green', linestyle='--')
        ax2.set_ylim(0, 110)
        ax2.set_ylabel('Cumulative %')

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

def plot_hourly_trend(df, periods, site_name):
    st.subheader(f"Hourly Defect Trend - {site_name}")
    fig, ax = plt.subplots(figsize=(15, 6))
    colors = ['blue', 'red']
    has_data = False

    for i, (start, end, label) in enumerate(periods):
        mask = (df['date'].dt.date >= start) & (df['date'].dt.date <= end) & (df['site'] == site_name)
        df_site = df.loc[mask].copy()
        df_site['hour'] = df_site['time'].apply(extract_hour)
        h_data = df_site[df_site['severity_desc'] != 'ผ่าน'].dropna(subset=['hour']).groupby('hour').size().reset_index(name='count')
        
        if not h_data.empty:
            sns.lineplot(data=h_data, x='hour', y='count', marker='o', label=label, color=colors[i % 2], ax=ax)
            has_data = True
    
    if has_data:
        ax.set_title(f'Hourly Defect Trend - {site_name}', fontsize=16)
        ax.set_xlabel('Hour of Day')
        ax.set_ylabel('Number of Defects')
        ax.grid(True, linestyle='--', alpha=0.5)
        ax.legend()
        st.pyplot(fig)
    else:
        st.info("No hourly defect data found.")
    plt.close(fig)

def plot_line_qc_analysis(df, periods, site_name):
    st.subheader(f"Line & QC Analysis - {site_name}")
    loc_col = get_location_col(df)

    for start, end, label in periods:
        st.markdown(f"#### Analysis for {label}")
        mask = (df['date'].dt.date >= start) & (df['date'].dt.date <= end) & (df['site'] == site_name)
        df_site = df.loc[mask].copy()
        if df_site.empty: continue
        
        df_site['IsDefect'] = df_site['severity_desc'] != 'ผ่าน'
        
        col1, col2 = st.columns(2)
        with col1:
            line_defects = df_site[df_site['IsDefect']].groupby('line').size().sort_values(ascending=False)
            if not line_defects.empty:
                fig1, ax1 = plt.subplots()
                sns.barplot(x=line_defects.index, y=line_defects.values, palette='Reds_r', ax=ax1, hue=line_defects.index, legend=False)
                ax1.set_title(f"Defects by Line ({label})")
                plt.xticks(rotation=45)
                for p in ax1.patches:
                    ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width()/2., p.get_height()), ha='center', va='bottom')
                st.pyplot(fig1)
                plt.close(fig1)
            else:
                st.write("No line defects found.")

        with col2:
            qc_defects = df_site[df_site['IsDefect']].groupby('qc_name').size().sort_values(ascending=False).head(15)
            if not qc_defects.empty:
                fig2, ax2 = plt.subplots()
                sns.barplot(x=qc_defects.index, y=qc_defects.values, palette='viridis', ax=ax2, hue=qc_defects.index, legend=False)
                ax2.set_title(f"Top 15 QC Inspectors ({label})")
                plt.xticks(rotation=45)
                for p in ax2.patches:
                    ax2.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width()/2., p.get_height()), ha='center', va='bottom')
                st.pyplot(fig2)
                plt.close(fig2)
            else:
                st.write("No QC inspector data.")

        if loc_col and 'machine' in df_site.columns:
            df_det = df_site[df_site['IsDefect']].dropna(subset=['line', loc_col, 'machine'])
            if not df_det.empty:
                st.markdown(f"**การแจกแจงปัญหาโดยละเอียด - {label} (Location & Machine)**")
                detailed = df_det.groupby(['line', loc_col, 'machine']).size().reset_index(name='count')
                lines = sorted(detailed['line'].unique())
                n_lines = len(lines)
                n_cols = 2
                n_rows = (n_lines + n_cols - 1) // n_cols
                
                fig3, axes3 = plt.subplots(n_rows, n_cols, figsize=(16, 5 * n_rows))
                axes3 = axes3.flatten() if n_lines > 1 else [axes3]
                
                for idx, line in enumerate(lines):
                    sns.barplot(data=detailed[detailed['line'] == line], x=loc_col, y='count', hue='machine', ax=axes3[idx], palette='Paired')
                    axes3[idx].set_title(f"Line: {line}")
                    axes3[idx].tick_params(axis='x', rotation=45)
                    for p in axes3[idx].patches:
                        if p.get_height() > 0:
                            axes3[idx].annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width()/2., p.get_height()), ha='center', va='bottom', fontsize=8)
                
                for i in range(n_lines, len(axes3)): axes3[i].set_visible(False)
                plt.tight_layout()
                st.pyplot(fig3)
                plt.close(fig3)

def plot_heatmap_st(df, periods, site_name):
    st.subheader(f"Heatmap: Top Defects - {site_name}")
    n_periods = len(periods)
    fig, axes = plt.subplots(1, n_periods, figsize=(14 * n_periods, 8))
    if n_periods == 1: axes = [axes]

    for i, (start, end, label) in enumerate(periods):
        mask = (df['date'].dt.date >= start) & (df['date'].dt.date <= end) & (df['site'] == site_name) & (df['severity_desc'] != 'ผ่าน')
        df_defects = df.loc[mask].copy()
        ax = axes[i]
        
        if df_defects.empty:
            ax.text(0.5, 0.5, 'No Defect Data', ha='center', va='center', transform=ax.transAxes)
            continue
        
        matrix = df_defects.groupby(['line', 'defect_description']).size().unstack(fill_value=0)
        top_10 = df_defects['defect_description'].value_counts().head(10).index
        matrix_top = matrix.loc[:, matrix.columns.intersection(top_10)]
        
        if not matrix_top.empty:
            sns.heatmap(matrix_top.T, annot=True, fmt='d', cmap='YlOrRd', ax=ax, cbar_kws={'label': 'Count'})
            ax.set_title(f'Heatmap ({label})', fontsize=14)
            ax.tick_params(axis='x', rotation=45)
        else:
            ax.text(0.5, 0.5, 'No data for Top 10 defects', ha='center', va='center', transform=ax.transAxes)

    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

# --- 5. Table & Summary Helpers ---
def display_analysis_results(df, periods, site_name, chart_keys):
    for site in [site_name]:
        if 'production_line' in chart_keys: plot_production_line_performance(df, periods, site)
        if 'pareto' in chart_keys: plot_pareto_chart(df, periods, site)
        if 'hourly_trend' in chart_keys: plot_hourly_trend(df, periods, site)
        if 'line_qc' in chart_keys: plot_line_qc_analysis(df, periods, site)
        if 'heatmap' in chart_keys: plot_heatmap_st(df, periods, site)
        
        if 'summary' in chart_keys or 'line_perf' in chart_keys:
            st.subheader(f"Summary Tables - {site}")
            for start, end, label in periods:
                st.markdown(f"**Period: {label}**")
                mask = (df['date'].dt.date >= start) & (df['date'].dt.date <= end) & (df['site'] == site)
                df_p = df.loc[mask]
                
                if df_p.empty:
                    st.info(f"No data for {label}")
                    continue
                
                col1, col2 = st.columns(2)
                with col1:
                    if 'summary' in chart_keys:
                        st.write(f"Total inspections: {len(df_p)}")
                        st.dataframe(df_p['severity_desc'].value_counts().to_frame(), use_container_width=True)
                with col2:
                    if 'line_perf' in chart_keys:
                        lp, _, _ = get_line_perf_data(df, start, end, site)
                        if lp is not None:
                            st.dataframe(lp[['Pass', 'Defect', 'Pass Rate (%)']].sort_values('Pass Rate (%)', ascending=False).style.format("{:.2f}%", subset=['Pass Rate (%)']), use_container_width=True)
                
                if len(periods) > 1 and 'summary' in chart_keys:
                    loc_col = get_location_col(df)
                    if loc_col and 'machine' in df_p.columns:
                        df_def = df_p[df_p['severity_desc'] != 'ผ่าน'].dropna(subset=['line', loc_col, 'machine', 'defect_description'])
                        if not df_def.empty:
                            st.markdown(f"Detailed Defect Breakdown ({label})")
                            detailed = df_def.groupby(['line', loc_col, 'machine', 'defect_description']).size().reset_index(name='Count')
                            st.dataframe(detailed.sort_values(['line', 'Count'], ascending=[True, False]), use_container_width=True)

# --- Main App ---
def main():
    setup_thai_font()
    st.set_page_config(layout="wide", page_title="QC Analysis Dashboard", page_icon="📊")
    st.title('📊 QC Analysis Dashboard')
    st.markdown("Upload your Excel file to analyze Quality Control data and trends.")

    with st.sidebar:
        st.header("Configuration")
        uploaded_file = st.file_uploader("Choose your Excel file", type=["xlsx", "xlsm"])
        comparison_mode = st.checkbox("Enable Comparison Mode (เปรียบเทียบสองช่วงเวลา)")

    if uploaded_file:
        df_clean = load_data(uploaded_file)
        if df_clean is not None:
            min_d, max_d = df_clean['date'].min().date(), df_clean['date'].max().date()
            
            with st.sidebar:
                if comparison_mode:
                    st.subheader("Period 1")
                    sd1, ed1 = st.date_input('Start (P1)', min_d, key='sd1'), st.date_input('End (P1)', max_d, key='ed1')
                    st.subheader("Period 2")
                    sd2, ed2 = st.date_input('Start (P2)', min_d, key='sd2'), st.date_input('End (P2)', max_d, key='ed2')
                    
                    if sd1 > ed1 or sd2 > ed2:
                        st.error("Error: End date must be after start date.")
                        st.stop()
                        
                    site = st.selectbox('Select Site for Comparison', sorted(df_clean['site'].unique()))
                    selected_sites = [site]
                    periods = [(sd1, ed1, "Period 1"), (sd2, ed2, "Period 2")]
                else:
                    start, end = st.date_input('Start Date', min_d), st.date_input('End Date', max_d)
                    if start > end:
                        st.error("Error: End date must be after start date.")
                        st.stop()
                    
                    sites = sorted(df_clean['site'].unique())
                    selected_sites = st.multiselect('Select Site(s)', sites, default=sites)
                    periods = [(start, end, "Analysis Period")]

                st.markdown("---")
                st.markdown("### 📊 เลือกกราฟที่ต้องการ")
                selected_charts = st.multiselect('Charts to Display', options=list(CHART_OPTIONS.keys()), default=list(CHART_OPTIONS.keys()))
                chart_keys = [CHART_OPTIONS[c] for c in selected_charts]

            if not selected_sites:
                st.warning("Please select at least one site.")
            else:
                for site in selected_sites:
                    st.divider()
                    st.header(f'Analysis for Site: {site}')
                    display_analysis_results(df_clean, periods, site, chart_keys)
    else:
        st.info("Please upload an Excel file to start the analysis.")

if __name__ == "__main__":
    main()
