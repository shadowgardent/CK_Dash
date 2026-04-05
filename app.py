import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm
from datetime import date
import sys

# --- Streamlit Page Configuration ---
st.set_page_config(layout="wide", page_title="QC Analysis Dashboard", page_icon="📊")

st.title('📊 QC Analysis Dashboard')
st.markdown("Upload your Excel file to analyze Quality Control data and trends.")

# --- Thai Font Setup for Matplotlib ---
# Configure Thai font based on OS
try:
    if sys.platform == 'win32':
        # Windows: Use Tahoma which supports Thai and comes with Windows
        plt.rcParams['font.family'] = 'Tahoma'
    else:
        # Linux/Mac: Try to find Thai font
        font_path = '/usr/share/fonts/truetype/tlwg/Loma.ttf'
        fm.fontManager.addfont(font_path)
        plt.rcParams['font.family'] = 'Loma'
    plt.rcParams['axes.unicode_minus'] = False  # This prevents minus signs from being squares
except FileNotFoundError:
    # Fallback: Use DejaVu Sans which has reasonable Thai support
    plt.rcParams['font.family'] = 'DejaVu Sans'
    plt.rcParams['axes.unicode_minus'] = False
    st.warning("Using fallback font. Thai characters may not display perfectly.")
except Exception as e:
    st.warning(f"Font configuration warning: {e}")

# --- Helper Functions ---
@st.cache_data
def load_data(uploaded_file):
    """Loads and cleans the Excel data."""
    try:
        df_clean = pd.read_excel(uploaded_file, sheet_name='Data', skiprows=3, engine='openpyxl')
        df_clean.columns = df_clean.columns.str.strip()
        df_clean['date'] = pd.to_datetime(df_clean['date'], errors='coerce')
        df_clean = df_clean.dropna(subset=['date'])
        return df_clean
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None

def extract_hour(t):
    """Extracts hour from various time formats."""
    try:
        if isinstance(t, str):
            parts = t.split(':')
            if len(parts) > 0:
                return int(parts[0])
            else:
                return None
        elif pd.notna(t) and hasattr(t, 'hour'): # Check for pd.NaT and then attribute
            return t.hour
        return pd.to_datetime(str(t)).hour
    except (ValueError, TypeError):
        return None

# --- Streamlit UI Components ---
with st.sidebar:
    st.header("Configuration")
    uploaded_file = st.file_uploader("Choose your Excel file", type=["xlsx", "xlsm"])

df_clean = None
if uploaded_file is not None:
    df_clean = load_data(uploaded_file)

    if df_clean is not None:
        st.sidebar.success("File uploaded and processed successfully!")
        min_date_data = df_clean['date'].min().date()
        max_date_data = df_clean['date'].max().date()

        start_date = st.sidebar.date_input('Start Date', value=min_date_data)
        end_date = st.sidebar.date_input('End Date', value=max_date_data)

        if start_date > end_date:
            st.sidebar.error("Error: End date must be after start date.")
            st.stop()

        # Filter data based on selected dates
        date_mask = (df_clean['date'].dt.date >= start_date) & (df_clean['date'].dt.date <= end_date)
        df_filtered_date = df_clean.loc[date_mask].copy()

        if df_filtered_date.empty:
            st.warning("No data found for the selected date range. Please adjust the dates.")
            st.stop()

        # Get unique sites for multiselect
        unique_sites = sorted(df_filtered_date['site'].dropna().unique().tolist())
        selected_sites = st.sidebar.multiselect('Select Site(s) for Analysis', unique_sites, default=unique_sites)

        if not selected_sites:
            st.warning("Please select at least one site for analysis.")
            st.stop()

        st.markdown(f"### Analyzing data from **{start_date}** to **{end_date}** for Site(s): **{', '.join(selected_sites)}**")

        # --- Analysis and Plotting Functions (Adapted for Streamlit) ---

        def plot_bar_summary_st(df, start_d, end_d, site_name=None):
            st.subheader(f'Summary of Inspection Results: {site_name if site_name else "All Sites"}')
            df_plot = df.copy()
            if site_name: # Filter by site if provided
                df_plot = df_plot[df_plot['site'] == site_name]

            if df_plot.empty:
                st.info(f"No data for summary bar chart in the selected range for {site_name if site_name else 'All Sites'}.")
                return

            status_col = 'severity_desc' if 'severity_desc' in df_plot.columns else 'severity'
            summary = df_plot[status_col].value_counts().reset_index()
            summary.columns = ['Status', 'Count']

            fig, ax = plt.subplots(figsize=(10, 6))
            sns.barplot(data=summary, x='Status', y='Count', palette='viridis', ax=ax, hue='Status', legend=False)
            for p in ax.patches:
                ax.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()),
                            ha='center', va='center', xytext=(0, 9), textcoords='offset points')

            ax.set_title(f'Summary of Inspection Results: {start_d} to {end_d}', fontsize=16)
            ax.set_xlabel('Status', fontsize=12)
            ax.set_ylabel('Count', fontsize=12)
            ax.grid(axis='y', linestyle='--', alpha=0.6)
            st.pyplot(fig)
            plt.close(fig)

        def plot_pareto_chart_by_site_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Pareto Chart: Top Defects for Site {site_name}')
            df_site_defects = df_data[df_data['site'] == site_name].copy()

            # Filter out 'ผ่าน' and 'ไม่พบปัญหา'
            df_site_defects = df_site_defects[~df_site_defects['severity_desc'].isin(['ผ่าน'])]
            df_site_defects = df_site_defects[df_site_defects['defect_description'] != 'ไม่พบปัญหา']

            if df_site_defects.empty or df_site_defects['defect_description'].isnull().all():
                st.info(f'No defect data for Site {site_name} in the selected period to generate Pareto chart.')
                return

            total_defects = df_site_defects['defect_description'].count()
            if total_defects == 0:
                st.info(f'No defects found for Site {site_name} after filtering for Pareto chart.')
                return

            pareto_df = df_site_defects['defect_description'].value_counts().reset_index()
            pareto_df.columns = ['Defect', 'Count']
            pareto_df['percent'] = (pareto_df['Count'] / total_defects) * 100
            pareto_df['cumpercent'] = pareto_df['Count'].cumsum() / total_defects * 100
            pareto_df = pareto_df.head(15) # Limit to top 15 defects

            fig, ax1 = plt.subplots(figsize=(14, 7))
            sns.barplot(x='Defect', y='Count', data=pareto_df, ax=ax1, palette='magma', hue='Defect', legend=False)
            ax1.set_title(f'Pareto Chart: Top Defects for Site {site_name} ({start_d} to {end_d})', fontsize=18)
            ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45, ha='right')
            ax1.set_xlabel('Defect Description')
            ax1.set_ylabel('Count')

            for i, p in enumerate(ax1.patches):
                percentage = pareto_df['percent'].iloc[i]
                ax1.annotate(f'{percentage:.1f}%',
                            (p.get_x() + p.get_width() / 2., p.get_height()),
                            ha='center', va='center', xytext=(0, 9), textcoords='offset points', fontsize=10, fontweight='bold')

            ax2 = ax1.twinx()
            ax2.plot(pareto_df['Defect'], pareto_df['cumpercent'], color='red', marker='D', ms=7, linestyle='-')
            for i, txt in enumerate(pareto_df['cumpercent']):
                ax2.annotate(f'{txt:.1f}%',
                             (pareto_df['Defect'].iloc[i], pareto_df['cumpercent'].iloc[i]),
                             xytext=(0, 10), textcoords='offset points', color='red', ha='center', fontweight='bold')

            ax2.axhline(80, color='green', linestyle='--', label='80% Line')
            ax2.set_ylim(0, 110)
            ax2.set_ylabel('Cumulative Percent (%)')
            ax2.legend(loc='upper right')

            plt.tight_layout()
            st.pyplot(fig)
            plt.close(fig)

        def display_summary_by_site_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Summary for Site: {site_name}')
            mask = (df_data['date'].dt.date >= start_d) & (df_data['date'].dt.date <= end_d) & (df_data['site'] == site_name)
            df_period_site = df_data.loc[mask].copy()

            if df_period_site.empty:
                st.info(f"No data for Site {site_name} in the selected period.")
                return

            st.write(f"Total inspections: {len(df_period_site)} times")
            status_summary = df_period_site['severity_desc'].value_counts().to_frame()
            st.dataframe(status_summary.style.set_table_styles([{'selector': 'td', 'props': 'text-align: center;'}, {'selector': 'th', 'props': 'text-align: center;'}]), use_container_width=True)

        def display_line_performance_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Line Performance for Site: {site_name}')
            mask = (df_data['date'].dt.date >= start_d) & (df_data['date'].dt.date <= end_d) & (df_data['site'] == site_name)
            df_period_site = df_data.loc[mask].copy()

            if df_period_site.empty:
                st.info(f"No data for line performance for Site {site_name} in the selected period.")
                return

            line_perf = df_period_site.groupby(['line', df_period_site['severity_desc'].apply(lambda x: 'Pass' if x=='ผ่าน' else 'Defect')]).size().unstack(fill_value=0)
            if 'Pass' not in line_perf.columns: line_perf['Pass'] = 0
            if 'Defect' not in line_perf.columns: line_perf['Defect'] = 0

            total_checks_per_line = line_perf['Pass'] + line_perf['Defect']
            line_perf['Pass Rate (%)'] = (line_perf['Pass'] / total_checks_per_line) * 100
            line_perf['Pass Rate (%)'] = line_perf['Pass Rate (%)'].fillna(0)
            st.dataframe(line_perf.sort_values('Pass Rate (%)', ascending=False).style.set_table_styles([{'selector': 'td', 'props': 'text-align: center;'}, {'selector': 'th', 'props': 'text-align: center;'}]), use_container_width=True)

        def plot_hourly_trend_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Hourly Defect Trend for Site: {site_name}')
            mask = (df_data['date'].dt.date >= start_d) & (df_data['date'].dt.date <= end_d) & (df_data['site'] == site_name)
            df_period_site = df_data.loc[mask].copy()

            df_period_site['hour'] = df_period_site['time'].apply(extract_hour)
            df_defects_only = df_period_site[df_period_site['severity_desc'] != 'ผ่าน'].dropna(subset=['hour']).copy()

            if df_defects_only.empty:
                st.info(f"No defect data for Site {site_name} in the selected period to analyze hourly trends.")
                return

            hourly_defects = df_defects_only.groupby('hour').size().reset_index(name='defect_count')

            if hourly_defects.empty:
                st.info(f"Could not create hourly trend graph for Site {site_name} due to missing or invalid data.")
                return

            fig, ax = plt.subplots(figsize=(12, 6))
            sns.lineplot(data=hourly_defects, x='hour', y='defect_count', marker='o', color='red', linewidth=2.5, ax=ax)
            ax.fill_between(hourly_defects['hour'], hourly_defects['defect_count'], color='red', alpha=0.1)

            ax.set_title(f'Hourly Defect Trend - Site {site_name}', fontsize=16)
            ax.set_xlabel('Hour of Day', fontsize=12)
            ax.set_ylabel('Number of Defects', fontsize=12)
            ax.set_xticks(range(int(hourly_defects['hour'].min()), int(hourly_defects['hour'].max()) + 1))
            ax.grid(True, linestyle='--', alpha=0.5)
            st.pyplot(fig)
            plt.close(fig)

            st.markdown(f"**Top 3 Hours with Most Defects for Site {site_name}:**")
            st.dataframe(hourly_defects.sort_values('defect_count', ascending=False).head(3).style.set_table_styles([{'selector': 'td', 'props': 'text-align: center;'}, {'selector': 'th', 'props': 'text-align: center;'}]), use_container_width=True)

        def plot_line_qc_analysis_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Line and QC Analysis for Site: {site_name}')
            mask = (df_data['date'].dt.date >= start_d) & (df_data['date'].dt.date <= end_d) & (df_data['site'] == site_name)
            df_perf_site = df_data.loc[mask].copy()

            if df_perf_site.empty:
                st.info(f"No data for line and QC analysis for Site {site_name} in the selected period.")
                return

            df_perf_site['Result'] = df_perf_site['severity_desc'].apply(lambda x: 'Pass' if x == 'ผ่าน' else 'Defect')

            # --- Plot 1: Defects by Production Line ---
            st.markdown("#### Defects by Production Line")
            line_summary = df_perf_site.groupby(['line', 'Result']).size().unstack(fill_value=0)
            if 'Pass' not in line_summary.columns: line_summary['Pass'] = 0
            if 'Defect' not in line_summary.columns: line_summary['Defect'] = 0

            total_checks_per_line = line_summary['Pass'] + line_summary['Defect']
            line_summary['Pass Rate (%)'] = (line_summary['Pass'] / total_checks_per_line) * 100
            line_summary['Pass Rate (%)'] = line_summary['Pass Rate (%)'].fillna(0)
            line_summary = line_summary.sort_values('Defect', ascending=False)

            if line_summary.empty:
                st.info(f"No line defect data for Site {site_name}.")
            else:
                fig1, ax1 = plt.subplots(figsize=(15, 6))
                sns.barplot(x=line_summary.index, y=line_summary['Defect'], palette='Reds_r', ax=ax1, hue=line_summary.index, legend=False)
                ax1.set_title(f'Number of Defects by Production Line - Site {site_name}', fontsize=16)
                ax1.set_xlabel('Production Line')
                ax1.set_ylabel('Number of Defects')
                ax1.tick_params(axis='x', rotation=45)

                for p in ax1.patches:
                    ax1.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()),
                                ha='center', va='center', xytext=(0, 7), textcoords='offset points', fontweight='bold')

                plt.tight_layout()
                st.pyplot(fig1)
                plt.close(fig1)

            # --- Plot 2: Top 15 QC Inspectors by Defect Count ---
            st.markdown("#### Top 15 QC Inspectors by Defect Count")
            qc_summary = df_perf_site.groupby(['qc_name', 'Result']).size().unstack(fill_value=0)
            if 'Defect' not in qc_summary.columns: qc_summary['Defect'] = 0
            qc_summary = qc_summary.sort_values('Defect', ascending=False).head(15)

            if qc_summary.empty:
                st.info(f"No QC data for Site {site_name}.")
            else:
                fig2, ax2 = plt.subplots(figsize=(15, 6))
                sns.barplot(x=qc_summary.index, y=qc_summary['Defect'], palette='viridis', ax=ax2, hue=qc_summary.index, legend=False)
                ax2.set_title(f'Top 15 QC Inspectors with Most Defects - Site {site_name}', fontsize=16)
                ax2.set_xlabel('QC Inspector Name')
                ax2.set_ylabel('Number of Defects Detected')
                ax2.tick_params(axis='x', rotation=45)

                for p in ax2.patches:
                    ax2.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2., p.get_height()),
                                ha='center', va='center', xytext=(0, 7), textcoords='offset points', fontweight='bold')

                plt.tight_layout()
                st.pyplot(fig2)
                plt.close(fig2)

        def plot_heatmap_and_top_defects_st(df_data, start_d, end_d, site_name):
            st.subheader(f'Heatmap and Top Defects for Site: {site_name}')

            mask = ((df_data['date'].dt.date >= start_d) & 
                    (df_data['date'].dt.date <= end_d) & 
                    (df_data['severity_desc'] != 'ผ่าน') & 
                    (df_data['site'] == site_name))
            df_defects_period_site = df_data.loc[mask].copy()

            if df_defects_period_site.empty:
                st.info(f"No defect data for Site {site_name} in the selected period to create heatmap.")
                return

            defect_matrix = df_defects_period_site.groupby(['line', 'defect_description']).size().unstack(fill_value=0)

            if defect_matrix.empty:
                st.info(f"No data to create Heatmap for Site {site_name}.")
            else:
                if not df_defects_period_site['defect_description'].value_counts().empty:
                    top_defects_overall_site = df_defects_period_site['defect_description'].value_counts().head(10).index
                    defect_matrix_top = defect_matrix.loc[:, defect_matrix.columns.intersection(top_defects_overall_site)]

                    if not defect_matrix_top.empty:
                        fig, ax = plt.subplots(figsize=(16, 8))
                        sns.heatmap(defect_matrix_top.T, annot=True, fmt='d', cmap='YlOrRd', cbar_kws={'label': 'Count'}, ax=ax)
                        ax.set_title(f'Heatmap: Top Defects by Production Line - Site {site_name}', fontsize=18)
                        ax.set_xlabel('Production Line', fontsize=12)
                        ax.set_ylabel('Defect Description', fontsize=12)
                        st.pyplot(fig)
                        plt.close(fig)
                    else:
                        st.info(f"No Top 10 defect data to create Heatmap for Site {site_name}.")
                else:
                    st.info(f"No defect data to create Heatmap for Site {site_name}.")

            st.markdown(f"#### Top 3 Defects by Line for Site {site_name}")
            unique_lines_in_site = df_defects_period_site['line'].dropna().unique()
            if unique_lines_in_site.size > 0:
                for line in sorted(unique_lines_in_site):
                    line_data = df_defects_period_site[df_defects_period_site['line'] == line]
                    top_3 = line_data['defect_description'].value_counts().head(3)
                    if not top_3.empty:
                        st.markdown(f"**Line: {line}**")
                        for defect, count in top_3.items():
                            st.write(f"- {defect}: {count} times")
                    else:
                        st.markdown(f"**Line: {line}**: No defects found.")
            else:
                st.info("No production lines with defects found for this site.")

        # --- Display Results for Selected Sites ---
        for site in selected_sites:
            st.divider()
            st.header(f'Analysis for Site: {site}')

            # Summary Bar Chart for current site (or overall if only one site selected)
            plot_bar_summary_st(df_filtered_date, start_date, end_date, site)

            # Pareto Chart
            plot_pareto_chart_by_site_st(df_filtered_date, start_date, end_date, site)

            # Site Summary (table)
            display_summary_by_site_st(df_filtered_date, start_date, end_date, site)

            # Line Performance
            display_line_performance_st(df_filtered_date, start_date, end_date, site)

            # Hourly Trend
            plot_hourly_trend_st(df_filtered_date, start_date, end_date, site)

            # Line and QC Analysis
            plot_line_qc_analysis_st(df_filtered_date, start_date, end_date, site)

            # Heatmap and Top Defects
            plot_heatmap_and_top_defects_st(df_filtered_date, start_date, end_date, site)

elif uploaded_file is None:
    st.info("Please upload an Excel file to start the analysis.")
