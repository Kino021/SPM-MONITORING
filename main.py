import streamlit as st
import pandas as pd
from io import BytesIO
import math

st.set_page_config(layout="wide", page_title="DIALER PRODUCTIVITY PER CRITERIA OF BALANCE", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode and custom header styling
st.markdown(
    """
    <style>
    .reportview-container {
        background: #2E2E2E;
        color: white;
    }
    .sidebar .sidebar-content {
        background: #2E2E2E;
    }
    h1, h2, h3 {
        color: #87CEEB !important;  /* Light blue color */
        font-weight: bold !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('SPM DIALING MONITORING ALL ENVI')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to convert single DataFrame to Excel bytes with formatting
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_copy = df.copy()
        if 'Date' in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy['Date']):
                df_copy['Date'] = df_copy['Date'].dt.strftime('%d-%m-%Y')
            elif pd.api.types.is_object_dtype(df_copy['Date']):
                df_copy['Date'] = pd.to_datetime(df_copy['Date'], errors='coerce').dt.strftime('%d-%m-%Y')
        df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
        
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#87CEEB',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1})
        
        for col_num, value in enumerate(df_copy.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        for row_num in range(1, len(df_copy) + 1):
            for col_num in range(len(df_copy.columns)):
                worksheet.write(row_num, col_num, df_copy.iloc[row_num-1, col_num], cell_format)
        
        for i, col in enumerate(df_copy.columns):
            max_length = max(
                df_copy[col].astype(str).map(len).max(),
                len(str(col))
            )
            worksheet.set_column(i, i, max_length + 2)
    
    return output.getvalue()

# Function to combine all DataFrames into one Excel file
def to_excel_all(dfs, sheet_names):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for df, sheet_name in zip(dfs, sheet_names):
            df_copy = df.copy()
            if 'Date' in df_copy.columns:
                if pd.api.types.is_datetime64_any_dtype(df_copy['Date']):
                    df_copy['Date'] = df_copy['Date'].dt.strftime('%d-%m-%Y')
                elif pd.api.types.is_object_dtype(df_copy['Date']):
                    df_copy['Date'] = pd.to_datetime(df_copy['Date'], errors='coerce').dt.strftime('%d-%m-%Y')
            df_copy.to_excel(writer, index=False, sheet_name=sheet_name)
            
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#87CEEB',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            cell_format = workbook.add_format({'border': 1})
            
            for col_num, value in enumerate(df_copy.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            for row_num in range(1, len(df_copy) + 1):
                for col_num in range(len(df_copy.columns)):
                    worksheet.write(row_num, col_num, df_copy.iloc[row_num-1, col_num], cell_format)
            
            for i, col in enumerate(df_copy.columns):
                max_length = max(
                    df_copy[col].astype(str).map(len).max(),
                    len(str(col))
                )
                worksheet.set_column(i, i, max_length + 2)
    
    return output.getvalue()

with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    # Debug: Display column names and sample of Date column
    st.write("Column names:", df.columns.tolist())
    st.write("Sample of 'Date' column:", df['Date'].head(10).tolist())
    
    # Convert date column with error handling (using column name 'Date')
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
    
    # Check for invalid dates and warn user
    if df['Date'].isna().any():
        st.warning("Some dates could not be parsed. Check these rows:")
        st.write(df[df['Date'].isna()])
    
    # Define roles to exclude
    exclude_roles = [
        "Supervisor",
        "Superuser",
        "Dialer specialist",
        "Supervisor (without Predictive Dialer Monitor)"
    ]
    
    # Filter out rows with excluded roles and zero talk time
    df_filtered = df[~df["Role"].isin(exclude_roles)]
    df_filtered = df_filtered[df_filtered["Talk Time Duration"] != "00:00:00"]
    
    # Check if df_filtered is empty
    if df_filtered.empty:
        st.error("No data remains after filtering. Check your 'Role' and 'Talk Time Duration' columns.")
    else:
        # 1. Per Client and Date Summary
        st.subheader("Summary Report Per Client and Date")
        summary_table = pd.DataFrame(columns=[
            'Date', 'CLIENT', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
            'AVG CONNECTED', 'AVG ACCOUNT', 'AVG TALKTIME'
        ])
        
        grouped = df_filtered.groupby([df_filtered['Date'].dt.date, df_filtered['Client']])
        summary_data = []
        for (date, client), group in grouped:
            environment = ', '.join(group['ENVIRONMENT'].unique())
            unique_collectors = group['Collector'].nunique()
            total_connected = group.shape[0]
            total_accounts = group['Account'].nunique()
            talk_times = pd.to_timedelta(group['Talk Time Duration'].astype(str))
            total_talk_time = talk_times.sum()
            
            total_seconds = int(total_talk_time.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            avg_connected = math.ceil(total_connected / unique_collectors) if (total_connected / unique_collectors) % 1 >= 0.5 else round(total_connected / unique_collectors)
            avg_account = math.ceil(total_accounts / unique_collectors) if (total_accounts / unique_collectors) % 1 >= 0.5 else round(total_accounts / unique_collectors)
            avg_talktime_seconds = total_seconds / unique_collectors
            avg_t_hours = int(avg_talktime_seconds // 3600)
            avg_t_minutes = int((avg_talktime_seconds % 3600) // 60)
            avg_t_seconds = int(avg_talktime_seconds % 60)
            avg_talktime_str = f"{avg_t_hours:02d}:{avg_t_minutes:02d}:{avg_t_seconds:02d}"
            
            summary_data.append({
                'Date': date,
                'CLIENT': client,
                'ENVIRONMENT': environment,
                'COLLECTOR': unique_collectors,
                'TOTAL CONNECTED': total_connected,
                'TOTAL ACCOUNT': total_accounts,
                'TOTAL TALK TIME': total_talk_time_str,
                'AVG CONNECTED': avg_connected,
                'AVG ACCOUNT': avg_account,
                'AVG TALKTIME': avg_talktime_str
            })
        
        summary_table = pd.DataFrame(summary_data)
        st.dataframe(
            summary_table.style.format({
                'Date': '{:%d-%m-%Y}',
                'COLLECTOR': '{:,.0f}',
                'TOTAL CONNECTED': '{:,.0f}',
                'TOTAL ACCOUNT': '{:,.0f}',
                'AVG CONNECTED': '{:.0f}',
                'AVG ACCOUNT': '{:.0f}'
            }),
            height=500,
            use_container_width=True
        )
        
        st.download_button(
            label="Download Per Client and Date Summary as XLSX",
            data=to_excel_single(summary_table, "Client_Date_Summary"),
            file_name="dialer_client_date_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-client-date"
        )
        
        # 2. Summary Per Day
        st.subheader("Summary Report Per Day")
        daily_summary_table = pd.DataFrame(columns=[
            'Date', 'CLIENT', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
            'AVG CONNECTED', 'AVG ACCOUNT', 'AVG TALKTIME'
        ])
        
        daily_grouped = df_filtered.groupby(df_filtered['Date'].dt.date)
        daily_summary_data = []
        for date, group in daily_grouped:
            clients = ', '.join(group['Client'].unique())  # List all unique clients
            environment = ', '.join(group['ENVIRONMENT'].unique())
            unique_collectors = group['Collector'].nunique()
            total_connected = group.shape[0]
            total_accounts = group['Account'].nunique()
            talk_times = pd.to_timedelta(group['Talk Time Duration'].astype(str))
            total_talk_time = talk_times.sum()
            
            total_seconds = int(total_talk_time.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            avg_connected = math.ceil(total_connected / unique_collectors) if (total_connected / unique_collectors) % 1 >= 0.5 else round(total_connected / unique_collectors)
            avg_account = math.ceil(total_accounts / unique_collectors) if (total_accounts / unique_collectors) % 1 >= 0.5 else round(total_accounts / unique_collectors)
            avg_talktime_seconds = total_seconds / unique_collectors
            avg_t_hours = int(avg_talktime_seconds // 3600)
            avg_t_minutes = int((avg_talktime_seconds % 3600) // 60)
            avg_t_seconds = int(avg_talktime_seconds % 60)
            avg_talktime_str = f"{avg_t_hours:02d}:{avg_t_minutes:02d}:{avg_t_seconds:02d}"
            
            daily_summary_data.append({
                'Date': date,
                'CLIENT': clients,
                'ENVIRONMENT': environment,
                'COLLECTOR': unique_collectors,
                'TOTAL CONNECTED': total_connected,
                'TOTAL ACCOUNT': total_accounts,
                'TOTAL TALK TIME': total_talk_time_str,
                'AVG CONNECTED': avg_connected,
                'AVG ACCOUNT': avg_account,
                'AVG TALKTIME': avg_talktime_str
            })
        
        daily_summary_table = pd.DataFrame(daily_summary_data)
        st.dataframe(
            daily_summary_table.style.format({
                'Date': '{:%d-%m-%Y}',
                'COLLECTOR': '{:,.0f}',
                'TOTAL CONNECTED': '{:,.0f}',
                'TOTAL ACCOUNT': '{:,.0f}',
                'AVG CONNECTED': '{:.0f}',
                'AVG ACCOUNT': '{:.0f}'
            }),
            height=500,
            use_container_width=True
        )
        
        st.download_button(
            label="Download Per Day Summary as XLSX",
            data=to_excel_single(daily_summary_table, "Daily_Summary"),
            file_name="dialer_daily_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-daily"
        )
        
        # 3. Overall Summary with Header
        st.header("Overall Summary Report")
        overall_summary = pd.DataFrame(columns=[
            'DATE RANGE', 'TOTAL COLLECTORS', 'TOTAL CONNECTED', 'TOTAL ACCOUNTS', 'TOTAL TALK TIME',
            'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG CONNECTED/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
        ])
        
        min_date = df_filtered['Date'].min()
        max_date = df_filtered['Date'].max()
        if pd.isna(min_date) or pd.isna(max_date):
            date_range = "Invalid date range (check Date column)"
        else:
            min_date_str = min_date.strftime('%B %d, %Y')
            max_date_str = max_date.strftime('%B %d, %Y')
            date_range = f"{min_date_str} - {max_date_str}"

        # Calculate totals from filtered data
        total_collectors = df_filtered['Collector'].nunique()
        total_connected = df_filtered.shape[0]
        total_accounts = df_filtered['Account'].nunique()
        total_talk_time = pd.to_timedelta(df_filtered['Talk Time Duration'].astype(str)).sum()

        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"

        # Use averages from daily_summary_table
        avg_agents_per_day = math.ceil(daily_summary_table['COLLECTOR'].mean()) if daily_summary_table['COLLECTOR'].mean() % 1 >= 0.5 else round(daily_summary_table['COLLECTOR'].mean())
        avg_calls_per_day = math.ceil(daily_summary_table['TOTAL CONNECTED'].mean()) if daily_summary_table['TOTAL CONNECTED'].mean() % 1 >= 0.5 else round(daily_summary_table['TOTAL CONNECTED'].mean())
        avg_connected_per_day = math.ceil(daily_summary_table['AVG CONNECTED'].mean()) if daily_summary_table['AVG CONNECTED'].mean() % 1 >= 0.5 else round(daily_summary_table['AVG CONNECTED'].mean())
        avg_accounts_per_day = math.ceil(daily_summary_table['TOTAL ACCOUNT'].mean()) if daily_summary_table['TOTAL ACCOUNT'].mean() % 1 >= 0.5 else round(daily_summary_table['TOTAL ACCOUNT'].mean())

        def str_to_timedelta(time_str):
            h, m, s = map(int, time_str.split(':'))
            return pd.Timedelta(hours=h, minutes=m, seconds=s)

        avg_talktimes = daily_summary_table['AVG TALKTIME'].apply(str_to_timedelta)
        avg_talktime_per_day = avg_talktimes.mean()
        avg_t_seconds = int(avg_talktime_per_day.total_seconds())
        avg_hours = avg_t_seconds // 3600
        avg_minutes = (avg_t_seconds % 3600) // 60
        avg_seconds = avg_t_seconds % 60
        avg_talktime_str = f"{avg_hours:02d}:{avg_minutes:02d}:{avg_seconds:02d}"

        overall_summary = pd.DataFrame([{
            'DATE RANGE': date_range,
            'TOTAL COLLECTORS': total_collectors,
            'TOTAL CONNECTED': total_connected,
            'TOTAL ACCOUNTS': total_accounts,
            'TOTAL TALK TIME': total_talk_time_str,
            'AVG AGENTS/DAY': avg_agents_per_day,
            'AVG CALLS/DAY': avg_calls_per_day,
            'AVG CONNECTED/DAY': avg_connected_per_day,
            'AVG ACCOUNTS/DAY': avg_accounts_per_day,
            'AVG TALKTIME/DAY': avg_talktime_str
        }])

        st.dataframe(
            overall_summary.style.format({
                'TOTAL COLLECTORS': '{:,.0f}',
                'TOTAL CONNECTED': '{:,.0f}',
                'TOTAL ACCOUNTS': '{:,.0f}',
                'AVG AGENTS/DAY': '{:.0f}',
                'AVG CALLS/DAY': '{:.0f}',
                'AVG CONNECTED/DAY': '{:.0f}',
                'AVG ACCOUNTS/DAY': '{:.0f}'
            }),
            use_container_width=True
        )
        
        st.download_button(
            label="Download Overall Summary as XLSX",
            data=to_excel_single(overall_summary, "Overall_Summary"),
            file_name="dialer_overall_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-overall"
        )
        
        # 4. Overall Per Client Summary
        st.subheader("Overall Per Client Summary")
        client_summary = pd.DataFrame(columns=[
            'CLIENT', 'DATE RANGE', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
            'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
        ])
        
        grouped_clients = df_filtered.groupby(df_filtered['Client'])
        client_summary_data = []
        for client, group in grouped_clients:
            client_min_date = group['Date'].min()
            client_max_date = group['Date'].max()
            if pd.isna(client_min_date) or pd.isna(client_max_date):
                client_date_range = "Invalid date range"
            else:
                client_min_date_str = client_min_date.strftime('%B %d, %Y')
                client_max_date_str = client_max_date.strftime('%B %d, %Y')
                client_date_range = f"{client_min_date_str} - {client_max_date_str}"
            
            environment = ', '.join(group['ENVIRONMENT'].unique())
            unique_collectors = group['Collector'].nunique()
            total_connected = group.shape[0]
            total_accounts = group['Account'].nunique()
            talk_times = pd.to_timedelta(group['Talk Time Duration'].astype(str))
            total_talk_time = talkWait times.sum()
            
            total_seconds = int(total_talk_time.total_seconds())
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            seconds = total_seconds % 60
            total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            client_summary_subset = summary_table[summary_table['CLIENT'] == client]
            avg_agents_per_day = math.ceil(client_summary_subset['COLLECTOR'].mean()) if client_summary_subset['COLLECTOR'].mean() % 1 >= 0.5 else round(client_summary_subset['COLLECTOR'].mean())
            avg_calls_per_day = math.ceil(client_summary_subset['TOTAL CONNECTED'].mean()) if client_summary_subset['TOTAL CONNECTED'].mean() % 1 >= 0.5 else round(client_summary_subset['TOTAL CONNECTED'].mean())
            avg_accounts_per_day = math.ceil(client_summary_subset['TOTAL ACCOUNT'].mean()) if client_summary_subset['TOTAL ACCOUNT'].mean() % 1 >= 0.5 else round(client_summary_subset['TOTAL ACCOUNT'].mean())
            
            client_avg_talktimes = client_summary_subset['AVG TALKTIME'].apply(str_to_timedelta)
            avg_talktime_per_day = client_avg_talktimes.mean()
            if pd.isna(avg_talktime_per_day):
                avg_talktime_str = "00:00:00"
            else:
                avg_t_seconds = int(avg_talktime_per_day.total_seconds())
                avg_hours = avg_t_seconds // 3600
                avg_minutes = (avg_t_seconds % 3600) // 60
                avg_seconds = avg_t_seconds % 60
                avg_talktime_str = f"{avg_hours:02d}:{avg_minutes:02d}:{avg_seconds:02d}"
            
            client_summary_data.append({
                'CLIENT': client,
                'DATE RANGE': client_date_range,
                'ENVIRONMENT': environment,
                'COLLECTOR': unique_collectors,
                'TOTAL CONNECTED': total_connected,
                'TOTAL ACCOUNT': total_accounts,
                'TOTAL TALK TIME': total_talk_time_str,
                'AVG AGENTS/DAY': avg_agents_per_day,
                'AVG CALLS/DAY': avg_calls_per_day,
                'AVG ACCOUNTS/DAY': avg_accounts_per_day,
                'AVG TALKTIME/DAY': avg_talktime_str
            })
        
        client_summary = pd.DataFrame(client_summary_data)
        st.dataframe(
            client_summary.style.format({
                'COLLECTOR': '{:,.0f}',
                'TOTAL CONNECTED': '{:,.0f}',
                'TOTAL ACCOUNT': '{:,.0f}',
                'AVG AGENTS/DAY': '{:.0f}',
                'AVG CALLS/DAY': '{:.0f}',
                'AVG ACCOUNTS/DAY': '{:.0f}'
            }),
            height=500,
            use_container_width=True
        )
        
        st.download_button(
            label="Download Overall Per Client Summary as XLSX",
            data=to_excel_single(client_summary, "Client_Summary"),
            file_name="dialer_overall_client_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-client-overall"
        )
        
        # 5. Download All Categories Button
        st.subheader("Download All Reports")
        st.download_button(
            label="Download All Categories as XLSX",
            data=to_excel_all(
                [summary_table, daily_summary_table, overall_summary, client_summary],
                ["Client_Date_Summary", "Daily_Summary", "Overall_Summary", "Client_Summary"]
            ),
            file_name="dialer_all_categories_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-all"
        )

else:
    st.info("Please upload an Excel file using the sidebar to generate the report.")

# Add basic table styling
st.markdown(
    """
    <style>
    thead tr th {
        color: white !important;
        background-color: #4A4A4A !important;
    }
    tbody tr:nth-child(odd) {
        background-color: #3A3A3A;
    }
    </style>
    """,
    unsafe_allow_html=True
)
