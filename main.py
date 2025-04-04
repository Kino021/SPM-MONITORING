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

st.title('SPM MONITORING ALL ENVI')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to convert single DataFrame to Excel bytes with formatting
def to_excel_single(df, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Format dates properly before writing if 'Date' column exists
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
                'valign': ' Ascending('vcenter')
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

# Define the MC-to-Client mapping
mc_mapping = {
    'MC 10': ['UBP PREWRITE CURING PL BIF', 'UBP PREWRITE BLUE', 'UBP PREWRITE NTB', 'UBP PREWRITE ORANGE', 'UBP SEEKCAP', 'UBP SME 90DPD', 'UBP RECOV NEWLY', 'UBP RECOV OLD1', 'UBP RECOV OLD2', 'UBP RECOV YOUNG MODEL', 'UBP COMBANK RMG', 'UBP RECOV BLUE'],
    'MC 16': ['FHL SALMON', 'GARNET', 'TONIK', 'HOME CREDIT', 'MAYA', 'UNIONDIGITAL', 'NCCF SALARY LOAN', 'TENDOPAY'],
    'MC 04': ['SBC CARDS CURING B2', 'SBF SALAD', 'SBF PL', 'SBC CARDS & LOAN L6', 'SBC CARDS RECOV L1', 'SBC PL RECOV L1', 'SBC CURING B4'],
    'MC 03': ['PSB AUTO CURING', 'SBC AUTO LOAN CURING', 'CSB REPO', 'MALAYAN SAVINGS BANK', 'BPI AUTO SKIP', 'ORICO AUTO LOAN', 'PSB AUTO', 'BPI AUTO', 'SUMISHO', 'MITSUKOSHI RECOV', 'BDO AUTO RMU', 'CBS AUTO LOAN', 'EWB AUTO', 'TOYOTA LEVEL 3', 'BANK OF MAKATI', 'RCBC AUTO LOAN', 'ROB BIKE', 'ROB AUTO', 'SBF MOTOR LOAN', 'TFS LEVEL 2 REGULAR'],
    'MC 17': ['TALA EDC'],
    'MC 21': ['BPI PL 30DPD SL', 'BPI RBANK CARDS 30DPD SL', 'BPI CARDS 30DPD', 'BPI CARDS XDAYS', 'BPI CARDS 30DPD SL', 'BPI CARDS XDAYS SL', 'RBANK BPI HOMELOAN SL', 'BPI RBANK PL SL', 'BPI PL XDAYS SL', 'BDO AUTO LOAN SL', 'BPI AUTO CURING SL', 'BDO HOMELOAN SL', 'BPI PL 60DPD SL', 'BPI RBANK CARDS 60DPD SL'],
    'MC 18': ['FUSE'],
    'MC 07': ['BDO 30DPD', 'BDO MSME NWOFF', 'BDO MSME WOFF', 'BDO PRU', 'BDO RECO', 'BNB SL'],
    'MC 09': ['PSB', 'CBS PERSONAL LOAN', 'CHINA BANK SAVINGS', 'MALAYAN SB PL', 'CBC', 'MALAYAN SB PLS', 'GLOBE WIRELESS', 'GLOBE GPON', 'GLOBE WIRELINES', 'CSB', 'GLOBE B2B'],
    'MC 06': ['BPI BANKO', 'BPI BANKO 210DPD', 'BPI CARDS', 'RBANK CARDS RECOV', 'BPI BUSINESS LOAN', 'BPI VISAYAS', 'RBANK BUCKET 4', 'BPI MINDANAO', 'BPI RECOV PL', 'BPI PERSONAL LOAN', 'BPI RECOV 3', 'BPI RECOV 1', 'BPI RECOV 2', 'RBANK PERSONAL LOAN', 'RBANK INSTABALE'],
    'MC 02': ['SBF HOMELOAN', 'RBANK BPI HOMELOAN', 'PIF BACOLOD', 'PIF BATANGAS', 'PIF CALAMBA', 'PIF CDO', 'PIF CEBU NORTH', 'PIF CEBU SOUTH', 'PIF DAVAO', 'PIF GENSAN', 'PIF ILOILO', 'PIF MALOLOS', 'PIF NCR', 'PIF PAGADIAN', 'PIF PAMPANGA', 'PIF PANGASINAN', 'PIF TAGUM', 'PIF ZAMBOANGA', 'CBS HOMELOAN', 'PIF FCL', 'PIF BACOLOD PROVIDENT', 'PIF CEBU PROVIDENT', 'PIF ILOILO PROVIDENT', 'PIF MALOLOS PROVIDENT', 'PIF PANGASINAN PROVIDENT', 'PIF ZAMBOANGA PROVIDENT', 'PIF NCR PROVIDENT', 'PIF CALAMBA PROVIDENT', 'PIF PAGADIAN PROVIDENT', 'PIF PAMPANGA PROVIDENT'],
    'MC 08': ['EWB', 'RCBC', 'BANKARD'],
    'MC 05': ['MBTC P1', 'MBTC P2', 'MBTC P4', 'MBTC PA', 'MBTC P90', 'MBTC 30DPD', 'MBTC APR'],
    'MC 11': ['MANULIFE', 'COLLECTIUS', 'AIQON', 'MBA'],
    'MC 12': ['PRELEGAL SBF', 'LEGAL CSB PL', 'PRELEGAL CSB PL']
}

# Reverse mapping for easier lookup (client -> MC)
client_to_mc = {client: mc for mc, clients in mc_mapping.items() for client in clients}

with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df.iloc[:, 2], format='%d-%m-%Y')
    
    # Define roles to exclude
    exclude_roles = [
        "Supervisor",
        "Superuser",
        "Dialer specialist",
        "Supervisor (without Predictive Dialer Monitor)"
    ]
    
    # Filter out rows with excluded roles and zero talk time
    df_filtered = df[~df["Role"].isin(exclude_roles)]
    df_filtered = df_filtered[df_filtered.iloc[:, 8] != "00:00:00"]
    
    # 1. Per Client and Date Summary
    st.subheader("Summary Report Per Client and Date")
    summary_table = pd.DataFrame(columns=[
        'Date', 'CLIENT', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
        'AVG CONNECTED', 'AVG ACCOUNT', 'AVG TALKTIME'
    ])
    
    grouped = df_filtered.groupby([df_filtered['Date'].dt.date, df_filtered.iloc[:, 6]])
    summary_data = []
    for (date, client), group in grouped:
        environment = ', '.join(group.iloc[:, 0].unique())
        unique_collectors = group.iloc[:, 4].nunique()
        total_connected = group.shape[0]
        total_accounts = group.iloc[:, 3].nunique()
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))
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
    
    unique_days = df_filtered['Date'].dt.date.nunique()
    
    # 2. Overall Summary with Header
    st.header("Overall Summary Report")
    overall_summary = pd.DataFrame(columns=[
        'DATE RANGE', 'TOTAL COLLECTORS', 'TOTAL CONNECTED', 'TOTAL ACCOUNTS', 'TOTAL TALK TIME',
        'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
    ])
    
    min_date = df_filtered['Date'].min().strftime('%B %d, %Y')
    max_date = df_filtered['Date'].max().strftime('%B %d, %Y')
    date_range = f"{min_date} - {max_date}"
    
    total_collectors = df_filtered.iloc[:, 4].nunique()
    total_connected = df_filtered.shape[0]
    total_accounts = df_filtered.iloc[:, 3].nunique()
    total_talk_time = pd.to_timedelta(df_filtered.iloc[:, 8].astype(str)).sum()
    
    total_seconds = int(total_talk_time.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    avg_agents_per_day = math.ceil(summary_table['COLLECTOR'].mean()) if summary_table['COLLECTOR'].mean() % 1 >= 0.5 else round(summary_table['COLLECTOR'].mean())
    avg_calls_per_day = math.ceil(summary_table['AVG CONNECTED'].mean()) if summary_table['AVG CONNECTED'].mean() % 1 >= 0.5 else round(summary_table['AVG CONNECTED'].mean())
    avg_accounts_per_day = math.ceil(summary_table['AVG ACCOUNT'].mean()) if summary_table['AVG ACCOUNT'].mean() % 1 >= 0.5 else round(summary_table['AVG ACCOUNT'].mean())
    
    def str_to_timedelta(time_str):
        h, m, s = map(int, time_str.split(':'))
        return pd.Timedelta(hours=h, minutes=m, seconds=s)
    
    avg_talktimes = summary_table['AVG TALKTIME'].apply(str_to_timedelta)
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
    
    # 3. Overall Per Client Summary
    st.subheader("Overall Per Client Summary")
    client_summary = pd.DataFrame(columns=[
        'CLIENT', 'DATE RANGE', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
        'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
    ])
    
    grouped_clients = df_filtered.groupby(df_filtered.iloc[:, 6])
    client_summary_data = []
    for client, group in grouped_clients:
        client_min_date = group['Date'].min().strftime('%B %d, %Y')
        client_max_date = group['Date'].max().strftime('%B %d, %Y')
        client_date_range = f"{client_min_date} - {client_max_date}"
        client_unique_days = group['Date'].dt.date.nunique()
        
        environment = ', '.join(group.iloc[:, 0].unique())
        unique_collectors = group.iloc[:, 4].nunique()
        total_connected = group.shape[0]
        total_accounts = group.iloc[:, 3].nunique()
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))
        total_talk_time = talk_times.sum()
        
        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        client_summary_subset = summary_table[summary_table['CLIENT'] == client]
        avg_agents_per_day = math.ceil(client_summary_subset['COLLECTOR'].mean()) if client_summary_subset['COLLECTOR'].mean() % 1 >= 0.5 else round(client_summary_subset['COLLECTOR'].mean())
        avg_calls_per_day = math.ceil(client_summary_subset['TOTAL CONNECTED'].mean()) if client_summary_subset['TOTAL CONNECTED'].mean() % 1 >= 0.5 else round(client_summary_subset['TOTAL CONNECTED'].mean())
        avg_accounts_per_day = math.ceil(client_summary_subset['TOTAL ACCOUNT'].mean()) if client_summary_subset['TOTAL ACCOUNT'].mean() % 1 >= 0.5 else round(client_summary_subset['TOTAL ACCOUNT'].mean())
        avg_talktime_seconds = total_talk_time.total_seconds() / client_unique_days
        avg_hours = int(avg_talktime_seconds // 3600)
        avg_minutes = int((avg_talktime_seconds % 3600) // 60)
        avg_seconds = int(avg_talktime_seconds % 60)
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
    
    # 4. Overall Per Management Center (MC) Summary
    st.subheader("Overall Per Management Center (MC) Summary")
    mc_summary = pd.DataFrame(columns=[
        'MANAGEMENT CENTER', 'DATE RANGE', 'CLIENTS', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
        'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
    ])
    
    # Add MC column to filtered DataFrame based on client
    df_filtered['MC'] = df_filtered.iloc[:, 6].map(client_to_mc)
    
    # Group by MC
    grouped_mc = df_filtered.groupby('MC')
    mc_summary_data = []
    for mc, group in grouped_mc:
        mc_min_date = group['Date'].min().strftime('%B %d, %Y')
        mc_max_date = group['Date'].max().strftime('%B %d, %Y')
        mc_date_range = f"{mc_min_date} - {mc_max_date}"
        mc_unique_days = group['Date'].dt.date.nunique()
        
        clients = ', '.join(group.iloc[:, 6].unique())
        environment = ', '.join(group.iloc[:, 0].unique())
        unique_collectors = group.iloc[:, 4].nunique()
        total_connected = group.shape[0]
        total_accounts = group.iloc[:, 3].nunique()
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))
        total_talk_time = talk_times.sum()
        
        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        mc_summary_subset = summary_table[summary_table['CLIENT'].isin(group.iloc[:, 6].unique())]
        avg_agents_per_day = math.ceil(mc_summary_subset['COLLECTOR'].mean()) if mc_summary_subset['COLLECTOR'].mean() % 1 >= 0.5 else round(mc_summary_subset['COLLECTOR'].mean())
        avg_calls_per_day = math.ceil(mc_summary_subset['TOTAL CONNECTED'].mean()) if mc_summary_subset['TOTAL CONNECTED'].mean() % 1 >= 0.5 else round(mc_summary_subset['TOTAL CONNECTED'].mean())
        avg_accounts_per_day = math.ceil(mc_summary_subset['TOTAL ACCOUNT'].mean()) if mc_summary_subset['TOTAL ACCOUNT'].mean() % 1 >= 0.5 else round(mc_summary_subset['TOTAL ACCOUNT'].mean())
        avg_talktime_seconds = total_talk_time.total_seconds() / mc_unique_days
        avg_hours = int(avg_talktime_seconds // 3600)
        avg_minutes = int((avg_talktime_seconds % 3600) // 60)
        avg_seconds = int(avg_talktime_seconds % 60)
        avg_talktime_str = f"{avg_hours:02d}:{avg_minutes:02d}:{avg_seconds:02d}"
        
        mc_summary_data.append({
            'MANAGEMENT CENTER': mc,
            'DATE RANGE': mc_date_range,
            'CLIENTS': clients,
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
    
    mc_summary = pd.DataFrame(mc_summary_data)
    st.dataframe(
        mc_summary.style.format({
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
        label="Download Overall Per MC Summary as XLSX",
        data=to_excel_single(mc_summary, "MC_Summary"),
        file_name="dialer_overall_mc_summary_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download-mc-overall"
    )
    
    # 5. Download All Categories Button
    st.subheader("Download All Reports")
    st.download_button(
        label="Download All Categories as XLSX",
        data=to_excel_all(
            [summary_table, overall_summary, client_summary, mc_summary],
            ["Client_Date_Summary", "Overall_Summary", "Client_Summary", "MC_Summary"]
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