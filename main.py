import streamlit as st
import pandas as pd

st.set_page_config(layout="wide", page_title="DIALER PRODUCTIVITY PER CRITERIA OF BALANCE", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode
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
    </style>
    """,
    unsafe_allow_html=True
)

st.title('SPM MONITORING')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Move file uploader to the sidebar
with st.sidebar:
    st.subheader("Upload File")
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Load the data
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
    
    # Filter out rows with excluded roles in Column E (index 4)
    df_filtered = df[~df["Role"].isin(exclude_roles)]
    
    # 1. Per Client and Date Summary
    st.subheader("Summary Report Per Client")
    summary_table = pd.DataFrame(columns=[
        'Date', 'CLIENT', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
        'AVG CONNECTED', 'AVG ACCOUNT', 'AVG TALKTIME'
    ])
    
    grouped = df_filtered.groupby([df_filtered['Date'].dt.date, df_filtered.iloc[:, 6]])  # Column C for date, Column G for client
    summary_data = []
    for (date, client), group in grouped:
        environment = ', '.join(group.iloc[:, 0].unique())  # Column A
        unique_collectors = group.iloc[:, 4].nunique()  # Column E
        total_connected = group.shape[0]  # Count rows for this specific date and client
        total_accounts = group.iloc[:, 3].nunique()  # Column D
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))  # Column I
        total_talk_time = talk_times.sum()
        
        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        # Calculate averages per collector
        avg_connected = total_connected / unique_collectors
        avg_account = total_accounts / unique_collectors
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
            'AVG CONNECTED': '{:.2f}',
            'AVG ACCOUNT': '{:.2f}'
        }),
        height=500,
        use_container_width=True
    )
    
    st.download_button(
        label="Download Per Client and Date Summary as CSV",
        data=summary_table.to_csv(index=False),
        file_name="dialer_client_date_summary_report.csv",
        mime="text/csv",
    )
    
    # Calculate number of unique days
    unique_days = df_filtered['Date'].dt.date.nunique()
    
    # 2. Overall Summary
    st.subheader("Overall Summary")
    overall_summary = pd.DataFrame(columns=[
        'DATE RANGE', 'TOTAL COLLECTORS', 'TOTAL CONNECTED', 'TOTAL ACCOUNTS', 'TOTAL TALK TIME',
        'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
    ])
    
    min_date = df_filtered['Date'].min().strftime('%B %d, %Y')
    max_date = df_filtered['Date'].max().strftime('%B %d, %Y')
    date_range = f"{min_date} - {max_date}"
    
    total_collectors = df_filtered.iloc[:, 4].nunique()  # Column E
    total_connected = df_filtered.shape[0]  # Total rows across all filtered data
    total_accounts = df_filtered.iloc[:, 3].nunique()  # Column D
    total_talk_time = pd.to_timedelta(df_filtered.iloc[:, 8].astype(str)).sum()
    
    total_seconds = int(total_talk_time.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    # Calculate averages based on daily averages from summary_table
    avg_agents_per_day = summary_table['COLLECTOR'].mean()
    avg_calls_per_day = summary_table['AVG CONNECTED'].mean()  # Average of daily per-collector connected
    avg_accounts_per_day = summary_table['AVG ACCOUNT'].mean()  # Average of daily per-collector accounts
    
    # Convert AVG TALKTIME strings to timedelta for averaging
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
            'AVG AGENTS/DAY': '{:.2f}',
            'AVG CALLS/DAY': '{:.2f}',
            'AVG ACCOUNTS/DAY': '{:.2f}'
        }),
        use_container_width=True
    )
    
    st.download_button(
        label="Download Overall Summary as CSV",
        data=overall_summary.to_csv(index=False),
        file_name="dialer_overall_summary_report.csv",
        mime="text/csv",
    )
    
    # 3. Overall Per Client Summary
    st.subheader("Overall Per Client Summary")
    client_summary = pd.DataFrame(columns=[
        'CLIENT', 'DATE RANGE', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME',
        'AVG AGENTS/DAY', 'AVG CALLS/DAY', 'AVG ACCOUNTS/DAY', 'AVG TALKTIME/DAY'
    ])
    
    grouped_clients = df_filtered.groupby(df_filtered.iloc[:, 6])  # Column G for client
    client_summary_data = []
    for client, group in grouped_clients:
        client_min_date = group['Date'].min().strftime('%B %d, %Y')
        client_max_date = group['Date'].max().strftime('%B %d, %Y')
        client_date_range = f"{client_min_date} - {client_max_date}"
        client_unique_days = group['Date'].dt.date.nunique()
        
        environment = ', '.join(group.iloc[:, 0].unique())  # Column A
        unique_collectors = group.iloc[:, 4].nunique()  # Column E
        total_connected = group.shape[0]
        total_accounts = group.iloc[:, 3].nunique()
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))
        total_talk_time = talk_times.sum()
        
        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        # Calculate averages per client
        client_summary_subset = summary_table[summary_table['CLIENT'] == client]
        avg_agents_per_day = client_summary_subset['COLLECTOR'].mean()
        avg_calls_per_day = client_summary_subset['TOTAL CONNECTED'].mean()
        avg_accounts_per_day = client_summary_subset['TOTAL ACCOUNT'].mean()
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
            'AVG AGENTS/DAY': '{:.2f}',
            'AVG CALLS/DAY': '{:.2f}',
            'AVG ACCOUNTS/DAY': '{:.2f}'
        }),
        height=500,
        use_container_width=True
    )
    
    st.download_button(
        label="Download Overall Per Client Summary as CSV",
        data=client_summary.to_csv(index=False),
        file_name="dialer_overall_client_summary_report.csv",
        mime="text/csv",
    )
else:
    st.info("Please upload an Excel file using the sidebar to generate the report.")

# Add some basic styling to the table
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
