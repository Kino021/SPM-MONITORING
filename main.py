import pandas as pd
import streamlit as st
import math
from io import BytesIO

# Set up the page configuration
st.set_page_config(layout="wide", page_title="MC06 MONITORING", page_icon="📊", initial_sidebar_state="expanded"import pandas as pd
import streamlit as st
import math
from io import BytesIO

# Set up the page configuration
st.set_page_config(layout="wide", page_title="MC06 MONITORING", page_icon="📊", initial_sidebar_state="expanded")

# Title of the app
st.title('MC06 MONITORING')

# Data loading function with file upload support
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to create a single Excel file with multiple sheets (unchanged)
def create_combined_excel_file(summary_dfs, overall_summary_df, sheet_prefix, main_header_text):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        main_header_format = workbook.add_format({
            'bg_color': '#000080', 'font_color': '#FFFFFF', 'bold': True, 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'font_size': 14
        })
        header_format = workbook.add_format({
            'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True, 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        date_range_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        time_format = workbook.add_format({
            'num_format': 'hh:mm:ss', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })

        for key, summary_df in summary_dfs.items():
            sheet_name = f"{sheet_prefix}_{key[:31]}"
            summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, header=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.merge_range('A1:V1', f"{main_header_text} {key}", main_header_format)
            for col_idx, col in enumerate(summary_df.columns):
                worksheet.write(1, col_idx, col, header_format)
            for row_idx in range(len(summary_df)):
                for col_idx, value in enumerate(summary_df.iloc[row_idx]):
                    if col_idx == 0:
                        worksheet.write(row_idx + 2, col_idx, value, date_format if sheet_prefix == "Summary" else date_range_format)
                    elif col_idx in [12, 14, 15, 16]:
                        worksheet.write(row_idx + 2, col_idx, value, time_format)
                    else:
                        worksheet.write(row_idx + 2, col_idx, value, cell_format)
            for col_idx, col in enumerate(summary_df.columns):
                max_length = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(col_idx, col_idx, max_length + 2)

        overall_summary_df.to_excel(writer, sheet_name=f"Overall_{sheet_prefix}", index=False, startrow=2, header=False)
        worksheet = writer.sheets[f"Overall_{sheet_prefix}"]
        worksheet.merge_range('A1:W1', f"Overall {main_header_text}", main_header_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            worksheet.write(1, col_idx, col, header_format)
        for row_idx in range(len(overall_summary_df)):
            for col_idx, value in enumerate(overall_summary_df.iloc[row_idx]):
                if col_idx == 0:
                    worksheet.write(row_idx + 2, col_idx, value, date_range_format)
                elif col_idx in [14, 16, 17, 18]:
                    worksheet.write(row_idx + 2, col_idx, value, time_format)
                else:
                    worksheet.write(row_idx + 2, col_idx, value, cell_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            max_length = max(overall_summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(col_idx, col_idx, max_length + 2)

    return output.getvalue()

# File uploader for Excel file
uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

# Define columns
col1, col2 = st.columns(2)

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    # Create summary table
    with st.container():
        st.subheader("Daily Summary")
        
        # Debugging: Show column names and first few rows
        st.write("Column names:", df.columns.tolist())
        st.write("First 5 rows of the DataFrame:", df.head())
        
        # Convert date column to datetime with explicit format
        df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y', errors='coerce')
        
        # Debugging: Show unique values and type after conversion
        st.write("Unique values in Date column:", df['Date'].unique())
        st.write("Type of Date column after conversion:", df['Date'].dtype)
        
        # Filter out rows where date conversion failed (NaT)
        df = df[df['Date'].notna()].copy()
        
        # Group by date
        summary_data = df.groupby(df['Date'].dt.date).agg({
            'Collector': 'nunique',  # Total unique collectors
            'Contact Number': lambda x: x.notna().sum(),  # Total calls (count non-empty)
            'Customer': 'nunique',  # Total unique accounts
            'Talk Time Duration': 'sum'  # Total talk time
        }).reset_index()
        
        # Rename columns
        summary_data.columns = ['DATE', 'TOTAL COLLECTOR', 'TOTAL CALL', 'TOTAL ACCOUNT', 'TOTAL TALK TIME']
        
        # Format date column to dd/mm/yyyy for display
        summary_data['DATE'] = summary_data['DATE'].apply(lambda x: x.strftime('%d/%m/%Y'))
        
        # Convert total talk time to proper time format
        summary_data['TOTAL TALK TIME'] = pd.to_timedelta(summary_data['TOTAL TALK TIME']).apply(
            lambda x: f"{int(x.total_seconds() // 3600):02d}:{int((x.total_seconds() % 3600) // 60):02d}:{int(x.total_seconds() % 60):02d}"
        )
        
        # Display the summary table
        st.dataframe(
            summary_data,
            use_container_width=True,
            column_config={
                'DATE': st.column_config.DateColumn(format="DD/MM/YYYY"),
                'TOTAL COLLECTOR': st.column_config.NumberColumn(format="%d"),
                'TOTAL CALL': st.column_config.NumberColumn(format="%d"),
                'TOTAL ACCOUNT': st.column_config.NumberColumn(format="%d"),
                'TOTAL TALK TIME': st.column_config.TextColumn()
            }
        )
        
        # Download button for summary
        summary_dfs = {'Daily': summary_data}
        overall_summary_df = summary_data
        excel_file = create_combined_excel_file(summary_dfs, overall_summary_df, "Summary", "MC06 Monitoring Report")
        st.download_button(
            label="Download Summary Excel",
            data=excel_file,
            file_name="MC06_Monitoring_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Title of the app
st.title('MC06 MONITORING')

# Data loading function with file upload support
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Function to create a single Excel file with multiple sheets (unchanged)
def create_combined_excel_file(summary_dfs, overall_summary_df, sheet_prefix, main_header_text):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        main_header_format = workbook.add_format({
            'bg_color': '#000080', 'font_color': '#FFFFFF', 'bold': True, 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'font_size': 14
        })
        header_format = workbook.add_format({
            'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True, 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        cell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        date_range_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        time_format = workbook.add_format({
            'num_format': 'hh:mm:ss', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })

        for key, summary_df in summary_dfs.items():
            sheet_name = f"{sheet_prefix}_{key[:31]}"
            summary_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2, header=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.merge_range('A1:V1', f"{main_header_text} {key}", main_header_format)
            for col_idx, col in enumerate(summary_df.columns):
                worksheet.write(1, col_idx, col, header_format)
            for row_idx in range(len(summary_df)):
                for col_idx, value in enumerate(summary_df.iloc[row_idx]):
                    if col_idx == 0:
                        worksheet.write(row_idx + 2, col_idx, value, date_format if sheet_prefix == "Summary" else date_range_format)
                    elif col_idx in [12, 14, 15, 16]:
                        worksheet.write(row_idx + 2, col_idx, value, time_format)
                    else:
                        worksheet.write(row_idx + 2, col_idx, value, cell_format)
            for col_idx, col in enumerate(summary_df.columns):
                max_length = max(summary_df[col].astype(str).map(len).max(), len(str(col)))
                worksheet.set_column(col_idx, col_idx, max_length + 2)

        overall_summary_df.to_excel(writer, sheet_name=f"Overall_{sheet_prefix}", index=False, startrow=2, header=False)
        worksheet = writer.sheets[f"Overall_{sheet_prefix}"]
        worksheet.merge_range('A1:W1', f"Overall {main_header_text}", main_header_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            worksheet.write(1, col_idx, col, header_format)
        for row_idx in range(len(overall_summary_df)):
            for col_idx, value in enumerate(overall_summary_df.iloc[row_idx]):
                if col_idx == 0:
                    worksheet.write(row_idx + 2, col_idx, value, date_range_format)
                elif col_idx in [14, 16, 17, 18]:
                    worksheet.write(row_idx + 2, col_idx, value, time_format)
                else:
                    worksheet.write(row_idx + 2, col_idx, value, cell_format)
        for col_idx, col in enumerate(overall_summary_df.columns):
            max_length = max(overall_summary_df[col].astype(str).map(len).max(), len(str(col)))
            worksheet.set_column(col_idx, col_idx, max_length + 2)

    return output.getvalue()

# File uploader for Excel file
uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

# Define columns
col1, col2 = st.columns(2)

if uploaded_file is not None:
    df = load_data(uploaded_file)
    
    # Create summary table
    with st.container():
        st.subheader("Daily Summary")
        
        # Debugging: Show column names and first few rows
        st.write("Column names:", df.columns.tolist())
        st.write("First 5 rows of the DataFrame:", df.head())
        
        # Convert date column (Column C, index 2) to datetime with explicit format
        date_col = pd.to_datetime(df.iloc[:, 2], format='%d-%m-%Y', errors='coerce')
        
        # Debugging: Show unique values and type after conversion
        st.write("Unique values in Date column (Column C):", df.iloc[:, 2].unique())
        st.write("Type of date column before filtering:", df.iloc[:, 2].dtype)
        
        # Filter out rows where date conversion failed (NaT)
        df = df[~date_col.isna()].copy()
        
        # Ensure the column is datetime after filtering
        df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], format='%d-%m-%Y')
        st.write("Type of date column after conversion:", df.iloc[:, 2].dtype)
        
        # Group by date
        summary_data = df.groupby(df.iloc[:, 2].dt.date).agg({
            df.columns[4]: 'nunique',  # Column E - Total unique collectors
            df.columns[7]: lambda x: x.notna().sum(),  # Column H - Total calls (count non-empty)
            df.columns[3]: 'nunique',  # Column D - Total unique accounts
            df.columns[8]: 'sum'  # Column I - Total talk time
        }).reset_index()
        
        # Rename columns
        summary_data.columns = ['DATE', 'TOTAL COLLECTOR', 'TOTAL CALL', 'TOTAL ACCOUNT', 'TOTAL TALK TIME']
        
        # Format date column to dd/mm/yyyy for display
        summary_data['DATE'] = summary_data['DATE'].apply(lambda x: x.strftime('%d/%m/%Y'))
        
        # Convert total talk time to proper time format
        summary_data['TOTAL TALK TIME'] = pd.to_timedelta(summary_data['TOTAL TALK TIME']).apply(
            lambda x: f"{int(x.total_seconds() // 3600):02d}:{int((x.total_seconds() % 3600) // 60):02d}:{int(x.total_seconds() % 60):02d}"
        )
        
        # Display the summary table
        st.dataframe(
            summary_data,
            use_container_width=True,
            column_config={
                'DATE': st.column_config.DateColumn(format="DD/MM/YYYY"),
                'TOTAL COLLECTOR': st.column_config.NumberColumn(format="%d"),
                'TOTAL CALL': st.column_config.NumberColumn(format="%d"),
                'TOTAL ACCOUNT': st.column_config.NumberColumn(format="%d"),
                'TOTAL TALK TIME': st.column_config.TextColumn()
            }
        )
        
        # Download button for summary
        summary_dfs = {'Daily': summary_data}
        overall_summary_df = summary_data
        excel_file = create_combined_excel_file(summary_dfs, overall_summary_df, "Summary", "MC06 Monitoring Report")
        st.download_button(
            label="Download Summary Excel",
            data=excel_file,
            file_name="MC06_Monitoring_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
