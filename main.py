import pandas as pd
import streamlit as st
from datetime import timedelta

# Set up the page configuration
st.set_page_config(layout="wide", page_title="SPM MONITORING", page_icon="📊", initial_sidebar_state="expanded")

# Function to convert HH:MM:SS to seconds
def time_to_seconds(time_str):
    try:
        h, m, s = map(int, time_str.split(':'))
        return h * 3600 + m * 60 + s
    except:
        return 0

# Function to convert seconds back to HH:MM:SS
def seconds_to_time(seconds):
    return str(timedelta(seconds=seconds))

# Data loading function
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Streamlit file uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Load the data
    df = load_data(uploaded_file)

    # Ensure the required columns exist
    required_columns = ["DATE", "CLIENT", "CONTACT NUMBER", "CUSTOMER", "Talk Time Duration"]
    if not all(col in df.columns for col in required_columns):
        st.error("The Excel file must contain the following columns: DATE, CLIENT, CONTACT NUMBER, CUSTOMER, Talk Time Duration")
    else:
        # Clean the data
        df["CONTACT NUMBER"] = df["CONTACT NUMBER"].astype(str)

        # Calculate Total Call: Count rows where CONTACT NUMBER is numeric
        df["Is_Contact"] = df["CONTACT NUMBER"].str.isnumeric()

        # Calculate Total Talk Time: Convert Talk Time Duration to seconds
        df["Talk Time Seconds"] = df["Talk Time Duration"].apply(time_to_seconds)

        # Group by DATE and CLIENT
        summary = df.groupby(["DATE", "CLIENT"]).agg(
            Total_Call=("Is_Contact", "sum"),  # Sum of rows with numeric contact numbers
            Total_Account=("CUSTOMER", "nunique"),  # Count unique customers
            Total_Talk_Time_Seconds=("Talk Time Seconds", "sum")  # Sum talk time in seconds
        ).reset_index()

        # Convert Total Talk Time from seconds back to HH:MM:SS
        summary["Total Talk Time"] = summary["Total_Talk_Time_Seconds"].apply(seconds_to_time)

        # Drop the temporary seconds column
        summary = summary.drop(columns=["Total_Talk_Time_Seconds"])

        # Display the summary table
        st.write("### Summary Table")
        st.dataframe(summary)

else:
    st.write("Please upload an Excel file to generate the summary table.")
