# Import Libraries
import streamlit as st
import pandas as pd

# Streamlit configurations
st.set_page_config(page_title="Sub Balancing - ME App", layout="wide")
hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)

# App info -- sidebar
with st.sidebar:
    st.title("App info:")
    st.write("This app is a QCC entry of Manufacturing Engineering Department.")
    st.write("It is made to easily identify and count the right and left insertions and connector number per sub, giving us the ability to easily see the gap and be able to balance the number of insertions.")
    st.title("How to use:")
    st.write("Drag and drop the excel file of sub balancing raw data and the processed data will be automatically generated.")
    st.title("Developer's Note:")
    st.write("1. Make sure that the file extension is xlsx. If not, open the file and save as .xlsx.")
    st.write("2. Make sure that the columns SubNo, Wi_No, Ins_L, Ins_R, ConnNo_L, ConnNo_R exist.")

# App title
st.title("Sub Balancing App")

# Upload excel file
raw_data = st.file_uploader("Upload sub balancing data", type=["xlsx"])

# Read and process uploaded excel file

if raw_data is not None:
    raw_data2 = pd.read_excel(raw_data)

    # Ordering columns
    raw_data3 = raw_data2[["SubNo", "Wi_No", "Ins_L", "Ins_R", "ConnNo_L", "ConnNo_R"]]

    # Group data by "Sub No"
    grouped_data = raw_data3.groupby("SubNo")

    # Calculate grand total
    grand_total = len(raw_data3)

    # Display tables for each "Sub No"
    for sub_no, group_data in grouped_data:
        count_wire_no = len(group_data)
        percent_of_grand_total = (count_wire_no / grand_total) * 100
        st.subheader(f"Sub No: {sub_no} - Wire Count: {count_wire_no} ({percent_of_grand_total:.2f}% of Total Insertions)")
        st.write(group_data)

        @st.cache
        def convert_df(group_data):
            # IMPORTANT: Cache the conversion to prevent computation on every rerun
            return group_data.to_csv().encode('utf-8')
        
        csv = convert_df(group_data)
        
        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name='Sub Balancing.csv',
            mime='text/csv',
        )
