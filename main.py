import streamlit as st
# Importing useful libraries
import streamlit as st
import utils

# For operating system related tasks
import shutil
import os

# Creating a title and icon to the webpage
st.set_page_config(
    page_title="Sub ID Analysis Automation",
    page_icon="📚"
)

st.header("Sub ID Analysis Automation")

st.subheader("Data Uploading Section")

click_data = st.file_uploader(
    'Please upload the click vs Registration data here',
    type = ['xls']
)

revenue_data = st.file_uploader(
    'Please upload the revenue data here',
    type = ['xls']
)

if st.button("Analyze"):

    temp_click_path = 'temp_click_data.xls'
    temp_revenue_path = 'temp_revenue_data.xls'
    zip_file_path = 'files.zip'
    
    if click_data and revenue_data:

        if not os.path.exists('synthesized_data'): os.mkdir('synthesized_data')
            
        # Save the uploaded file to the local directory
        with open(temp_click_path, "wb") as f:
            f.write(click_data.getbuffer())

        # Save the uploaded file to the local directory
        with open(temp_revenue_path, "wb") as f:
            f.write(revenue_data.getbuffer())

        with st.spinner("Analyzing..."):
            utils.get_report(temp_click_path,temp_revenue_path,zip_file_path)

        st.success('Report generation complete!')
        st.download_button(
            label="Download ZIP",
            data=open(zip_file_path, "rb").read(),
            file_name="generated_report.zip",
            mime="application/zip"
        )

        os.remove(zip_file_path)
        os.remove(temp_click_path)
        os.remove(temp_revenue_path)

    elif click_data and not revenue_data:
        st.warning("Please Upload the Revenue Data to continue")
    elif revenue_data and not click_data:
        st.warning("Please Upload the Click vs Registration Data to continue")
    else: st.warning("Upload data first please!!!")
