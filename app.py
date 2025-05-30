# -*- coding: utf-8 -*-
"""
Created on Fri May 30 12:59:52 2025

@author: AsfiyaKhanam
"""

# app.py

import streamlit as st
import time
from automation_test1 import generate_master_datasheet

st.set_page_config(page_title="Master Equipment Sheet Generator", layout="centered")
st.title("📊 Master Equipment Data Sheet Generator")

st.markdown("""
            This tool allows you to upload your equipment datasheet Excel file and generate a Master datasheet
            The parameters in each equipment sheet will be organized into five input categoris:
                -SysCAD Inputs
                -Engineering Inputs
                -Lab/Pilot Inputs
                -Project Constants
                -Vendor Inputs
            Upload your file, click **Generate Master Datasheet**, and download the sheet!""")

uploaded_file = st.file_uploader("Upload your Datasheets Excel (.xlsm)", type=["xlsm"])

if uploaded_file:
    if st.button("Generate Master DataSheet"):
        #with st.spinner("Processing..."):
            #output,filename = generate_master_datasheet(uploaded_file)
        progress = st.progress(0,text="Starting generation...")
        for i in range(1,6):
            time.sleep(0.2) #simulate work
            progress.progress(i*20,text=f"Generating...{i*20}%")
            output,filename = generate_master_datasheet(uploaded_file)
            progress.empty()
        st.success("Done! ✅")
        #with open(output_path, "rb") as f:
        st.download_button("📥 Download Master DataSheet", output, file_name=filename)

