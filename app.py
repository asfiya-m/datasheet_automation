# -*- coding: utf-8 -*-
"""
Created on Fri May 30 12:59:52 2025

@author: AsfiyaKhanam
"""

# app.py

import streamlit as st
from automation_test1 import generate_master_datasheet

st.set_page_config(page_title="Master Equipment Sheet Generator", layout="centered")
st.title("📊 Master Equipment Data Sheet Generator")

uploaded_file = st.file_uploader("Upload your Datasheets Excel (.xlsm)", type=["xlsm"])

if uploaded_file:
    if st.button("Generate Master DataSheet"):
        with st.spinner("Processing..."):
            output,filename = generate_master_datasheet(uploaded_file)
        st.success("Done! ✅")
        #with open(output_path, "rb") as f:
        st.download_button("📥 Download Master DataSheet", output, file_name=filename)

