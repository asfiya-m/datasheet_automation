# -*- coding: utf-8 -*-
"""
Created on Fri May 30 12:59:52 2025

@author: AsfiyaKhanam
"""

# app.py
from io import BytesIO
import streamlit as st
from populate_syscad_inputs_rev1 import populate_syscad_inputs

st.set_page_config(page_title="Master Equipment Sheet Generator", layout="centered")
st.title("📊 Master Datasheet Automation tool")

st.markdown("""
Welcome to the **Master Equipment Sheet Automation Tool**!  
This tool generate and populate equipment datasheets in two steps:

### Step 1: Generate Master Datasheet  
Upload your raw `.xlsm` input sheet. This tool will automatically:
- Extract equipment sheets
- Detect and organize parameter categories:
  - **SysCAD Inputs**
  - **Engineering Inputs**
  - **Lab/Pilot Inputs**
  - **Project Constants**
  - **Vendor Inputs**
- Apply formatting, sorting, and structure

 Click **Generate Master Datasheet** to begin.

### Step 2: Populate with SysCAD Data  
Once you’ve generated the master sheet, upload it here along with a **SysCAD streamtable Excel file**.  
We’ll:
- Match unit tags and parameter values
- Fill in the **SysCAD Inputs** section for each equipment
- Overwrite units in the master sheet to match the streamtable (no conversions)

📤 Click **Populate SysCAD Inputs** to enrich the datasheet.

""")

# uploaded_file = st.file_uploader("Upload your Datasheets Excel (.xlsm)", type=["xlsm"])
# st.markdown("#### 📤 **Upload Equipment Datasheet (.xlsm)**")
# uploaded_file = st.file_uploader(" ", type=["xlsm"])

uploaded_master = st.file_uploader("Upload your master file", type=["xlsm", "xlsx"])
uploaded_streamtable = st.file_uploader("Upload your SysCAD streamtable Excel", type=["xls", "xlsx"])
if uploaded_master and uploaded_streamtable:
    if st.button("▶️ Populate SysCAD Inputs"):
        # Convert to in-memory streams
        master_bytes = BytesIO(uploaded_master.read())
        stream_bytes = BytesIO(uploaded_streamtable.read())

        # Run the population function
        result,missing_types = populate_syscad_inputs(master_bytes, stream_bytes)
        # Show warning if any equipment types were missing
        if missing_types:
            st.warning(f"⚠️ No streamtable data found for: {', '.join(missing_types)}")

        # Download button
        st.download_button(
            label="📥 Download Populated Master Sheet",
            data=result,
            file_name="Populated_Master.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )