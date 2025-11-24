#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2025/11/5 20:09
# @Author  : Healer
# @File    : myapp.py
# @Software: PyCharm


import streamlit as st
import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string
from datetime import datetime
import tempfile
import zipfile
from io import BytesIO
import hashlib
import re

# 设置页面配置（必须在其他streamlit命令之前）
st.set_page_config(
    page_title="SI Generator Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


class SIGeneratorWeb:
    def __init__(self):
        pass

    def setup_page(self):
        st.title("📊 SI Generator Tool")
        st.markdown("---")

    def run(self):
        self.setup_page()

        # Sidebar for file upload
        with st.sidebar:
            st.header("Upload Files")
            uploaded_files = st.file_uploader(
                "Choose Excel files",
                type=['xlsx', 'xlsm'],
                accept_multiple_files=True,
                help="Upload one or more Excel files containing 'No SI Order' and 'SI Template' sheets"
            )

            st.header("Settings")
            group_column = st.selectbox(
                "Group by Column",
                ["O", "P", "Q", "R", "S", "T"],
                index=0,
                help="Select the column to group data by"
            )

            if st.button("Generate SI Files", type="primary", use_container_width=True):
                if uploaded_files:
                    self.process_files(uploaded_files, group_column)
                else:
                    st.error("Please upload at least one Excel file")

        # Main area for instructions
        st.markdown("""
        ### 📋 Instructions

        1. **Upload Excel Files**: Select one or more Excel files (.xlsx or .xlsm) that contain:
           - 'No SI Order' sheet with data
           - 'SI Template' sheet with formatting

        2. **Select Group Column**: Choose which column to group the data by (default is Column O)

        3. **Generate**: Click the 'Generate SI Files' button to process all files

        ### 📁 Output Structure
        For each input file, the tool will create:
        - Individual SI files for each group
        - A consolidated file with all SIs
        - All files packaged in a ZIP for download
        """)

    def process_files(self, uploaded_files, group_column):
        # 这里放置您的文件处理逻辑
        # 确保使用相对路径，避免绝对路径
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_container = st.container()

        # 您的处理逻辑...
        pass

    # 其他方法保持不变...
    # 确保所有文件路径操作都使用相对路径或tempfile


def main():
    app = SIGeneratorWeb()
    app.run()


if __name__ == "__main__":
    main()