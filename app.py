import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import re
from datetime import datetime

# 1. 頁面配置 (這一定要在最上面)
st.set_page_config(page_title="出勤報表系統", page_icon="📊")

# 2. 插入藍色虛線框區塊 (這就是你要的視覺效果)
st.markdown("""
    <style>
    .header-box {
        border: 2px dashed #9fb6ff;
        border-radius: 12px;
        padding: 18px;
        background-color: #f8faff;
        margin-bottom: 20px;
        text-align: center;
    }
    .header-title {
        color: #1e40af;
        font-size: 1.3rem; 
        font-weight: bold;
        margin: 0;
        white-space: nowrap;
    }
    .header-subtitle {
        color: #475569;
        font-size: 0.9rem;
        margin-top: 8px;
        font-weight: 500;
    }
    .header-notice {
        color: #64748b;
        font-size: 0.75rem;
        margin-top: 4px;
        font-style: italic;
    }
    </style>
    
    <div class="header-box">
        <p class="header-title">📊 出勤報表自動彙整系統</p>
        <p class="header-subtitle">社團法人中華民國更生少年關懷協會</p>
        <p class="header-notice">(本系統僅供內部行政作業使用，請勿外流)</p>
    </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. 以下完全是你原本提供的「完整設定」與「功能語法」
# ---------------------------------------------------------

# (請在此處貼上你原本那份包含日期轉換、Word 表格填寫、df 處理的所有程式碼)
# 只要確保原本代碼中的 st.set_page_config 不要重複出現即可。
