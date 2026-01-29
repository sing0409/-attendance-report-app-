import streamlit as st
import pandas as pd
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt
import io

設定頁面配置

st.set_page_config(
page_title="出勤報表自動彙整系統",
page_icon="📊",
layout="wide",
initial_sidebar_state="collapsed"
)

自訂 CSS 樣式

st.markdown("""
