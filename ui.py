import streamlit as st
import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import date
from datetime import datetime

def setup_sidebar():
    with st.sidebar:
        st.title("功能简介")
        st.markdown("---")
        st.markdown("- 晶圆文件处理")
        
def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    # ✅ 合并上传框：所有主+明细文件统一上传
    st.subheader("📁 上传晶圆文件")
    all_cp_files = st.file_uploader(
        "关键字：华虹/先进/DB/上华（支持多选）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="all_cp_files"
    )

    uploaded_cp_files = {}
    if all_cp_files:
        for file in all_cp_files:
            uploaded_cp_files[file.name] = file
        st.success(f"✅ 共上传 {len(uploaded_cp_files)} 个文件：")
        st.write(list(uploaded_cp_files.keys()))
    else:
        st.info("📂 尚未上传文件。")

    # 📁 上传辅助文件
    st.subheader("📁 上传辅助文件（如无更新可跳过）")
    unfulfilled_file = st.file_uploader("📄 上传未交订单文件", type="xlsx", key="unfulfilled")
    cp_wip_file = st.file_uploader("🧪 上传 CP 在制文件", type="xlsx", key="cp_wip")
    wafer_inventory_file = st.file_uploader("💾 上传晶圆库存文件", type="xlsx", key="wafer_inventory")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")

    # 🚀 生成按钮
    start = st.button("🚀 生成汇总 Excel")

    return (
        uploaded_cp_files,
        forecast_file,
        safety_file,
        unfulfilled_file,
        cp_wip_file,
        wafer_inventory_file,
        start
    )
