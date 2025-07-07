import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd

from pivot_processor import PivotProcessor
from ui import setup_sidebar, get_uploaded_files
from github_utils import load_file_with_github_fallback
from urllib.parse import quote

def main():
    st.set_page_config(page_title="Excel数据透视汇总工具", layout="wide")
    setup_sidebar()

    # 获取上传文件
    uploaded_cp_files, forecast_file, safety_file, unfulfilled_file, cp_wip_file, wafer_inventory_file, start = get_uploaded_files()

    if start:
        # 初始化处理器
        buffer = BytesIO()
        processor = PivotProcessor()

        # 将所有上传的辅助文件打包成一个 dict（便于传入 processor）
        additional_files = {
            "forecast": forecast_file,
            "safety": safety_file,
            "unfulfilled": unfulfilled_file,
            "cp_wip": cp_wip_file,
            "wafer_inventory": wafer_inventory_file,
        }

        # 调用处理方法（你可能需要在 PivotProcessor 中添加对这些辅助文件的处理逻辑）
        processor.process(uploaded_cp_files, buffer, additional_files)

        # 下载文件按钮
        file_name = f"FAB-WIP数据汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ 汇总完成！你可以下载结果文件：")
        st.download_button(
            label="📥 下载 Excel 汇总报告",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Sheet 预览
        try:
            buffer.seek(0)
            with pd.ExcelFile(buffer, engine="openpyxl") as xls:
                sheet_names = xls.sheet_names
                tabs = st.tabs(sheet_names)
                for i, sheet_name in enumerate(sheet_names):
                    try:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        with tabs[i]:
                            st.subheader(f"📄 {sheet_name}")
                            st.dataframe(df, use_container_width=True)
                    except Exception as e:
                        with tabs[i]:
                            st.error(f"❌ 无法读取工作表 `{sheet_name}`: {e}")
        except Exception as e:
            st.warning(f"⚠️ 无法预览生成的 Excel 文件：{e}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("❌ Streamlit app crashed:", e)
        traceback.print_exc()
