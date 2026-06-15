import io

import streamlit as st

from tools.generate_kouji_schedule import generate_workbook

st.set_page_config(page_title="工程表ジェネレーター", layout="centered")

st.title("住戸内工事日程表ジェネレーター")
st.write("入居者一覧（.xls）をアップロードすると、住戸内工事日程表（.xlsx）を作成します。")

uploaded = st.file_uploader("入居者一覧 (.xls)", type=["xls"])

if uploaded is not None:
    try:
        wb = generate_workbook(uploaded.read())
    except Exception as e:
        st.error(f"工程表の作成に失敗しました: {e}")
    else:
        building_name = wb.active["A1"].value
        st.success(f"「{building_name}」の工程表を作成しました。")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

        st.download_button(
            label="工程表をダウンロード (.xlsx)",
            data=buf,
            file_name="工程表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
