import streamlit as st
from aggregate_tantousha import aggregate

st.title("担当者集計アプリ")
st.write("Excel ファイルをアップロードすると、自動で集計します。")

uploaded = st.file_uploader("📂 Excel を選択", type=["xlsx","xlsm","xls"])
if uploaded is not None:
    with st.spinner("集計中…"):
        try:
            total, first = aggregate(uploaded)
        except Exception as e:
            st.error("集計処理中にエラーが発生しました:")
            st.error(str(e))
            st.stop()

    st.header("📊 全体担当者")
    for name, cnt in total.most_common():
        st.write(f"- **{name}**：{cnt}件")

    st.header("🚩 『1本目』担当者")
    for name, cnt in first.most_common():
        st.write(f"- **{name}**：{cnt}件")
