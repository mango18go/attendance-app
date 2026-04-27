import streamlit as st
import pandas as pd

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])
flex_file = st.file_uploader("유연근무 엑셀 업로드", type=["xlsx", "xls"])

if attendance_file and flex_file:
    attendance = pd.read_excel(attendance_file)
    flex = pd.read_excel(flex_file)

    st.subheader("근태 원본")
    st.dataframe(attendance)

    st.subheader("유연근무 기준")
    st.dataframe(flex)
