import streamlit as st
import pandas as pd

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])
flex_file = st.file_uploader("유연근무 엑셀 업로드", type=["xlsx", "xls"])

if attendance_file:
    attendance = pd.read_excel(attendance_file)

    attendance["출근시간"] = pd.to_datetime(attendance["출근시간"])
    attendance["퇴근시간"] = pd.to_datetime(attendance["퇴근시간"])

    if flex_file:
        flex = pd.read_excel(flex_file)
        flex["퇴근기준"] = pd.to_datetime(flex["퇴근기준"])
        df = pd.merge(attendance, flex, on="이름", how="left")
    else:
        attendance["퇴근기준"] = pd.to_datetime("18:00")
        df = attendance

    def calc_overtime(row):
        if pd.isna(row["퇴근기준"]):
            return 0
        if row["퇴근시간"] > row["퇴근기준"]:
            return (row["퇴근시간"] - row["퇴근기준"]).total_seconds() / 3600
        return 0

    df["추가근무(시간)"] = df.apply(calc_overtime, axis=1)

    st.subheader("결과")
    st.dataframe(df)
