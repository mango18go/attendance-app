import streamlit as st
import pandas as pd

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])
flex_file = st.file_uploader("유연근무 엑셀 업로드", type=["xlsx", "xls"])

if attendance_file:
    attendance = pd.read_excel(attendance_file)

    # 컬럼 이름 정리
    attendance.columns = attendance.columns.astype(str).str.strip()

    st.write("현재 근태 엑셀 컬럼 목록:", list(attendance.columns))

    # 컬럼 이름 자동 통일
    attendance = attendance.rename(columns={
        "성명": "이름",
        "직원명": "이름",
        "이름": "이름",
        "출근": "출근시간",
        "출근시각": "출근시간",
        "출근 시간": "출근시간",
        "퇴근": "퇴근시간",
        "퇴근시각": "퇴근시간",
        "퇴근 시간": "퇴근시간"
    })

    required_cols = ["이름", "출근시간", "퇴근시간"]

    missing_cols = [col for col in required_cols if col not in attendance.columns]

    if missing_cols:
        st.error(f"필수 컬럼이 없습니다: {missing_cols}")
        st.stop()

    attendance["출근시간"] = pd.to_datetime(attendance["출근시간"], errors="coerce")
    attendance["퇴근시간"] = pd.to_datetime(attendance["퇴근시간"], errors="coerce")

    if flex_file:
        flex = pd.read_excel(flex_file)
        flex.columns = flex.columns.astype(str).str.strip()

        st.write("현재 유연근무 엑셀 컬럼 목록:", list(flex.columns))

        flex = flex.rename(columns={
            "성명": "이름",
            "직원명": "이름",
            "퇴근": "퇴근기준",
            "퇴근시간": "퇴근기준",
            "퇴근 기준": "퇴근기준"
        })

        if "이름" not in flex.columns or "퇴근기준" not in flex.columns:
            st.error("유연근무 엑셀에는 이름, 퇴근기준 컬럼이 필요합니다.")
            st.stop()

        flex["퇴근기준"] = pd.to_datetime(flex["퇴근기준"], errors="coerce")
        df = pd.merge(attendance, flex, on="이름", how="left")

        # 유연근무 기준이 없는 사람은 기본 18:00 적용
        df["퇴근기준"] = df["퇴근기준"].fillna(pd.to_datetime("18:00"))
    else:
        attendance["퇴근기준"] = pd.to_datetime("18:00")
        df = attendance

    def calc_overtime(row):
        if pd.isna(row["퇴근시간"]) or pd.isna(row["퇴근기준"]):
            return 0

        if row["퇴근시간"] > row["퇴근기준"]:
            return round((row["퇴근시간"] - row["퇴근기준"]).total_seconds() / 3600, 2)

        return 0

    df["추가근무(시간)"] = df.apply(calc_overtime, axis=1)

    st.subheader("결과")
    st.dataframe(df)
