import streamlit as st
import pandas as pd
from io import BytesIO

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])

def extract_end_time(value):
    if pd.isna(value):
        return None
    text = str(value).replace(" ", "")
    if "~" in text:
        return text.split("~")[-1]
    return text

def safe_sheet_name(name):
    invalid = ["\\", "/", "*", "[", "]", ":", "?"]
    for ch in invalid:
        name = str(name).replace(ch, "")
    return name[:31]

if attendance_file:

    # 👉 헤더 문제 해결 (중요)
    attendance = pd.read_excel(attendance_file, sheet_name="가공", header=1)

    # 컬럼 정리
    attendance.columns = attendance.columns.astype(str).str.strip()

    # 컬럼 이름 통일
    attendance = attendance.rename(columns={
        "성명": "사용자명",
        "이름": "사용자명",
        "직원명": "사용자명",
        "출근시간": "첫출근",
        "출근": "첫출근",
        "첫출근시간": "첫출근",
        "퇴근시간": "마지막퇴근",
        "퇴근": "마지막퇴근",
        "최종퇴근": "마지막퇴근"
    })

    st.write("현재 컬럼:", list(attendance.columns))

    # 필수 컬럼 체크
    needed = ["날짜", "사용자명", "첫출근", "마지막퇴근"]
    missing = [c for c in needed if c not in attendance.columns]

    if missing:
        st.error(f"필수 컬럼 없음: {missing}")
        st.stop()

    # 부서 없으면 생성
    if "부서명" not in attendance.columns:
        attendance["부서명"] = ""

    # 날짜/시간 변환
    attendance["날짜"] = pd.to_datetime(attendance["날짜"], errors="coerce").dt.date
    attendance["첫출근"] = pd.to_datetime(attendance["첫출근"], errors="coerce")
    attendance["마지막퇴근"] = pd.to_datetime(attendance["마지막퇴근"], errors="coerce")

    # 👉 하루 1명 기준으로 묶기
    df = attendance.groupby(["날짜", "사용자명", "부서명"], as_index=False).agg({
        "첫출근": "min",
        "마지막퇴근": "max"
    })

    # 기본 퇴근시간
    df["퇴근기준"] = "18:00:00"

    # 👉 유연근무 처리 (있을 때만)
    try:
        flex = pd.read_excel(attendance_file, sheet_name="4월 유연근무", header=0)
        flex.columns = flex.columns.astype(str).str.strip()

        if "성명" in flex.columns and "유연근무시간" in flex.columns:
            flex["퇴근기준"] = flex["유연근무시간"].apply(extract_end_time)
            flex = flex.rename(columns={"성명": "사용자명"})
            flex = flex[["사용자명", "퇴근기준"]].drop_duplicates()

            df = df.merge(flex, on="사용자명", how="left", suffixes=("", "_유연"))
            df["퇴근기준"] = df["퇴근기준_유연"].fillna(df["퇴근기준"])
            df = df.drop(columns=["퇴근기준_유연"])

    except:
        pass

    # 추가근무 계산
    def calc_overtime(row):
        if pd.isna(row["마지막퇴근"]):
            return 0

        base = pd.to_datetime(f"{row['날짜']} {row['퇴근기준']}", errors="coerce")
        actual = row["마지막퇴근"]

        if pd.isna(base) or pd.isna(actual):
            return 0

        if actual > base:
            return round((actual - base).total_seconds() / 3600, 2)

        return 0

    df["추가근무(시간)"] = df.apply(calc_overtime, axis=1)

    # 보기용 포맷
    df["첫출근"] = df["첫출근"].dt.strftime("%H:%M:%S")
    df["마지막퇴근"] = df["마지막퇴근"].dt.strftime("%H:%M:%S")

    # ---------------------------
    # 1. 전체 직원 요약
    # ---------------------------
    all_summary = df.groupby(["사용자명", "부서명"], as_index=False).agg({
        "날짜": "count",
        "추가근무(시간)": "sum"
    }).rename(columns={
        "날짜": "근무일수",
        "추가근무(시간)": "총 추가근무시간"
    })

    # ---------------------------
    # 2. 부서별 요약
    # ---------------------------
    dept_summary = df.groupby("부서명", as_index=False).agg({
        "사용자명": "nunique",
        "날짜": "count",
        "추가근무(시간)": "sum"
    }).rename(columns={
        "사용자명": "인원수",
        "날짜": "총 근무건수",
        "추가근무(시간)": "부서 총 추가근무시간"
    })

    st.subheader("전체 직원 요약")
    st.dataframe(all_summary)

    st.subheader("부서별 요약")
    st.dataframe(dept_summary)

    st.subheader("상세 결과")
    st.dataframe(df)

    # ---------------------------
    # 엑셀 다운로드
    # ---------------------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        all_summary.to_excel(writer, index=False, sheet_name="전체직원요약")
        dept_summary.to_excel(writer, index=False, sheet_name="부서별요약")
        df.to_excel(writer, index=False, sheet_name="전체상세")

        for name in df["사용자명"].dropna().unique():
            person_df = df[df["사용자명"] == name]
            person_df.to_excel(writer, index=False, sheet_name=safe_sheet_name(name))

    st.download_button(
        "결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name="근태_최종결과.xlsx"
    )
