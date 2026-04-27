import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
from io import BytesIO

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])
flex_file = st.file_uploader("유연근무 엑셀 업로드 선택사항", type=["xlsx", "xls"])

DEFAULT_END = time(18, 0)
OVERTIME_LIMIT = time(22, 0)

def read_attendance(file):
    xls = pd.ExcelFile(file)
    sheet = "가공" if "가공" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet, header=3)
    df.columns = df.columns.astype(str).str.strip()
    return df, sheet, xls.sheet_names

def read_flex_from_attendance(file):
    xls = pd.ExcelFile(file)
    if "4월 유연근무" not in xls.sheet_names:
        return pd.DataFrame()
    return pd.read_excel(file, sheet_name="4월 유연근무", header=2)

def read_flex(file):
    xls = pd.ExcelFile(file)
    return pd.read_excel(file, sheet_name=xls.sheet_names[0], header=2)

def parse_time_range(value):
    if pd.isna(value):
        return None, None
    text = str(value).replace(" ", "")
    if "~" not in text:
        return None, None
    start, end = text.split("~")
    return pd.to_datetime(start).time(), pd.to_datetime(end).time()

def korean_weekday(dt):
    names = ["월", "화", "수", "목", "금", "토", "일"]
    return names[dt.weekday()]

def is_flex_day(day_text, weekday):
    if pd.isna(day_text):
        return False
    text = str(day_text).replace(" ", "").replace(",", "")
    return weekday in text

def get_standard_end(name, date, flex_df):
    if flex_df.empty:
        return DEFAULT_END

    person_rows = flex_df[flex_df["성명"] == name]

    if person_rows.empty:
        return DEFAULT_END

    weekday = korean_weekday(date)

    for _, row in person_rows.iterrows():
        if is_flex_day(row.get("요일"), weekday):
            _, end_time = parse_time_range(row.get("유연근무시간"))
            if end_time:
                return end_time

    return DEFAULT_END

def combine_datetime(date_value, time_value):
    if pd.isna(date_value) or pd.isna(time_value):
        return pd.NaT

    date_part = pd.to_datetime(date_value).date()

    if isinstance(time_value, time):
        time_part = time_value
    else:
        time_part = pd.to_datetime(str(time_value)).time()

    return datetime.combine(date_part, time_part)

def calc_group_overtime(group, flex_df):
    group = group.sort_values("일시").copy()

    name = group["사용자명"].iloc[0]
    date = group["수신날짜"].iloc[0].date()

    standard_end_time = get_standard_end(name, date, flex_df)
    standard_end_dt = datetime.combine(date, standard_end_time)
    overtime_limit_dt = datetime.combine(date, OVERTIME_LIMIT)

    overtime_minutes = 0
    active_start = None

    for _, row in group.iterrows():
        event_time = row["일시"]
        event_type = str(row["출퇴근"])

        if pd.isna(event_time):
            continue

        if "출근" in event_type and event_time >= standard_end_dt:
            active_start = event_time

        elif "퇴근" in event_type and active_start is not None:
            end_time = min(event_time, overtime_limit_dt)

            if end_time > active_start:
                overtime_minutes += (end_time - active_start).total_seconds() / 60

            active_start = None

    return round(overtime_minutes / 60, 2)

if attendance_file:
    attendance, used_sheet, sheet_names = read_attendance(attendance_file)

    st.write("읽은 시트:", used_sheet)
    st.write("전체 시트:", sheet_names)
    st.write("근태 컬럼:", list(attendance.columns))

    required_cols = ["수신날짜", "24H", "사용자명", "출퇴근"]

    missing = [col for col in required_cols if col not in attendance.columns]

    if missing:
        st.error(f"필수 컬럼이 없습니다: {missing}")
        st.stop()

    if flex_file:
        flex = read_flex(flex_file)
        st.write("유연근무 파일을 따로 사용했습니다.")
    else:
        attendance_file.seek(0)
        flex = read_flex_from_attendance(attendance_file)
        st.write("근태 엑셀 안의 '4월 유연근무' 시트를 사용했습니다.")

    flex.columns = flex.columns.astype(str).str.strip()

    if not flex.empty:
        st.write("유연근무 컬럼:", list(flex.columns))

    attendance["수신날짜"] = pd.to_datetime(attendance["수신날짜"], errors="coerce")
    attendance["일시"] = attendance.apply(
        lambda row: combine_datetime(row["수신날짜"], row["24H"]),
        axis=1
    )

    summary_rows = []

    for (date, name), group in attendance.groupby(["수신날짜", "사용자명"]):
        overtime = calc_group_overtime(group, flex)

        dept = group["부서명"].dropna().iloc[0] if "부서명" in group.columns and not group["부서명"].dropna().empty else ""

        first_in = group[group["출퇴근"].astype(str).str.contains("출근", na=False)]["일시"].min()
        last_out = group[group["출퇴근"].astype(str).str.contains("퇴근", na=False)]["일시"].max()

        standard_end = get_standard_end(name, pd.to_datetime(date), flex)

        summary_rows.append({
            "날짜": pd.to_datetime(date).date(),
            "사용자명": name,
            "부서명": dept,
            "첫출근": first_in.time() if pd.notna(first_in) else "",
            "마지막퇴근": last_out.time() if pd.notna(last_out) else "",
            "퇴근기준": standard_end,
            "추가근무(시간)": overtime
        })

    result = pd.DataFrame(summary_rows)

    st.subheader("추가근무 계산 결과")
    st.dataframe(result)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="추가근무결과")
        attendance.to_excel(writer, index=False, sheet_name="원본가공데이터")
        if not flex.empty:
            flex.to_excel(writer, index=False, sheet_name="유연근무기준")

    st.download_button(
        label="결과 엑셀 다운로드",
        data=output.getvalue(),
        file_name="근태_자동판정_결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
