import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date, time

st.title("근태 자동판정 프로그램")

attendance_file = st.file_uploader("근태 엑셀 업로드", type=["xlsx", "xls"])

def clean_text(x):
    return str(x).strip().replace("\n", "").replace(" ", "")

def safe_sheet_name(name):
    invalid = ["\\", "/", "*", "[", "]", ":", "?"]
    name = str(name)
    for ch in invalid:
        name = name.replace(ch, "")
    return name[:31] if name else "직원"

def find_header_row(raw):
    for i in range(len(raw)):
        row_text = " ".join(raw.iloc[i].astype(str).tolist())
        if ("날짜" in row_text or "일자" in row_text) and ("출근" in row_text or "퇴근" in row_text):
            return i
    return None

def find_col(df, keywords):
    for col in df.columns:
        c = clean_text(col)
        for k in keywords:
            if k in c:
                return col
    return None

def combine_date_time(d, t):
    if pd.isna(d) or pd.isna(t):
        return pd.NaT

    d = pd.to_datetime(d, errors="coerce")
    t = pd.to_datetime(t, errors="coerce")

    if pd.isna(d) or pd.isna(t):
        return pd.NaT

    return datetime.combine(d.date(), t.time())

def extract_end_time(value):
    if pd.isna(value):
        return "18:00:00"

    text = str(value).replace(" ", "")

    if "~" in text:
        text = text.split("~")[-1]

    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return "18:00:00"

    return parsed.strftime("%H:%M:%S")

if attendance_file:
    try:
        xls = pd.ExcelFile(attendance_file)

        target_sheet = "가공" if "가공" in xls.sheet_names else xls.sheet_names[0]

        raw = pd.read_excel(attendance_file, sheet_name=target_sheet, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            st.error("엑셀에서 날짜/출근/퇴근 컬럼이 있는 행을 찾지 못했습니다.")
            st.write("시트명:", target_sheet)
            st.write(raw.head(15))
            st.stop()

        attendance = pd.read_excel(attendance_file, sheet_name=target_sheet, header=header_row)
        attendance.columns = attendance.columns.astype(str).str.strip()

        date_col = find_col(attendance, ["날짜", "일자"])
        name_col = find_col(attendance, ["사용자명", "성명", "직원명", "이름"])
        dept_col = find_col(attendance, ["부서명", "부서", "소속"])
        in_col = find_col(attendance, ["첫출근", "출근시간", "출근"])
        out_col = find_col(attendance, ["마지막퇴근", "최종퇴근", "퇴근시간", "퇴근"])

        missing = []
        if date_col is None:
            missing.append("날짜")
        if name_col is None:
            missing.append("사용자명/성명")
        if in_col is None:
            missing.append("첫출근/출근")
        if out_col is None:
            missing.append("마지막퇴근/퇴근")

        if missing:
            st.error(f"필수 컬럼을 찾지 못했습니다: {missing}")
            st.write("현재 컬럼:", list(attendance.columns))
            st.stop()

        df = pd.DataFrame()
        df["날짜"] = attendance[date_col]
        df["사용자명"] = attendance[name_col]
        df["부서명"] = attendance[dept_col] if dept_col else ""
        df["첫출근"] = attendance[in_col]
        df["마지막퇴근"] = attendance[out_col]

        df = df.dropna(subset=["날짜", "사용자명"])

        df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce").dt.date
        df["첫출근_dt"] = df.apply(lambda r: combine_date_time(r["날짜"], r["첫출근"]), axis=1)
        df["마지막퇴근_dt"] = df.apply(lambda r: combine_date_time(r["날짜"], r["마지막퇴근"]), axis=1)

        df = df.groupby(["날짜", "사용자명", "부서명"], as_index=False).agg({
            "첫출근_dt": "min",
            "마지막퇴근_dt": "max"
        })

        df["퇴근기준"] = "18:00:00"

        if "4월 유연근무" in xls.sheet_names:
            flex = pd.read_excel(attendance_file, sheet_name="4월 유연근무")
            flex.columns = flex.columns.astype(str).str.strip()

            flex_name_col = find_col(flex, ["성명", "사용자명", "직원명", "이름"])
            flex_time_col = find_col(flex, ["유연근무시간", "근무시간"])

            if flex_name_col and flex_time_col:
                flex_df = pd.DataFrame()
                flex_df["사용자명"] = flex[flex_name_col]
                flex_df["퇴근기준"] = flex[flex_time_col].apply(extract_end_time)
                flex_df = flex_df.dropna(subset=["사용자명"]).drop_duplicates("사용자명")

                df = df.merge(flex_df, on="사용자명", how="left", suffixes=("", "_유연"))
                df["퇴근기준"] = df["퇴근기준_유연"].fillna(df["퇴근기준"])
                df = df.drop(columns=["퇴근기준_유연"])

        def calc_overtime(row):
            if pd.isna(row["마지막퇴근_dt"]):
                return 0

            base = pd.to_datetime(f"{row['날짜']} {row['퇴근기준']}", errors="coerce")

            if pd.isna(base):
                return 0

            if row["마지막퇴근_dt"] > base:
                return round((row["마지막퇴근_dt"] - base).total_seconds() / 3600, 2)

            return 0

        df["추가근무(시간)"] = df.apply(calc_overtime, axis=1)

        df["첫출근"] = df["첫출근_dt"].dt.strftime("%H:%M:%S")
        df["마지막퇴근"] = df["마지막퇴근_dt"].dt.strftime("%H:%M:%S")

        final = df[["날짜", "사용자명", "부서명", "첫출근", "마지막퇴근", "퇴근기준", "추가근무(시간)"]]

        전체직원요약 = final.groupby(["사용자명", "부서명"], as_index=False).agg({
            "날짜": "count",
            "추가근무(시간)": "sum"
        }).rename(columns={
            "날짜": "근무일수",
            "추가근무(시간)": "총 추가근무시간"
        })

        부서별요약 = final.groupby("부서명", as_index=False).agg({
            "사용자명": "nunique",
            "날짜": "count",
            "추가근무(시간)": "sum"
        }).rename(columns={
            "사용자명": "인원수",
            "날짜": "총 근무건수",
            "추가근무(시간)": "부서 총 추가근무시간"
        })

        st.subheader("전체 직원 요약")
        st.dataframe(전체직원요약)

        st.subheader("부서별 요약")
        st.dataframe(부서별요약)

        st.subheader("전체 상세")
        st.dataframe(final)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            전체직원요약.to_excel(writer, index=False, sheet_name="전체직원요약")
            부서별요약.to_excel(writer, index=False, sheet_name="부서별요약")
            final.to_excel(writer, index=False, sheet_name="전체상세")

            used_names = set()
            for name in sorted(final["사용자명"].dropna().unique()):
                sheet_name = safe_sheet_name(name)
                original = sheet_name
                n = 1
                while sheet_name in used_names:
                    sheet_name = safe_sheet_name(f"{original}_{n}")
                    n += 1
                used_names.add(sheet_name)

                person_df = final[final["사용자명"] == name]
                person_df.to_excel(writer, index=False, sheet_name=sheet_name)

        st.download_button(
            label="결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name="근태_최종결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("처리 중 오류가 발생했습니다.")
        st.write("오류 내용:", str(e))
        st.write("엑셀 구조나 컬럼명이 예상과 다를 수 있습니다.")
