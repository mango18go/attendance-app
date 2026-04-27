import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

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
        row_text = " ".join(raw.iloc[i].fillna("").astype(str).tolist())
        if ("수신날짜" in row_text or "날짜" in row_text) and ("사용자명" in row_text or "성명" in row_text):
            return i
    return None

def find_col(df, keywords):
    for col in df.columns:
        col_clean = clean_text(col)
        for keyword in keywords:
            if keyword in col_clean:
                return col
    return None

def combine_date_time(d, t):
    d = pd.to_datetime(d, errors="coerce")
    t = pd.to_datetime(t, errors="coerce")
    if pd.isna(d) or pd.isna(t):
        return pd.NaT
    return datetime.combine(d.date(), t.time())

def weekday_kr(d):
    return ["월", "화", "수", "목", "금", "토", "일"][pd.to_datetime(d).weekday()]

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

def calc_extra_work(person_day_df, base_end_time):
    """
    퇴근기준 이후 다시 출근한 시간부터
    이후 퇴근까지 추가근무로 계산.
    하루 최대 4시간.
    """
    person_day_df = person_day_df.sort_values("일시")

    base_time = pd.to_datetime(
        f"{person_day_df['날짜'].iloc[0]} {base_end_time}",
        errors="coerce"
    )

    if pd.isna(base_time):
        return 0, None, None

    extra_start = None
    extra_end = None
    total_hours = 0

    for _, row in person_day_df.iterrows():
        event_type = str(row["출퇴근"])
        event_time = row["일시"]

        if pd.isna(event_time):
            continue

        # 기준 퇴근 이후 다시 출근을 찍은 경우
        if "출근" in event_type and event_time > base_time:
            extra_start = event_time

        # 추가출근 이후 퇴근을 찍은 경우
        elif "퇴근" in event_type and extra_start is not None and event_time > extra_start:
            extra_end = event_time
            total_hours += (extra_end - extra_start).total_seconds() / 3600
            extra_start = None

    total_hours = min(total_hours, 4)

    return round(total_hours, 2), extra_start, extra_end

if attendance_file:
    try:
        xls = pd.ExcelFile(attendance_file)

        target_sheet = "가공" if "가공" in xls.sheet_names else xls.sheet_names[0]

        raw = pd.read_excel(attendance_file, sheet_name=target_sheet, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            st.error("엑셀에서 헤더 행을 찾지 못했습니다.")
            st.write(raw.head(20))
            st.stop()

        data = pd.read_excel(attendance_file, sheet_name=target_sheet, header=header_row)
        data.columns = data.columns.astype(str).str.strip()

        date_col = find_col(data, ["수신날짜", "날짜", "일자"])
        time_col = find_col(data, ["수신시간", "시간"])
        name_col = find_col(data, ["사용자명", "성명", "직원명", "이름"])
        dept_col = find_col(data, ["부서명", "부서", "소속"])
        type_col = find_col(data, ["출퇴근", "구분"])

        missing = []
        if date_col is None:
            missing.append("수신날짜/날짜")
        if time_col is None:
            missing.append("수신시간/시간")
        if name_col is None:
            missing.append("사용자명/성명")
        if type_col is None:
            missing.append("출퇴근")

        if missing:
            st.error(f"필수 컬럼을 찾지 못했습니다: {missing}")
            st.write("현재 컬럼:", list(data.columns))
            st.stop()

        logs = pd.DataFrame()
        logs["날짜"] = pd.to_datetime(data[date_col], errors="coerce").dt.date
        logs["시간"] = data[time_col]
        logs["사용자명"] = data[name_col]
        logs["부서명"] = data[dept_col] if dept_col else ""
        logs["출퇴근"] = data[type_col].astype(str)

        logs = logs.dropna(subset=["날짜", "사용자명"])
        logs["요일"] = logs["날짜"].apply(weekday_kr)
        logs["일시"] = logs.apply(lambda r: combine_date_time(r["날짜"], r["시간"]), axis=1)

        # 기본 퇴근기준
        기준표 = logs[["날짜", "요일", "사용자명", "부서명"]].drop_duplicates()
        기준표["퇴근기준"] = "18:00:00"

        # 유연근무 시트 반영
        if "4월 유연근무" in xls.sheet_names:
            flex = pd.read_excel(attendance_file, sheet_name="4월 유연근무")
            flex.columns = flex.columns.astype(str).str.strip()

            flex_name_col = find_col(flex, ["성명", "사용자명", "직원명", "이름"])
            flex_time_col = find_col(flex, ["유연근무시간", "근무시간"])
            flex_day_col = find_col(flex, ["요일"])

            if flex_name_col and flex_time_col:
                flex_df = pd.DataFrame()
                flex_df["사용자명"] = flex[flex_name_col]
                flex_df["퇴근기준"] = flex[flex_time_col].apply(extract_end_time)

                if flex_day_col:
                    flex_df["요일"] = flex[flex_day_col].astype(str).str.replace("요일", "").str.strip()
                    기준표 = 기준표.merge(
                        flex_df[["사용자명", "요일", "퇴근기준"]],
                        on=["사용자명", "요일"],
                        how="left",
                        suffixes=("", "_유연")
                    )
                else:
                    flex_df = flex_df.drop_duplicates("사용자명")
                    기준표 = 기준표.merge(
                        flex_df[["사용자명", "퇴근기준"]],
                        on="사용자명",
                        how="left",
                        suffixes=("", "_유연")
                    )

                기준표["퇴근기준"] = 기준표["퇴근기준_유연"].fillna(기준표["퇴근기준"])
                기준표 = 기준표.drop(columns=["퇴근기준_유연"])

        result_rows = []

        for _, 기준 in 기준표.iterrows():
            person_logs = logs[
                (logs["날짜"] == 기준["날짜"]) &
                (logs["사용자명"] == 기준["사용자명"])
            ].copy()

            first_in = person_logs[person_logs["출퇴근"].str.contains("출근", na=False)]["일시"].min()
            last_out = person_logs[person_logs["출퇴근"].str.contains("퇴근", na=False)]["일시"].max()

            extra_hours, extra_start, extra_end = calc_extra_work(
                person_logs,
                기준["퇴근기준"]
            )

            result_rows.append({
                "날짜": 기준["날짜"],
                "요일": 기준["요일"],
                "사용자명": 기준["사용자명"],
                "부서명": 기준["부서명"],
                "첫출근": first_in.strftime("%H:%M:%S") if pd.notna(first_in) else "",
                "마지막퇴근": last_out.strftime("%H:%M:%S") if pd.notna(last_out) else "",
                "퇴근기준": 기준["퇴근기준"],
                "추가근무시작": extra_start.strftime("%H:%M:%S") if extra_start is not None else "",
                "추가근무종료": extra_end.strftime("%H:%M:%S") if extra_end is not None else "",
                "추가근무(시간)": extra_hours
            })

        final = pd.DataFrame(result_rows)

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
