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
        if ("수신날짜" in row_text or "날짜" in row_text) and ("수신시간" in row_text or "시간" in row_text):
            return i
    return None

def find_col(df, keywords):
    for keyword in keywords:
        for col in df.columns:
            if keyword in clean_text(col):
                return col
    return None

def parse_datetime(d, t):
    d = pd.to_datetime(d, errors="coerce")
    if pd.isna(d) or pd.isna(t):
        return pd.NaT

    if isinstance(t, (int, float)) and not isinstance(t, bool):
        seconds = int(round(float(t) * 24 * 60 * 60)) % (24 * 60 * 60)
        t_str = f"{seconds//3600:02d}:{(seconds%3600)//60:02d}:{seconds%60:02d}"
    else:
        t_str = str(t).strip()

    parsed_t = pd.to_datetime(t_str, errors="coerce")
    if pd.isna(parsed_t):
        return pd.NaT

    return datetime.combine(d.date(), parsed_t.time())

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

def parse_period(value):
    if pd.isna(value):
        return None, None

    text = str(value).replace(" ", "")
    if "~" in text:
        a, b = text.split("~")[0], text.split("~")[-1]
    elif "-" in text and len(text) > 12:
        return None, None
    else:
        return None, None

    start = pd.to_datetime(a, errors="coerce")
    end = pd.to_datetime(b, errors="coerce")

    if pd.isna(start) or pd.isna(end):
        return None, None

    return start.date(), end.date()

def make_pairs(person_logs):
    person_logs = person_logs.sort_values("일시")
    pairs = []
    current_in = None

    for _, row in person_logs.iterrows():
        kind = str(row["출퇴근"])
        dt = row["일시"]

        if pd.isna(dt):
            continue

        if "출근" in kind:
            if current_in is None:
                current_in = dt

        elif "퇴근" in kind:
            if current_in is not None and dt >= current_in:
                pairs.append((current_in, dt))
                current_in = None

    return pairs

def get_flex_base(name, day, weekday, rules):
    for rule in rules:
        if rule["사용자명"] != name:
            continue

        if rule["요일"] and weekday not in rule["요일"]:
            continue

        if rule["시작일"] and day < rule["시작일"]:
            continue

        if rule["종료일"] and day > rule["종료일"]:
            continue

        return rule["퇴근기준"]

    return "18:00:00"

def calc_extra(day, pairs, base_end):
    base = pd.to_datetime(f"{day} {base_end}", errors="coerce")

    total = 0
    starts = []
    ends = []

    # 첫 번째 출근-퇴근 세트는 정상근무
    # 두 번째 세트부터 추가근무 후보
    for start, end in pairs[1:]:
        if pd.isna(base) or pd.isna(start) or pd.isna(end):
            continue

        count_start = max(start, base)

        if end > count_start:
            total += (end - count_start).total_seconds() / 3600
            starts.append(count_start.strftime("%H:%M:%S"))
            ends.append(end.strftime("%H:%M:%S"))

    total = min(total, 4)

    return round(total, 2), ", ".join(starts), ", ".join(ends)

if attendance_file:
    try:
        xls = pd.ExcelFile(attendance_file)
        target_sheet = "가공" if "가공" in xls.sheet_names else xls.sheet_names[0]

        raw = pd.read_excel(attendance_file, sheet_name=target_sheet, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            st.error("헤더 행을 찾지 못했습니다.")
            st.write(raw.head(20))
            st.stop()

        data = pd.read_excel(attendance_file, sheet_name=target_sheet, header=header_row)
        data.columns = data.columns.astype(str).str.strip()

        date_col = find_col(data, ["수신날짜", "날짜", "일자"])
        time_col = find_col(data, ["수신시간", "시간"])
        name_col = find_col(data, ["사용자명", "성명", "직원명", "이름"])
        dept_col = find_col(data, ["부서명", "부서", "소속"])
        type_col = find_col(data, ["출퇴근", "출/퇴근"])
        if type_col is None:
            type_col = find_col(data, ["구분", "신호"])

        if not all([date_col, time_col, name_col, type_col]):
            st.error("필수 컬럼을 찾지 못했습니다.")
            st.write("현재 컬럼:", list(data.columns))
            st.stop()

        logs = pd.DataFrame()
        logs["날짜"] = pd.to_datetime(data[date_col], errors="coerce").dt.date
        logs["시간"] = data[time_col]
        logs["사용자명"] = data[name_col].astype(str).str.strip()
        logs["부서명"] = data[dept_col] if dept_col else ""
        logs["출퇴근"] = data[type_col].astype(str)

        logs = logs.dropna(subset=["날짜", "사용자명"])
        logs["요일"] = logs["날짜"].apply(weekday_kr)
        logs["일시"] = logs.apply(lambda r: parse_datetime(r["날짜"], r["시간"]), axis=1)

        flex_rules = []

        if "4월 유연근무" in xls.sheet_names:
            flex = pd.read_excel(attendance_file, sheet_name="4월 유연근무")
            flex.columns = flex.columns.astype(str).str.strip()

            flex_name_col = find_col(flex, ["성명", "사용자명", "직원명", "이름"])
            flex_time_col = find_col(flex, ["유연근무시간", "근무시간"])
            flex_day_col = find_col(flex, ["요일"])
            flex_period_col = find_col(flex, ["기간"])

            if flex_name_col and flex_time_col:
                for _, r in flex.iterrows():
                    name = str(r[flex_name_col]).strip()
                    end_time = extract_end_time(r[flex_time_col])

                    days = ""
                    if flex_day_col:
                        days = str(r[flex_day_col]).replace("요일", "").replace(" ", "")

                    start_date, end_date = None, None
                    if flex_period_col:
                        start_date, end_date = parse_period(r[flex_period_col])

                    flex_rules.append({
                        "사용자명": name,
                        "요일": days,
                        "시작일": start_date,
                        "종료일": end_date,
                        "퇴근기준": end_time
                    })

        result_rows = []

        for (day, name, dept), group in logs.groupby(["날짜", "사용자명", "부서명"]):
            weekday = weekday_kr(day)
            pairs = make_pairs(group)
            base_end = get_flex_base(name, day, weekday, flex_rules)

            first_in = pairs[0][0] if pairs else pd.NaT
            last_out = pairs[-1][1] if pairs else pd.NaT

            extra_hours, extra_start, extra_end = calc_extra(day, pairs, base_end)

            result_rows.append({
                "날짜": day,
                "요일": weekday,
                "사용자명": name,
                "부서명": dept,
                "첫출근": first_in.strftime("%H:%M:%S") if pd.notna(first_in) else "",
                "마지막퇴근": last_out.strftime("%H:%M:%S") if pd.notna(last_out) else "",
                "출퇴근세트수": len(pairs),
                "퇴근기준": base_end,
                "추가근무시작": extra_start,
                "추가근무종료": extra_end,
                "추가근무(시간)": extra_hours
            })

        final = pd.DataFrame(result_rows)

        if final.empty:
            st.warning("처리된 근태 데이터가 없습니다.")
            st.stop()

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
