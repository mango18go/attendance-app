import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time, date
import re

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
        if ("수신날짜" in row_text or "날짜" in row_text) and ("출퇴근" in row_text or "사용자명" in row_text):
            return i
    return None

def find_col(df, keywords):
    for keyword in keywords:
        for col in df.columns:
            if keyword in clean_text(col):
                return col
    return None

def parse_time_value(v):
    if pd.isna(v):
        return None

    if isinstance(v, time):
        return v

    if isinstance(v, datetime):
        return v.time()

    if isinstance(v, pd.Timestamp):
        return v.time()

    if isinstance(v, (int, float)) and not isinstance(v, bool):
        seconds = int(round(float(v) * 24 * 60 * 60)) % 86400
        return time(seconds // 3600, (seconds % 3600) // 60, seconds % 60)

    s = str(v).strip()

    m = re.search(r"(오전|오후)\s*(\d{1,2}):(\d{2})(?::(\d{2}))?", s)
    if m:
        ap = m.group(1)
        h = int(m.group(2))
        mi = int(m.group(3))
        se = int(m.group(4) or 0)

        if ap == "오후" and h < 12:
            h += 12
        if ap == "오전" and h == 12:
            h = 0

        return time(h, mi, se)

    m = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", s)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2))
        se = int(m.group(3) or 0)
        return time(h, mi, se)

    return None

def combine_date_time(d, t):
    d = pd.to_datetime(d, errors="coerce")
    tv = parse_time_value(t)

    if pd.isna(d) or tv is None:
        return pd.NaT

    return datetime.combine(d.date(), tv)

def weekday_kr(d):
    return ["월", "화", "수", "목", "금", "토", "일"][pd.to_datetime(d).weekday()]

def parse_end_time(value):
    if pd.isna(value):
        return "18:00:00"

    text = str(value).replace(" ", "")

    if "~" in text:
        text = text.split("~")[-1]

    tv = parse_time_value(text)

    if tv is None:
        return "18:00:00"

    return tv.strftime("%H:%M:%S")

def parse_period(value, default_year=2026):
    if pd.isna(value):
        return None, None

    s = str(value).replace(" ", "")

    if "~" not in s:
        return None, None

    start_s, end_s = s.split("~")[0], s.split("~")[-1]

    def parse_one(x):
        x = x.replace(".", "-").replace("/", "-")
        nums = re.findall(r"\d+", x)

        if len(nums) >= 3:
            y, m, d = map(int, nums[:3])
        elif len(nums) == 2:
            y = default_year
            m, d = map(int, nums[:2])
        else:
            return None

        return date(y, m, d)

    return parse_one(start_s), parse_one(end_s)

def normalize_days(value):
    if pd.isna(value):
        return ""

    return str(value).replace("요일", "").replace(",", "").replace(" ", "")

def read_flex_rules(file_bytes, xls):
    rules = []

    if "4월 유연근무" not in xls.sheet_names:
        return rules

    raw = pd.read_excel(BytesIO(file_bytes), sheet_name="4월 유연근무", header=None)

    header_row = None
    for i in range(len(raw)):
        row_text = " ".join(raw.iloc[i].fillna("").astype(str).tolist())
        if "성명" in row_text and "유연근무시간" in row_text:
            header_row = i
            break

    if header_row is None:
        return rules

    flex = pd.read_excel(BytesIO(file_bytes), sheet_name="4월 유연근무", header=header_row)
    flex.columns = flex.columns.astype(str).str.strip()

    name_col = find_col(flex, ["성명", "사용자명", "직원명", "이름"])
    time_col = find_col(flex, ["유연근무시간", "근무시간"])
    day_col = find_col(flex, ["요일"])
    period_col = find_col(flex, ["기간"])

    if not name_col or not time_col:
        return rules

    for _, r in flex.iterrows():
        name = str(r[name_col]).strip()

        if not name or name == "nan":
            continue

        start_date, end_date = None, None
        if period_col:
            start_date, end_date = parse_period(r[period_col])

        rules.append({
            "사용자명": name,
            "요일": normalize_days(r[day_col]) if day_col else "",
            "시작일": start_date,
            "종료일": end_date,
            "퇴근기준": parse_end_time(r[time_col])
        })

    return rules

def get_base_time(name, day, weekday, rules):
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

def make_pairs(group):
    group = group.dropna(subset=["일시"]).sort_values("일시")

    pairs = []
    current_in = None

    for _, row in group.iterrows():
        kind = str(row["출퇴근"])
        dt = row["일시"]

        if "출근" in kind:
            if current_in is None:
                current_in = dt

        elif "퇴근" in kind:
            if current_in is not None and dt >= current_in:
                pairs.append((current_in, dt))
                current_in = None

    return pairs

def calc_extra(day, pairs, base_end):
    base = pd.to_datetime(f"{day} {base_end}", errors="coerce")
    weekday = pd.to_datetime(day).weekday()

    total = 0
    starts = []
    ends = []

    # 🔥 토/일이면 전부 추가근무
    if weekday in [5, 6]:
        target_pairs = pairs
    else:
        target_pairs = pairs[1:]

    for start, end in target_pairs:
        if pd.isna(start) or pd.isna(end):
            continue

        if weekday in [5, 6]:
            count_start = start
        else:
            count_start = max(start, base) if pd.notna(base) else start

        if end > count_start:
            hours = (end - count_start).total_seconds() / 3600
            allowed = min(hours, 4 - total)

            if allowed > 0:
                total += allowed
                starts.append(count_start.strftime("%H:%M:%S"))
                ends.append((count_start + pd.Timedelta(hours=allowed)).strftime("%H:%M:%S"))

        if total >= 4:
            break

    return total, ", ".join(starts), ", ".join(ends)
def hours_to_hms(hours):
    if pd.isna(hours):
        return "00:00:00"

    total_seconds = int(round(float(hours) * 3600))

    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60

    return f"{h:02d}:{m:02d}:{s:02d}"

if attendance_file:
    try:
        file_bytes = attendance_file.getvalue()
        xls = pd.ExcelFile(BytesIO(file_bytes))

        target_sheet = "가공" if "가공" in xls.sheet_names else xls.sheet_names[0]

        raw = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=None)
        header_row = find_header_row(raw)

        if header_row is None:
            st.error("헤더 행을 찾지 못했습니다.")
            st.write(raw.head(20))
            st.stop()

        data = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=header_row)
        data.columns = data.columns.astype(str).str.strip()

        date_col = find_col(data, ["수신날짜", "날짜", "일자"])
        time_col = find_col(data, ["24H", "수신시간", "시간"])
        name_col = find_col(data, ["사용자명", "성명", "직원명", "이름"])
        dept_col = find_col(data, ["부서명", "부서", "소속"])
        type_col = find_col(data, ["출퇴근", "출/퇴근", "구분"])

        if not all([date_col, time_col, name_col, type_col]):
            st.error("필수 컬럼을 찾지 못했습니다.")
            st.write("현재 컬럼:", list(data.columns))
            st.stop()

        logs = pd.DataFrame()
        logs["날짜"] = pd.to_datetime(data[date_col], errors="coerce").dt.date
        logs["시간"] = data[time_col]
        logs["사용자명"] = data[name_col].astype(str).str.strip()
        logs["부서명"] = data[dept_col].fillna("").astype(str).str.strip() if dept_col else ""
        logs["출퇴근"] = data[type_col].astype(str).str.strip()

        logs = logs.dropna(subset=["날짜", "사용자명"])
        logs = logs[(logs["사용자명"] != "") & (logs["사용자명"] != "nan")]
        logs["요일"] = logs["날짜"].apply(weekday_kr)
        logs["일시"] = logs.apply(lambda r: combine_date_time(r["날짜"], r["시간"]), axis=1)

        flex_rules = read_flex_rules(file_bytes, xls)

        result_rows = []

        for (day, name, dept), group in logs.groupby(["날짜", "사용자명", "부서명"], dropna=False):
            weekday = weekday_kr(day)
            pairs = make_pairs(group)

            first_in = group[group["출퇴근"].str.contains("출근", na=False)]["일시"].min()
            last_out = group[group["출퇴근"].str.contains("퇴근", na=False)]["일시"].max()

            base_end = get_base_time(name, day, weekday, flex_rules)
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
            st.warning("처리된 데이터가 없습니다.")
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

        # 🔥 여기 추가 (중요)
        final["추가근무(시간)"] = final["추가근무(시간)"].apply(hours_to_hms)
        전체직원요약["총 추가근무시간"] = 전체직원요약["총 추가근무시간"].apply(hours_to_hms)
        부서별요약["부서 총 추가근무시간"] = 부서별요약["부서 총 추가근무시간"].apply(hours_to_hms)

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
