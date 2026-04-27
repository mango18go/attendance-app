import streamlit as st
import pandas as pd
from datetime import datetime

st.title("근태 자동판정 프로그램")

file = st.file_uploader("엑셀 업로드", type=["xls", "xlsx"])

if file:
    try:
        raw = pd.read_excel(file, sheet_name="로우", header=None)

        # ✅ 헤더 2행으로 고정
        data = pd.read_excel(file, sheet_name="로우", header=2)

        data.columns = data.columns.astype(str).str.strip()

        # 컬럼 찾기
        time_col = [c for c in data.columns if "수신시간" in c][0]
        dept_col = [c for c in data.columns if "부서명" in c][0]
        id_col = [c for c in data.columns if "사원번호" in c][0]

        df = data[[time_col, dept_col, id_col]].copy()
        df.columns = ["일시", "부서명", "사용자명"]

        # 날짜+시간 변환
        df["일시"] = pd.to_datetime(df["일시"], errors="coerce")
        df = df.dropna(subset=["일시"])

        df["날짜"] = df["일시"].dt.date

        result = []

        # 👉 사람 + 날짜 기준 처리
        for (name, day), group in df.groupby(["사용자명", "날짜"]):
            group = group.sort_values("일시")

            times = list(group["일시"])

            pairs = []
            for i in range(0, len(times)-1, 2):
                pairs.append((times[i], times[i+1]))

            if not pairs:
                continue

            first_in = pairs[0][0]
            last_out = pairs[-1][1]

            base = datetime.combine(day, datetime.strptime("18:00:00", "%H:%M:%S").time())

            extra = 0

            for start, end in pairs[1:]:
                s = max(start, base)
                if end > s:
                    extra += (end - s).total_seconds() / 3600

            extra = min(extra, 4)

            result.append({
                "날짜": day,
                "사용자명": name,
                "부서명": group["부서명"].iloc[0],
                "첫출근": first_in.strftime("%H:%M:%S"),
                "마지막퇴근": last_out.strftime("%H:%M:%S"),
                "출퇴근세트수": len(pairs),
                "추가근무(시간)": round(extra,2)
            })

        final = pd.DataFrame(result)

        st.subheader("전체 결과")
        st.dataframe(final)

    except Exception as e:
        st.error(str(e))
