# -*- coding: utf-8 -*-
"""
国内批复核对工具 - Streamlit 版本
上传：批复汇总表 docx、航段数据导出 xlsx、文本版航班信息 txt、飞行员名单 txt
输出：差异列表 + 标红后的 docx 下载
"""

import io
import re
import csv
import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor
from openpyxl import load_workbook

# =========================================================
# 基础常量
# =========================================================
MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
    "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
    "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

st.set_page_config(
    page_title="国内批复核对工具",
    page_icon="✈️",
    layout="wide",
)


# =========================================================
# 工具函数
# =========================================================
def parse_date_token(token):
    token = token.strip().upper()
    day = int(token[:2])
    mon = MONTHS[token[2:5]]
    year = int(token[5:])
    return datetime.date(year, mon, day)


def parse_hhmm(token):
    token = str(token).strip().zfill(4)
    return datetime.time(int(token[:2]), int(token[2:]))


def is_b_reg(reg):
    return str(reg).strip().upper().startswith("B")


def to_beijing_datetime(reg, date_obj, hhmm_token):
    """
    B 注册：按北京时间直接使用；
    非 B 注册：按世界时 UTC，+8 小时转为北京时间。
    """
    dt = datetime.datetime.combine(date_obj, parse_hhmm(hhmm_token))
    if not is_b_reg(reg):
        dt += datetime.timedelta(hours=8)
    return dt


def fmt_time(value):
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%H:%M")
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M")
    s = str(value).strip()
    m = re.match(r"(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def fmt_date(value):
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def has_chinese(text):
    return bool(re.search(r"[\u4e00-\u9fff]", str(text)))


def time_diff_minutes(t1, t2):
    def to_min(t):
        h, m = t.split(":")
        return int(h) * 60 + int(m)
    try:
        return abs(to_min(t1) - to_min(t2))
    except Exception:
        return 9999


# =========================================================
# 读取飞行员名单
# =========================================================
def load_pilots_from_bytes(data: bytes):
    text = data.decode("utf-8-sig", errors="ignore")
    pilots = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            parts = next(csv.reader([line]))
        except Exception:
            parts = line.split(",")
        if len(parts) >= 2:
            pilots[parts[0].strip()] = parts[1].strip().strip('"')
    return pilots


def crew_all_chinese(crew_codes, pilots):
    pilot_codes = [c.strip() for c in crew_codes
                   if c.strip().startswith(("P", "W"))]
    if not pilot_codes:
        return None
    for code in pilot_codes:
        if code not in pilots:
            return None
        if not has_chinese(pilots[code]):
            return False
    return True


# =========================================================
# 解析批复行
# =========================================================
APPROVAL_RE = re.compile(
    r"^(?P<reg>[A-Z0-9]+)\s+"
    r"(?P<type>[A-Z0-9]+)\s+"
    r"(?P<dep>[A-Z]{4})(?P<dep_time>\d{4})\s+"
    r"(?P<arr_time>\d{4})(?P<arr>[A-Z]{4})\s+"
    r"ON\s+(?P<date>\d{2}[A-Z]{3}\d{4})\s+"
    r"(?P<rest>.+)$",
    re.IGNORECASE,
)


def parse_approval_line(text):
    m = APPROVAL_RE.match(text.strip())
    if not m:
        return None

    reg = m.group("reg").upper()
    dep = m.group("dep").upper()
    arr = m.group("arr").upper()
    dep_time_raw = m.group("dep_time")
    arr_time_raw = m.group("arr_time")
    date_raw = m.group("date").upper()
    rest = m.group("rest").strip()

    service, remark = "", ""
    sm = re.match(r"^(U/H|N/M)(.*)$", rest, re.IGNORECASE)
    if sm:
        service = sm.group(1).upper()
        remark = sm.group(2).strip()
    else:
        parts = rest.split(None, 1)
        service = parts[0].upper() if parts else ""
        remark = parts[1].strip() if len(parts) > 1 else ""

    date_obj = parse_date_token(date_raw)
    return {
        "raw": text.strip(),
        "reg": reg,
        "type": m.group("type").upper(),
        "dep": dep,
        "dep_time_raw": dep_time_raw,
        "arr_time_raw": arr_time_raw,
        "arr": arr,
        "date_raw": date_raw,
        "dep_dt_bj": to_beijing_datetime(reg, date_obj, dep_time_raw),
        "arr_dt_bj": to_beijing_datetime(reg, date_obj, arr_time_raw),
        "service": service,
        "remark": remark,
    }


def iter_doc_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            for np in ncell.paragraphs:
                                yield np


# =========================================================
# 读取 Excel
# =========================================================
def load_excel_rows_from_bytes(data: bytes):
    wb = load_workbook(io.BytesIO(data), data_only=True)
    ws = wb["航段(北京时)"] if "航段(北京时)" in wb.sheetnames else wb.active

    rows = []
    for r in range(3, ws.max_row + 1):
        reg = ws.cell(r, 3).value
        if not reg:
            continue
        rows.append({
            "reg": str(reg).strip().upper(),
            "use": str(ws.cell(r, 4).value or "").strip(),
            "dep_date": fmt_date(ws.cell(r, 7).value),
            "dep_time": fmt_time(ws.cell(r, 8).value),
            "dep": str(ws.cell(r, 11).value or "").strip().upper(),
            "dep_city": str(ws.cell(r, 12).value or "").strip(),
            "arr": str(ws.cell(r, 13).value or "").strip().upper(),
            "arr_city": str(ws.cell(r, 14).value or "").strip(),
            "arr_date": fmt_date(ws.cell(r, 15).value),
            "arr_time": fmt_time(ws.cell(r, 16).value),
        })
    return rows


# =========================================================
# 解析文本版航班信息
# =========================================================
FLIGHT_HEADER_RE = re.compile(
    r"^([A-Z0-9]+)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(?:\s*\+1)?$"
)


def load_text_flights_from_bytes(data: bytes):
    text = data.decode("utf-8-sig", errors="ignore")
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    flights = []
    i = 0
    while i < len(lines):
        m = FLIGHT_HEADER_RE.match(lines[i])
        if m:
            reg = m.group(1).upper()
            dep_time = m.group(2)
            arr_time = m.group(3)

            if i + 1 < len(lines):
                cm = re.match(r"^(.+?)\s+-\s+(.+)$", lines[i + 1])
                if cm:
                    dep_city = cm.group(1).strip()
                    arr_city = cm.group(2).strip()
                    crew = []
                    if i + 2 < len(lines):
                        cl = lines[i + 2].replace(" ", "")
                        if re.match(r"^[A-Z0-9,]+$", cl):
                            crew = [x for x in cl.split(",") if x]

                    flights.append({
                        "reg": reg,
                        "dep_time": dep_time,
                        "arr_time": arr_time,
                        "dep_city": dep_city,
                        "arr_city": arr_city,
                        "crew": crew,
                        "raw": f"{lines[i]}\n{lines[i+1]}\n{','.join(crew)}",
                    })
                    i += 3
                    continue
        i += 1
    return flights


# =========================================================
# 匹配
# =========================================================
def find_excel_match(approval, excel_rows):
    candidates = [
        r for r in excel_rows
        if r["reg"] == approval["reg"]
        and r["dep"] == approval["dep"]
        and r["arr"] == approval["arr"]
        and r["dep_date"] == approval["dep_dt_bj"].date()
    ]
    if not candidates:
        candidates = [
            r for r in excel_rows
            if r["reg"] == approval["reg"]
            and r["dep"] == approval["dep"]
            and r["arr"] == approval["arr"]
        ]
    if not candidates:
        candidates = [r for r in excel_rows if r["reg"] == approval["reg"]]
    if not candidates:
        return None

    target = approval["dep_dt_bj"].strftime("%H:%M")
    candidates.sort(key=lambda r: time_diff_minutes(r["dep_time"], target))
    return candidates[0]


def find_text_match(approval, text_flights, city_to_icao):
    candidates = []
    for tf in text_flights:
        if tf["reg"] != approval["reg"]:
            continue
        if (city_to_icao.get(tf["dep_city"]) == approval["dep"]
                and city_to_icao.get(tf["arr_city"]) == approval["arr"]):
            candidates.append(tf)
    if not candidates:
        candidates = [tf for tf in text_flights if tf["reg"] == approval["reg"]]
    if not candidates:
        return None

    target = approval["dep_dt_bj"].strftime("%H:%M")
    candidates.sort(key=lambda x: time_diff_minutes(x["dep_time"], target))
    return candidates[0]


# =========================================================
# docx 标红
# =========================================================
def set_paragraph_runs(paragraph, text, red_parts):
    p = paragraph._element
    for child in list(p):
        if child.tag.endswith("}r"):
            p.remove(child)

    ranges = []
    for part in red_parts:
        if not part:
            continue
        start = 0
        while True:
            idx = text.find(part, start)
            if idx == -1:
                break
            ranges.append((idx, idx + len(part)))
            start = idx + len(part)

    ranges.sort()
    merged = []
    for a, b in ranges:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    pos = 0
    for a, b in merged:
        if pos < a:
            paragraph.add_run(text[pos:a])
        run = paragraph.add_run(text[a:b])
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        pos = b
    if pos < len(text):
        paragraph.add_run(text[pos:])


# =========================================================
# 核对主流程
# =========================================================
def run_check(docx_bytes, excel_bytes, text_bytes, pilot_bytes):
    pilots = load_pilots_from_bytes(pilot_bytes)
    excel_rows = load_excel_rows_from_bytes(excel_bytes)
    text_flights = load_text_flights_from_bytes(text_bytes)

    city_to_icao = {}
    for row in excel_rows:
        if row["dep_city"]:
            city_to_icao[row["dep_city"]] = row["dep"]
        if row["arr_city"]:
            city_to_icao[row["arr_city"]] = row["arr"]

    doc = Document(io.BytesIO(docx_bytes))

    result_rows = []
    approval_red_map = {}

    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if not raw_text:
            continue

        approval = parse_approval_line(raw_text)
        if not approval:
            continue

        red_parts = []
        diffs = []

        excel_row = find_excel_match(approval, excel_rows)
        text_flight = find_text_match(approval, text_flights, city_to_icao)

        approval_dep_time = approval["dep_dt_bj"].strftime("%H:%M")
        approval_arr_time = approval["arr_dt_bj"].strftime("%H:%M")

        # 机场
        if excel_row:
            if approval["dep"] != excel_row["dep"]:
                diffs.append(f"起飞机场：批复 {approval['dep']} vs 计划 {excel_row['dep']}")
                red_parts.append(approval["dep"])
            if approval["arr"] != excel_row["arr"]:
                diffs.append(f"到达机场：批复 {approval['arr']} vs 计划 {excel_row['arr']}")
                red_parts.append(approval["arr"])

            # 时间
            if approval_dep_time != excel_row["dep_time"]:
                diffs.append(
                    f"起飞时间：批复 {approval_dep_time} vs 计划 {excel_row['dep_time']}"
                )
                red_parts.append(approval["dep_time_raw"])
            if approval_arr_time != excel_row["arr_time"]:
                diffs.append(
                    f"落地时间：批复 {approval_arr_time} vs 计划 {excel_row['arr_time']}"
                )
                red_parts.append(approval["arr_time_raw"])

            # 日期
            if excel_row["dep_date"] and approval["dep_dt_bj"].date() != excel_row["dep_date"]:
                diffs.append(
                    f"起飞日期：批复 {approval['dep_dt_bj'].date()} vs 计划 {excel_row['dep_date']}"
                )
                red_parts.append(approval["date_raw"])

            # 用途
            expected_use = ""
            if approval["service"] == "U/H":
                expected_use = "载客"
            elif approval["service"] == "N/M":
                expected_use = "调机"
            if expected_use and expected_use not in excel_row["use"]:
                diffs.append(
                    f"用途：批复 {approval['service']} -> {expected_use} vs 计划 {excel_row['use']}"
                )
                red_parts.append(approval["service"])
        else:
            diffs.append("未在 Excel 中找到匹配航段")

        # 中国籍备注
        if text_flight:
            all_cn = crew_all_chinese(text_flight["crew"], pilots)
            if all_cn is None:
                diffs.append("中国籍备注：机组名单不全，未核对")
            else:
                actual_cn = "中国籍" in approval["remark"]
                if all_cn and not actual_cn:
                    diffs.append("中国籍备注：机组全为中国籍，但批复未写“中国籍”")
                    red_parts.append(approval["remark"])
                elif not all_cn and actual_cn:
                    diffs.append("中国籍备注：机组含外籍飞行员，但批复写“中国籍”")
                    red_parts.append("中国籍")
        else:
            diffs.append("未找到对应文本版航班信息，无法核验机组")

        result_rows.append({
            "批复": raw_text,
            "飞机号": approval["reg"],
            "批复起飞(北京时)": approval_dep_time,
            "批复落地(北京时)": approval_arr_time,
            "批复日期": approval["dep_dt_bj"].date().isoformat(),
            "Excel 计划起飞": excel_row["dep_time"] if excel_row else "",
            "Excel 计划落地": excel_row["arr_time"] if excel_row else "",
            "Excel 出发地": excel_row["dep"] if excel_row else "",
            "Excel 到达地": excel_row["arr"] if excel_row else "",
            "差异": "；".join(diffs) if diffs else "无",
            "是否一致": "否" if diffs else "是",
        })

        if diffs:
            approval_red_map[raw_text] = red_parts

    # docx 标红
    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if raw_text in approval_red_map:
            set_paragraph_runs(p, raw_text, approval_red_map[raw_text])

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)

    return result_rows, out_buf


# =========================================================
# Streamlit UI
# =========================================================
st.title("✈️ 国内批复核对工具")
st.caption("上传以下四个文件，自动识别批复、比对差异、标红并下载 docx。")

with st.sidebar:
    st.header("📁 上传文件")
    docx_file = st.file_uploader("① 国内批复信息汇总表 (.docx)", type=["docx"])
    excel_file = st.file_uploader("② 航段数据导出 (.xlsx)", type=["xlsx"])
    text_file = st.file_uploader("③ 文本版航班信息 (.txt)", type=["txt"])
    pilot_file = st.file_uploader("④ 飞行员名单 (.txt)", type=["txt"])

    st.markdown("---")
    st.markdown(
        "**说明**\n\n"
        "- B 注册号：批复时间按北京时间\n"
        "- 其他注册号（N、T7、M…）：批复时间按世界时 UTC，+8 转为北京时间\n"
        "- `U/H` 视为载客，`N/M` 视为调机\n"
        "- 备注含“中国籍”则期望机组全为中国籍"
    )

if not (docx_file and excel_file and text_file and pilot_file):
    st.info("👈 请先在左侧上传四个文件。")
    st.stop()

if st.button("🚀 开始核对", type="primary"):
    with st.spinner("正在核对，请稍候..."):
        try:
            rows, out_buf = run_check(
                docx_file.getvalue(),
                excel_file.getvalue(),
                text_file.getvalue(),
                pilot_file.getvalue(),
            )
        except Exception as e:
            st.error(f"核对出错：{e}")
            st.stop()

    df = pd.DataFrame(rows)

    # 汇总
    total = len(df)
    diff_count = (df["是否一致"] == "否").sum() if total else 0
    c1, c2, c3 = st.columns(3)
    c1.metric("批复条数", total)
    c2.metric("一致", total - diff_count)
    c3.metric("有差异", diff_count, delta=None if diff_count == 0 else -diff_count)

    st.subheader("📋 核对结果")
    if total == 0:
        st.warning("未在 docx 中识别到任何批复行。")
    else:
        def highlight(row):
            if row["是否一致"] == "否":
                return ["background-color: #ffe5e5"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(highlight, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("🚨 差异明细")
        diffs_df = df[df["是否一致"] == "否"][["批复", "差异"]]
        if diffs_df.empty:
            st.success("✅ 所有批复与计划一致，未发现差异。")
        else:
            for _, r in diffs_df.iterrows():
                st.markdown(f"**`{r['批复']}`**")
                st.markdown(f"- {r['差异'].replace('；', chr(10) + '- ')}")

    st.subheader("📥 下载标红后的批复汇总表")
    st.download_button(
        label="下载 _已核对标红.docx",
        data=out_buf.getvalue(),
        file_name=docx_file.name.replace(".docx", "_已核对标红.docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
