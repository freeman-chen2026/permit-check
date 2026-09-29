# -*- coding: utf-8 -*-
"""
国内批复核对工具 - Streamlit 版本
- 飞行员名单内置
- 航班信息文本框粘贴
- 内置机型对照表
- 标红时只改颜色，绝不动文本
- 内机/外机分别处理
"""

import io
import re
import csv
import copy
import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from openpyxl import load_workbook

# =========================================================
# 常量
# =========================================================
MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
    "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
    "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

PILOT_RAW = """P001,庚凡,gengfan@amber-aviation.com
P002,张永一,zhangyongyi@amber-aviation.com
P003,梅峰,fmei@amber-aviation.com
P004,王斌,wangbin@amber-aviation.com
P019,"HEALY, Darran William",darranhealy@amber-aviation.com
P020,"BEEBE, Thaddeus John",thaddeusbeebe@amber-aviation.com
P032,林毅,ericlin@amber-aviation.com
P035,"Peter Robert, JACKSON",prjackson@amber-aviation.com
P036,王少雄,warrenwang@amber-aviation.com
P038,苗旺旺,johnmiao@amber-aviation.com
P039,"Yiftah, RAUCH",yiftahrauch@amber-aviation.com
P044,李辛欣,rockli@amber-aviation.com
P046,赵岩松,zhyszhao@amber-aviation.com
P051,彭罡,eugene.peng@humbleholding.com
P052,胡君量,brian.wu@humbleholding.com
P053,"Bruce Roderick, WAINES",brwaines@amber-aviation.com
P054,"Rodolfo, BONETTI",rbonetti@amber-aviation.com
P056,"Keith Robert, SHERREN",krsherren@amber-aviation.com
P057,"Oliver Viktor, RACZ",ovracz@amber-aviation.com
P059,蔡国俊,kctsai@amber-aviation.com
P061,李庆宏,qhli@amber-aviation.com
P065,宋炜,wsong@amber-aviation.com
P068,昝昭君,zjzan@amber-aviation.com
P069,"ROEDER, SIMONE ELKE",simoneroeder@amber-aviation.com
P070,"Herve Daniel, STAMM",hdstamm@amber-aviation.com
P071,孙浩,jasonsun@amber-aviation.com
P072,朱正宇,zyzhu@amber-aviation.com
P074,金尚明,smjin@amber-aviation.com
P075,"Eduard Pascal, Roski",eduardroski@amber-aviation.com
P077,刘凯,andyliu@amber-aviation.com
P078,张帆,fzhang@amber-aviation.com
P079,魏思远,wesleywei@amber-aviation.com
P080,刘爽,sliu@amber-aviation.com
P081,吴鹏,richardwu@amber-aviation.com
P082,刘汇川,frankliu@amber-aviation.com
P083,尤欣,xyou@amber-aviation.com
P084,李亚民,ymli@amber-aviation.com
P085,赵镭,lzhao@amber-aviation.com
P086,张贺新,hxzhang@amber-aviation.com
P087,孙赫,hesun@amber-aviation.com
P088,马坚,harryma@amber-aviation.com
P089,李晓龙,xlli@amber-aviation.com
P090,黄海东,hdhuang@amber-aviation.com
P091,马洪双,mikema@amber-aviation.com
PJZ001,张哲,zzhang@amber-aviation.com
PJZ002,郭春旭,charlesguo@amber-aviation.com
PJZ004,王国勤,leowang@amber-aviation.com
PJZ005,王莹,evawang@amber-aviation.com
PJZ007,徐卓,frankxu@amber-aviation.com
PJZ008,杨华,ariayang@amber-aviation.com
W070,王彦海,wang_yanhai@163.com
W213,沈志伟,cshum@tagaviation.com
W267,"Nathon Andrew G, NORBERG",naten7@hotmail.com
W268,"Daniel, RICHTER",pilotlocalizer@gmail.com
W270,杨涛,yang_tao2005@aliyun.com
W272,"Andrew Nigel, KING",Andrew.king@aero.bombardier.com
"""

AIRCRAFT_TYPE_MAP = {
    "B3926": "LJ60",
    "B652R": "GLF4",
    "B8105": "GLEX",
    "B8160": "GLF5",
    "B8262": "GLF4",
    "B8292": "GLF5",
    "B8309": "GLF5",
    "MLLIN": "GLEX",
    "N2QE": "GL5T",
    "N328LM": "GL7T",
    "N550DR": "GLF5",
    "N577QT": "F900",
    "N7777U": "GLEX",
    "N777ZH": "GLF5",
    "N88AY": "GLF5",
    "T7178HT": "GL7T",
    "T7CJK": "GLEX",
    "VPCSZ": "GL7T",
    "VPCVA": "GLF6",
    "B652Q": "GLF4",
    "B652S": "GLF4",
    "B65AP": "GLF4",
}

st.set_page_config(page_title="国内批复核对工具", page_icon="✈️", layout="wide")


# =========================================================
# 工具
# =========================================================
def parse_date_token(token):
    """支持 01OCT2026 和 01OCT26"""
    token = token.strip().upper()
    day = int(token[:2])
    mon = MONTHS[token[2:5]]
    year_str = token[5:]
    year = 2000 + int(year_str) if len(year_str) == 2 else int(year_str)
    return datetime.date(year, mon, day)


def parse_hhmm(token):
    token = str(token).strip().zfill(4)
    return datetime.time(int(token[:2]), int(token[2:]))


def is_b_reg(reg):
    return str(reg).strip().upper().startswith("B")


def to_beijing_datetime(reg, date_obj, hhmm_token):
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
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else s


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


@st.cache_data
def load_pilots():
    pilots = {}
    for line in PILOT_RAW.strip().splitlines():
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
    pilot_codes = [c.strip() for c in crew_codes if c.strip().startswith(("P", "W"))]
    if not pilot_codes:
        return None
    for code in pilot_codes:
        if code not in pilots:
            return None
        if not has_chinese(pilots[code]):
            return False
    return True


# =========================================================
# 解析批复
# =========================================================
APPROVAL_RE = re.compile(
    r"^(?P<reg>[A-Z0-9\-]+)\s+"
    r"(?P<second>[A-Z0-9]+)\s+"
    r"(?P<dep>[A-Z]{4})(?P<dep_time>\d{4})\s+"
    r"(?P<arr_time>\d{4})(?P<arr>[A-Z]{4})\s+"
    r"ON\s+(?P<date>\d{2}[A-Z]{3}\d{2,4})\s+"
    r"(?P<rest>.+)$",
    re.IGNORECASE,
)


def parse_approval_line(text):
    m = APPROVAL_RE.match(text.strip())
    if not m:
        return None

    reg = m.group("reg").upper().replace("-", "")
    second = m.group("second").upper()
    dep = m.group("dep").upper()
    arr = m.group("arr").upper()
    dep_raw = m.group("dep_time")
    arr_raw = m.group("arr_time")
    date_raw = m.group("date").upper()
    rest = m.group("rest").strip()

    # B 注册：第二字段是机型；外机：第二字段是航班号
    if is_b_reg(reg):
        ac_type = second
        flight_no = ""
    else:
        ac_type = ""
        flight_no = second

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
        "type": ac_type,
        "flight_no": flight_no,
        "is_domestic": is_b_reg(reg),
        "dep": dep,
        "dep_time_raw": dep_raw,
        "arr_time_raw": arr_raw,
        "arr": arr,
        "date_raw": date_raw,
        "dep_dt_bj": to_beijing_datetime(reg, date_obj, dep_raw),
        "arr_dt_bj": to_beijing_datetime(reg, date_obj, arr_raw),
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
# Excel
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
# 文本航班信息
# =========================================================
FLIGHT_HEADER_RE = re.compile(
    r"^([A-Z0-9]+)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(?:\s*\+1)?$"
)


def load_text_flights(text: str):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    flights = []
    i = 0
    while i < len(lines):
        m = FLIGHT_HEADER_RE.match(lines[i])
        if m:
            reg = m.group(1).upper()
            dep_time, arr_time = m.group(2), m.group(3)
            if i + 1 < len(lines):
                cm = re.match(r"^(.+?)\s+-\s+(.+)$", lines[i + 1])
                if cm:
                    crew = []
                    if i + 2 < len(lines):
                        cl = lines[i + 2].replace(" ", "")
                        if re.match(r"^[A-Z0-9,]+$", cl):
                            crew = [x for x in cl.split(",") if x]
                    flights.append({
                        "reg": reg,
                        "dep_time": dep_time,
                        "arr_time": arr_time,
                        "dep_city": cm.group(1).strip(),
                        "arr_city": cm.group(2).strip(),
                        "crew": crew,
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
# docx 标红（只改颜色，绝不动文本）
# =========================================================
W_R = qn('w:r')
W_RPR = qn('w:rPr')
W_COLOR = qn('w:color')
W_T = qn('w:t')


def _set_run_red(run_element):
    """给 run 的 rPr 增加红色，保留其他属性"""
    rPr = run_element.find(W_RPR)
    if rPr is None:
        rPr = run_element.makeelement(W_RPR, {})
        run_element.insert(0, rPr)
    for c in rPr.findall(W_COLOR):
        rPr.remove(c)
    color = rPr.makeelement(W_COLOR, {qn('w:val'): 'FF0000'})
    rPr.append(color)


def _make_run_like(src_run_elem, text, red):
    """复制 src_run_elem 的 XML（含字体），只把文本换成 text；red=True 时加红色"""
    new_r = copy.deepcopy(src_run_elem)
    # 删除所有 w:t，保留其它子节点（如 w:rPr）
    for t in new_r.findall(W_T):
        new_r.remove(t)
    t = new_r.makeelement(W_T, {})
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    new_r.append(t)
    if red:
        _set_run_red(new_r)
    return new_r


def set_paragraph_runs(paragraph, text, red_parts):
    """
    仅对需要标红的部分修改颜色：
    - 不含红色的 run 原封不动
    - 整个 run 都是红色的，直接给该 run 加红
    - 部分红色的 run，才拆分；拆出的片段复制原 run 格式
    """
    if not red_parts:
        return

    runs = list(paragraph.runs)
    if not runs:
        return

    full_text = "".join(r.text for r in runs)
    if not full_text:
        return

    # 计算段落内红色字符区间
    ranges = []
    for part in red_parts:
        if not part:
            continue
        start = 0
        while True:
            idx = full_text.find(part, start)
            if idx == -1:
                break
            ranges.append((idx, idx + len(part)))
            start = idx + len(part)

    if not ranges:
        return

    ranges.sort()
    merged = []
    for a, b in ranges:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b))
        else:
            merged.append((a, b))

    def is_red(pos):
        for a, b in merged:
            if a <= pos < b:
                return True
        return False

    pos = 0
    for run in runs:
        r_text = run.text
        if not r_text:
            continue

        r_start = pos
        r_len = len(r_text)

        # 该 run 的红色状态
        red_flags = [is_red(r_start + i) for i in range(r_len)]
        if not any(red_flags):
            pos += r_len
            continue

        if all(red_flags):
            # 整个 run 标红，直接改颜色，不动文本结构
            _set_run_red(run._element)
            pos += r_len
            continue

        # 部分红色：拆分这个 run
        run_elem = run._element
        parent = run_elem.getparent()
        idx_in_parent = list(parent).index(run_elem)

        # 切成片段
        pieces = []
        i = 0
        while i < r_len:
            red_now = red_flags[i]
            j = i + 1
            while j < r_len and red_flags[j] == red_now:
                j += 1
            pieces.append((r_text[i:j], red_now))
            i = j

        # 移除原 run，按顺序插入新片段
        parent.remove(run_elem)
        for k, (seg, red) in enumerate(pieces):
            new_r = _make_run_like(run_elem, seg, red)
            parent.insert(idx_in_parent + k, new_r)

        pos += r_len


# =========================================================
# 主核对
# =========================================================
def run_check(docx_bytes, excel_bytes, text_content, pilots):
    excel_rows = load_excel_rows_from_bytes(excel_bytes)
    text_flights = load_text_flights(text_content)

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

        # ---- 机型：仅对 B 注册核对 ----
        if approval["is_domestic"]:
            expected_type = AIRCRAFT_TYPE_MAP.get(approval["reg"])
            if expected_type is None:
                diffs.append(f"机型：注册号 {approval['reg']} 不在机型对照表中")
                red_parts.append(approval["type"])
            elif approval["type"] != expected_type:
                diffs.append(f"机型：批复 {approval['type']} vs 对照表 {expected_type}")
                red_parts.append(approval["type"])

        # ---- Excel 比对（内机、外机都做） ----
        if excel_row:
            if approval["dep"] != excel_row["dep"]:
                diffs.append(f"起飞机场：批复 {approval['dep']} vs 计划 {excel_row['dep']}")
                red_parts.append(approval["dep"])
            if approval["arr"] != excel_row["arr"]:
                diffs.append(f"到达机场：批复 {approval['arr']} vs 计划 {excel_row['arr']}")
                red_parts.append(approval["arr"])
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
            if excel_row["dep_date"] and approval["dep_dt_bj"].date() != excel_row["dep_date"]:
                diffs.append(
                    f"起飞日期：批复 {approval['dep_dt_bj'].date()} vs 计划 {excel_row['dep_date']}"
                )
                red_parts.append(approval["date_raw"])

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

        # ---- 中国籍：仅对 B 注册核对 ----
        if approval["is_domestic"]:
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
            "航班号": approval["flight_no"] if not approval["is_domestic"] else "",
            "机型": approval["type"] if approval["is_domestic"]
                    else "",
            "内/外机": "内机" if approval["is_domestic"] else "外机",
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

    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if raw_text in approval_red_map:
            set_paragraph_runs(p, raw_text, approval_red_map[raw_text])

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return result_rows, out_buf


# =========================================================
# UI
# =========================================================
st.title("✈️ 国内批复核对工具")
st.caption("上传批复汇总表 + 航段数据，粘贴文本航班信息，即可自动核对差异。")

with st.sidebar:
    st.header("📁 上传文件")
    docx_file = st.file_uploader("① 国内批复信息汇总表 (.docx)", type=["docx"])
    excel_file = st.file_uploader("② 航段数据导出 (.xlsx)", type=["xlsx"])

    st.markdown("---")
    st.markdown(
        "**说明**\n"
        "- 飞行员名单已内置，无需上传\n"
        "- 机型对照表已内置，只对 B 注册核对\n"
        "- **内机**：`注册号 机型 起飞机场+时间 到达时间+到达机场 ON 日期 U/H|N/M`\n"
        "- **外机**：`注册号 航班号 起飞机场+时间 到达时间+到达机场 ON 日期 U/H|N/M`\n"
        "- 外机不核对机型、不核对中国籍机组\n"
        "- B 注册批复时间按北京时间；其他按世界时 UTC +8\n"
        "- `U/H` → 载客，`N/M` → 调机"
    )

text_input = st.text_area(
    "③ 粘贴文本版航班信息",
    height=320,
    placeholder="例如：\nB8160 13:00 - 20:45\n马尔代夫马法鲁岛 - 北京大兴\nP083,PJZ005,C054,M041",
)

text_content = text_input.strip()

if not docx_file or not excel_file or not text_content:
    st.info("👈 请上传批复汇总表、航段数据，并粘贴文本版航班信息。")
    st.stop()

if st.button("🚀 开始核对", type="primary"):
    with st.spinner("正在核对..."):
        try:
            rows, out_buf = run_check(
                docx_file.getvalue(),
                excel_file.getvalue(),
                text_content,
                load_pilots(),
            )
        except Exception as e:
            st.exception(e)
            st.stop()

    df = pd.DataFrame(rows)
    total = len(df)
    diff_count = (df["是否一致"] == "否").sum() if total else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("批复条数", total)
    c2.metric("一致", total - diff_count)
    c3.metric("有差异", diff_count)

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
                for line in r["差异"].split("；"):
                    st.markdown(f"- {line}")

    st.subheader("📥 下载标红后的批复汇总表")
    st.download_button(
        label="下载 _已核对标红.docx",
        data=out_buf.getvalue(),
        file_name=docx_file.name.replace(".docx", "_已核对标红.docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
