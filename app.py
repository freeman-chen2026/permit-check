# -*- coding: utf-8 -*-
"""
国内批复核对工具 - Streamlit 版本
- 飞行员名单内置
- 航班信息文本框粘贴
- 内置机型对照表
- 标红时只改颜色，绝不动文本
  · 例外 1：待取消 → 段落末尾追加红色"待取消"
  · 例外 2：待申请 → 按日期插入到该飞机批复队列中，+ 红色"待申请"
- 内机/外机分别处理
- 用途：调机/维修 -> N/M，其余 -> U/H
- 用途三方交叉校验：批复 vs Excel vs 文本 F（漏 F / 多 F 均提示）
- 落地时间 30 分钟容差（仅当起飞时间一致时）
- 起降机场组合在 Excel 中找不到 -> 待取消
- Excel 有计划但批复没有（按日期判重）-> 待申请
- 国籍写错 → 直接标红错误标签
- 批复之间不插入空行
- 下载文件保持原文件名
"""

import io
import re
import csv
import copy
import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement
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

FERRY_KEYWORDS = ("调机", "维修")

st.set_page_config(page_title="国内批复核对工具", page_icon="✈️", layout="wide")


# =========================================================
# 工具
# =========================================================
def parse_date_token(token):
    token = token.strip().upper()
    day = int(token[:2])
    mon = MONTHS[token[2:5]]
    year_str = token[5:]
    year = 2000 + int(year_str) if len(year_str) == 2 else int(year_str)
    return datetime.date(year, mon, day)


def parse_hhmm(token):
    token = str(token).strip().zfill(4)
    return datetime.time(int(token[:2]), int(token[2:]))


def parse_hhmm_str(s):
    if not s:
        return None
    m = re.match(r"^(\d{1,2}):(\d{2})$", str(s).strip())
    if not m:
        return None
    return datetime.time(int(m.group(1)), int(m.group(2)))


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


def is_ferry_use(use_text):
    return any(k in use_text for k in FERRY_KEYWORDS)


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
    pending_f = False

    while i < len(lines):
        line = lines[i]

        if line.upper() == "F":
            pending_f = True
            i += 1
            continue

        if line.upper() in ("TBA",):
            i += 1
            continue

        m = FLIGHT_HEADER_RE.match(line)
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
                        "is_ferry": pending_f,
                    })
                    pending_f = False
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
    ]
    if not candidates:
        return None

    same_date = [r for r in candidates if r["dep_date"] == approval["dep_dt_bj"].date()]
    pool = same_date if same_date else candidates

    target = approval["dep_dt_bj"].strftime("%H:%M")
    pool.sort(key=lambda r: time_diff_minutes(r["dep_time"], target))
    return pool[0]


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
# docx 标红 / 追加
# =========================================================
W_R = qn('w:r')
W_RPR = qn('w:rPr')
W_COLOR = qn('w:color')
W_T = qn('w:t')


def _set_run_red(run_element):
    rPr = run_element.find(W_RPR)
    if rPr is None:
        rPr = run_element.makeelement(W_RPR, {})
        run_element.insert(0, rPr)
    for c in rPr.findall(W_COLOR):
        rPr.remove(c)
    color = rPr.makeelement(W_COLOR, {qn('w:val'): 'FF0000'})
    rPr.append(color)


def _make_run_like(src_run_elem, text, red):
    new_r = copy.deepcopy(src_run_elem)
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
    if not red_parts:
        return
    runs = list(paragraph.runs)
    if not runs:
        return

    full_text = "".join(r.text for r in runs)
    if not full_text:
        return

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
        red_flags = [is_red(r_start + i) for i in range(r_len)]

        if not any(red_flags):
            pos += r_len
            continue

        if all(red_flags):
            _set_run_red(run._element)
            pos += r_len
            continue

        run_elem = run._element
        parent = run_elem.getparent()
        idx_in_parent = list(parent).index(run_elem)

        pieces = []
        i = 0
        while i < r_len:
            red_now = red_flags[i]
            j = i + 1
            while j < r_len and red_flags[j] == red_now:
                j += 1
            pieces.append((r_text[i:j], red_now))
            i = j

        parent.remove(run_elem)
        for k, (seg, red) in enumerate(pieces):
            new_r = _make_run_like(run_elem, seg, red)
            parent.insert(idx_in_parent + k, new_r)

        pos += r_len


def append_red_text(paragraph, text):
    runs = list(paragraph.runs)
    p_elem = paragraph._element
    if runs:
        src = runs[-1]._element
        new_r = _make_run_like(src, text, red=True)
    else:
        new_r = p_elem.makeelement(W_R, {})
        rPr = new_r.makeelement(W_RPR, {})
        new_r.insert(0, rPr)
        color = rPr.makeelement(W_COLOR, {qn('w:val'): 'FF0000'})
        rPr.append(color)
        t = new_r.makeelement(W_T, {})
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        new_r.append(t)
    p_elem.append(new_r)


def find_target_cell(doc, reg):
    reg_norm = reg.upper().replace("-", "")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    ap = parse_approval_line(p.text.strip())
                    if ap and ap["reg"] == reg_norm:
                        return cell
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                text = cell.text.strip().strip("*").strip().replace("-", "").upper()
                if text == reg_norm:
                    if row_idx + 1 < len(table.rows):
                        next_row = table.rows[row_idx + 1]
                        return next_row.cells[0]
                    return cell
    return None


def find_global_template_paragraph(doc):
    for p in iter_doc_paragraphs(doc):
        if parse_approval_line(p.text.strip()):
            return p
    return None


def _make_red_paragraph_element(text, template_p):
    p_elem = OxmlElement('w:p')

    if template_p is not None:
        template_pPr = template_p._element.find(qn('w:pPr'))
        if template_pPr is not None:
            p_elem.append(copy.deepcopy(template_pPr))

    r_elem = OxmlElement('w:r')

    rPr = None
    if template_p is not None:
        for r in template_p.runs:
            rPr_src = r._element.find(W_RPR)
            if rPr_src is not None:
                rPr = copy.deepcopy(rPr_src)
                break

    if rPr is None:
        rPr = OxmlElement('w:rPr')

    for c in rPr.findall(W_COLOR):
        rPr.remove(c)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), 'FF0000')
    rPr.append(color)

    r_elem.append(rPr)

    t_elem = OxmlElement('w:t')
    t_elem.text = text
    t_elem.set(qn('xml:space'), 'preserve')
    r_elem.append(t_elem)

    p_elem.append(r_elem)
    return p_elem


def _reorder_cell_with_pending(cell, pending_items, global_template_p):
    existing_items = []
    template_p = None
    for p in cell.paragraphs:
        raw = p.text.strip()
        ap = parse_approval_line(raw)
        if ap:
            existing_items.append({
                "type": "existing",
                "dt": ap["dep_dt_bj"],
                "element": p._element,
            })
            if template_p is None:
                template_p = p

    if template_p is None:
        template_p = global_template_p

    for item in pending_items:
        existing_items.append({
            "type": "pending",
            "dt": item["dt"],
            "text": item["text"],
        })

    existing_items.sort(key=lambda x: x["dt"])

    tc = cell._tc
    for p_elem in list(tc.findall(qn('w:p'))):
        tc.remove(p_elem)

    for item in existing_items:
        if item["type"] == "existing":
            tc.append(item["element"])
        else:
            new_p = _make_red_paragraph_element(item["text"], template_p)
            tc.append(new_p)


# =========================================================
# 构造"待申请"批复文本
# =========================================================
def build_approval_text(excel_row, is_domestic):
    reg = excel_row["reg"]
    dep_date = excel_row["dep_date"]
    arr_date = excel_row["arr_date"] or dep_date
    dep_time = parse_hhmm_str(excel_row["dep_time"])
    arr_time = parse_hhmm_str(excel_row["arr_time"])

    if not (dep_date and arr_date and dep_time and arr_time):
        return None

    dep_dt = datetime.datetime.combine(dep_date, dep_time)
    arr_dt = datetime.datetime.combine(arr_date, arr_time)

    if is_domestic:
        date_fmt = "%d%b%Y"
        second_field = AIRCRAFT_TYPE_MAP.get(reg, "")
    else:
        dep_dt -= datetime.timedelta(hours=8)
        arr_dt -= datetime.timedelta(hours=8)
        date_fmt = "%d%b%y"
        second_field = reg

    service = "N/M" if is_ferry_use(excel_row["use"]) else "U/H"

    dep_time_str = dep_dt.strftime("%H%M")
    arr_time_str = arr_dt.strftime("%H%M")
    date_str = dep_dt.strftime(date_fmt).upper()

    return (f"{reg} {second_field} "
            f"{excel_row['dep']}{dep_time_str} "
            f"{arr_time_str}{excel_row['arr']} "
            f"ON {date_str} {service}")


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
    cancel_paragraphs = []
    approved_index = {}

    # ===== 第一遍：核对 docx 已有批复 =====
    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if not raw_text:
            continue

        approval = parse_approval_line(raw_text)
        if not approval:
            continue

        approved_index.setdefault(
            (approval["reg"], approval["dep"], approval["arr"]), set()
        ).add(approval["dep_dt_bj"].date())

        red_parts = []
        diffs = []
        cancel_note = ""

        airport_exists = any(
            r["reg"] == approval["reg"]
            and r["dep"] == approval["dep"]
            and r["arr"] == approval["arr"]
            for r in excel_rows
        )

        if not airport_exists:
            cancel_note = "待取消"
            red_parts.append(approval["dep"])
            red_parts.append(approval["arr"])

            result_rows.append({
                "批复": raw_text,
                "飞机号": approval["reg"],
                "航班号": approval["flight_no"] if not approval["is_domestic"] else "",
                "机型": approval["type"] if approval["is_domestic"] else "",
                "内/外机": "内机" if approval["is_domestic"] else "外机",
                "批复起飞(北京时)": approval["dep_dt_bj"].strftime("%H:%M"),
                "批复落地(北京时)": approval["arr_dt_bj"].strftime("%H:%M"),
                "批复日期": approval["dep_dt_bj"].date().isoformat(),
                "Excel 用途": "",
                "文本标记": "",
                "Excel 计划起飞": "",
                "Excel 计划落地": "",
                "Excel 出发地": "",
                "Excel 到达地": "",
                "差异": f"起降机场组合 {approval['dep']}→{approval['arr']} 在 Excel 中不存在",
                "是否一致": "否",
                "备注": cancel_note,
            })
            approval_red_map[raw_text] = red_parts
            cancel_paragraphs.append(raw_text)
            continue

        excel_row = find_excel_match(approval, excel_rows)
        text_flight = find_text_match(approval, text_flights, city_to_icao)

        approval_dep_time = approval["dep_dt_bj"].strftime("%H:%M")
        approval_arr_time = approval["arr_dt_bj"].strftime("%H:%M")

        if approval["is_domestic"]:
            expected_type = AIRCRAFT_TYPE_MAP.get(approval["reg"])
            if expected_type is None:
                diffs.append(f"机型：注册号 {approval['reg']} 不在机型对照表中")
                red_parts.append(approval["type"])
            elif approval["type"] != expected_type:
                diffs.append(f"机型：批复 {approval['type']} vs 对照表 {expected_type}")
                red_parts.append(approval["type"])

        if excel_row:
            dep_same = (approval_dep_time == excel_row["dep_time"])
            if not dep_same:
                diffs.append(
                    f"起飞时间：批复 {approval_dep_time} vs 计划 {excel_row['dep_time']}"
                )
                red_parts.append(approval["dep_time_raw"])

            arr_same = (approval_arr_time == excel_row["arr_time"])
            if not arr_same:
                arr_gap = time_diff_minutes(approval_arr_time, excel_row["arr_time"])
                if not (dep_same and arr_gap <= 30):
                    diffs.append(
                        f"落地时间：批复 {approval_arr_time} vs 计划 {excel_row['arr_time']}"
                    )
                    red_parts.append(approval["arr_time_raw"])

            if excel_row["dep_date"] and approval["dep_dt_bj"].date() != excel_row["dep_date"]:
                diffs.append(
                    f"起飞日期：批复 {approval['dep_dt_bj'].date()} vs 计划 {excel_row['dep_date']}"
                )
                red_parts.append(approval["date_raw"])

            excel_ferry = is_ferry_use(excel_row["use"])
            expected_service = "N/M" if excel_ferry else "U/H"

            if approval["service"] != expected_service:
                diffs.append(
                    f"用途：批复 {approval['service']} vs 计划 {excel_row['use']}"
                    f"（应为 {expected_service}）"
                )
                red_parts.append(approval["service"])

            # ---- 文本 F 标记核对（漏 F / 多 F）----
            if text_flight is not None:
                text_ferry = text_flight.get("is_ferry", False)
                if excel_ferry and not text_ferry:
                    diffs.append(
                        f"⚠ 文本漏 F 标记（Excel 为调机：{excel_row['use']}）"
                    )
                    if approval["service"] not in red_parts:
                        red_parts.append(approval["service"])
                elif not excel_ferry and text_ferry:
                    diffs.append(
                        f"⚠ 文本多标 F 标记（Excel 为 {excel_row['use']}，非调机）"
                    )
                    if approval["service"] not in red_parts:
                        red_parts.append(approval["service"])

        if approval["is_domestic"]:
            if text_flight:
                all_cn = crew_all_chinese(text_flight["crew"], pilots)
                if all_cn is None:
                    diffs.append("国籍标注：机组名单不全，无法核对")
                else:
                    has_cn = "中国籍" in approval["remark"]
                    has_foreign = "外籍" in approval["remark"]
                    if all_cn:
                        if has_foreign:
                            diffs.append("国籍标注错误（应为中国籍）")
                            red_parts.append("外籍")
                        elif not has_cn:
                            diffs.append("国籍未标注（应为中国籍）")
                    else:
                        if has_cn:
                            diffs.append("国籍标注错误（应为外籍）")
                            red_parts.append("中国籍")
                        elif not has_foreign:
                            diffs.append("国籍未标注（应为外籍）")
            else:
                diffs.append("未找到对应文本版航班信息，无法核验机组")

        text_ferry_label = ""
        if text_flight is not None:
            text_ferry_label = "调机(F)" if text_flight.get("is_ferry", False) else "载客(无F)"

        result_rows.append({
            "批复": raw_text,
            "飞机号": approval["reg"],
            "航班号": approval["flight_no"] if not approval["is_domestic"] else "",
            "机型": approval["type"] if approval["is_domestic"] else "",
            "内/外机": "内机" if approval["is_domestic"] else "外机",
            "批复起飞(北京时)": approval_dep_time,
            "批复落地(北京时)": approval_arr_time,
            "批复日期": approval["dep_dt_bj"].date().isoformat(),
            "Excel 用途": excel_row["use"] if excel_row else "",
            "文本标记": text_ferry_label,
            "Excel 计划起飞": excel_row["dep_time"] if excel_row else "",
            "Excel 计划落地": excel_row["arr_time"] if excel_row else "",
            "Excel 出发地": excel_row["dep"] if excel_row else "",
            "Excel 到达地": excel_row["arr"] if excel_row else "",
            "差异": "；".join(diffs) if diffs else "无",
            "是否一致": "否" if diffs else "是",
            "备注": cancel_note,
        })

        if diffs:
            approval_red_map[raw_text] = red_parts

    # 标红
    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if raw_text in approval_red_map:
            set_paragraph_runs(p, raw_text, approval_red_map[raw_text])

    # 追加"待取消"
    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if raw_text in cancel_paragraphs:
            append_red_text(p, "  待取消")

    # ===== 第二遍：待申请 =====
    pending_by_reg = {}
    for row in excel_rows:
        if not row["reg"]:
            continue
        if not (row["dep"].startswith("Z") or row["arr"].startswith("Z")):
            continue
        if row["dep_date"] is None:
            continue
        key = (row["reg"], row["dep"], row["arr"])
        if row["dep_date"] in approved_index.get(key, set()):
            continue
        pending_by_reg.setdefault(row["reg"], []).append(row)

    pending_rows = []
    global_template_p = find_global_template_paragraph(doc)

    for reg, rows in pending_by_reg.items():
        target_cell = find_target_cell(doc, reg)
        if target_cell is None:
            continue

        is_domestic = is_b_reg(reg)
        items = []
        for row in rows:
            text = build_approval_text(row, is_domestic)
            if not text:
                continue
            dep_t = parse_hhmm_str(row["dep_time"])
            if not row["dep_date"] or not dep_t:
                continue
            dep_dt_bj = datetime.datetime.combine(row["dep_date"], dep_t)

            note_kind = ""

            if is_domestic:
                fake_ap = {
                    "reg": reg,
                    "dep": row["dep"],
                    "arr": row["arr"],
                    "dep_dt_bj": dep_dt_bj,
                }
                tf = find_text_match(fake_ap, text_flights, city_to_icao)

                if tf is None or not tf["crew"]:
                    note_kind = "机组未定"
                else:
                    all_cn = crew_all_chinese(tf["crew"], pilots)
                    if all_cn is True:
                        note_kind = "中国籍"
                    elif all_cn is False:
                        note_kind = "外籍"
                    else:
                        note_kind = "机组未定"

            if note_kind == "机组未定":
                full_text = text + "  待申请（机组未定，需确认）"
            elif note_kind == "中国籍":
                full_text = text + " 中国籍  待申请"
            elif note_kind == "外籍":
                full_text = text + " 外籍  待申请"
            else:
                full_text = text + "  待申请"

            items.append({
                "dt": dep_dt_bj,
                "text": full_text,
                "row": row,
                "note_kind": note_kind,
            })

        if not items:
            continue

        _reorder_cell_with_pending(target_cell, items, global_template_p)

        for item in items:
            row = item["row"]
            if is_domestic:
                bj_dep = row["dep_time"]
                bj_arr = row["arr_time"]
            else:
                dep_t2 = parse_hhmm_str(row["dep_time"])
                arr_t2 = parse_hhmm_str(row["arr_time"])
                if row["dep_date"] and dep_t2:
                    bj_dep = (datetime.datetime.combine(row["dep_date"], dep_t2)
                              + datetime.timedelta(hours=8)).strftime("%H:%M")
                else:
                    bj_dep = ""
                if row["arr_date"] and arr_t2:
                    bj_arr = (datetime.datetime.combine(row["arr_date"], arr_t2)
                              + datetime.timedelta(hours=8)).strftime("%H:%M")
                else:
                    bj_arr = ""

            if item["note_kind"] == "机组未定":
                remark = "待申请（机组未定，需确认）"
                diff_text = "Excel 有计划，批复汇总表缺失；机组信息未定，需确认中外籍"
            else:
                remark = "待申请"
                diff_text = "Excel 有计划，批复汇总表缺失"

            pending_rows.append({
                "批复": item["text"],
                "飞机号": reg,
                "航班号": reg if not is_domestic else "",
                "机型": AIRCRAFT_TYPE_MAP.get(reg, "") if is_domestic else "",
                "内/外机": "内机" if is_domestic else "外机",
                "批复起飞(北京时)": bj_dep,
                "批复落地(北京时)": bj_arr,
                "批复日期": row["dep_date"].isoformat() if row["dep_date"] else "",
                "Excel 用途": row["use"],
                "文本标记": "",
                "Excel 计划起飞": row["dep_time"],
                "Excel 计划落地": row["arr_time"],
                "Excel 出发地": row["dep"],
                "Excel 到达地": row["arr"],
                "差异": diff_text,
                "是否一致": "否",
                "备注": remark,
            })

    result_rows.extend(pending_rows)

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
        "- 飞行员名单、机型对照表已内置\n"
        "- **用途规则**：Excel 为 `调机` / `维修` → `N/M`；其余 → `U/H`\n"
        "- **用途三方交叉校验**：批复 vs Excel vs 文本 F，漏 F / 多 F 均提示\n"
        "- **落地时间容差**：起飞时间一致时，落地时间差 ≤ 30 分钟视为一致\n"
        "- **待取消**：批复的起降机场组合在 Excel 找不到 → 段落末尾追加红色“待取消”\n"
        "- **待申请**：Excel 有、批复无该航段日期（含 Z 机场）→ 按日期插入到该飞机批复中\n"
        "- **国籍核对**：批复与文本机组国籍不符 → 标红批复里的标签\n"
        "- B 注册批复时间按北京时间；其他按世界时 UTC +8\n"
        "- 文本里单独的 `F` 属于**下一段**航班"
    )

text_input = st.text_area(
    "③ 粘贴文本版航班信息",
    height=320,
    placeholder=(
        "例如：\n"
        "B65AP 16:30 - 17:45\n"
        "香港 - 泉州晋江\n"
        "P057,P039,C046,M035\n"
        "\n"
        "F\n"
        "B652Q 19:15 - 21:30\n"
        "北京大兴 - 宁波栎社\n"
        "M021"
    ),
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
                if "机组未定" in str(row["备注"]):
                    return ["background-color: #fff3cd"] * len(row)
                return ["background-color: #ffe5e5"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(highlight, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("🚨 差异明细")
        diffs_df = df[df["是否一致"] == "否"][["批复", "差异", "备注"]]
        if diffs_df.empty:
            st.success("✅ 所有批复与计划一致，未发现差异。")
        else:
            for _, r in diffs_df.iterrows():
                if "机组未定" in str(r["备注"]):
                    note_html = (
                        " <span style='color:#d97706;font-weight:bold'>"
                        f"【{r['备注']}】</span>"
                    )
                elif r["备注"]:
                    note_html = (
                        f" <span style='color:red;font-weight:bold'>【{r['备注']}】</span>"
                    )
                else:
                    note_html = ""
                st.markdown(f"**`{r['批复']}`**{note_html}", unsafe_allow_html=True)
                for line in r["差异"].split("；"):
                    st.markdown(f"- {line}")

    st.subheader("📥 下载标红后的批复汇总表")
    st.download_button(
        label=f"下载 {docx_file.name}",
        data=out_buf.getvalue(),
        file_name=docx_file.name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
