# -*- coding: utf-8 -*-
"""
国内批复核对工具 - Streamlit 版本

主要特性：
- 软换行（Shift+Enter）自动拆分成独立段落
- 外机两种格式识别（已批 / 未批）
- 用途映射：FERRY → N/M；BUSINESS → U/H
- 待申请简洁格式：`VPCSZ ZGSZ-ZBAD 07OCT 待申请`
- 追加文字（待取消/待变更/待申请）红字 + 黄底
- 三态判定：是 / 否 / 待确认
- 文本覆盖率检查：粘贴后立即执行，缺失航段可展开复制
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

RED = "FF0000"
GREEN = "00B050"
HIGHLIGHT_YELLOW = "yellow"

MAX_CROSS_DAY_GAP_MIN = 600
EARLY_GREEN_THRESHOLD_MIN = 600

COVERAGE_ERROR_THRESHOLD = 0.60
COVERAGE_WARN_THRESHOLD = 0.85

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
    "B3926": "LJ60", "B652R": "GLF4", "B8105": "GLEX", "B8160": "GLF5",
    "B8262": "GLF4", "B8292": "GLF5", "B8309": "GLF5", "MLLIN": "GLEX",
    "N2QE": "GL5T", "N328LM": "GL7T", "N550DR": "GLF5", "N577QT": "F900",
    "N7777U": "GLEX", "N777ZH": "GLF5", "N88AY": "GLF5", "T7178HT": "GL7T",
    "T7CJK": "GLEX", "VPCSZ": "GL7T", "VPCVA": "GLF6", "B652Q": "GLF4",
    "B652S": "GLF4", "B65AP": "GLF4",
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


def hhmm_to_minutes(t):
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def flight_duration_minutes(dep_t, arr_t):
    try:
        d = hhmm_to_minutes(dep_t)
        a = hhmm_to_minutes(arr_t)
    except Exception:
        return None
    if a < d:
        a += 1440
    return a - d


def fmt_duration(mins):
    if mins is None:
        return ""
    sign = "-" if mins < 0 else ""
    m = abs(mins)
    return f"{sign}{m // 60}:{m % 60:02d}"


def compute_real_minute_diff(ap_date, ap_time_str, xl_date, xl_time_str):
    if ap_date is None or xl_date is None:
        return None
    try:
        ap_min = hhmm_to_minutes(ap_time_str)
        xl_min = hhmm_to_minutes(xl_time_str)
    except Exception:
        return None
    ap_abs = ap_date.toordinal() * 1440 + ap_min
    xl_abs = xl_date.toordinal() * 1440 + xl_min
    return ap_abs - xl_abs


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
        if second == reg:
            flight_no = second
            ac_type = ""
        else:
            ac_type = second
            flight_no = ""

    service, remark = "", ""
    sm = re.match(r"^(U/H|N/M)\s*(.*)$", rest, re.IGNORECASE)
    if sm:
        service = sm.group(1).upper()
        remark = sm.group(2).strip()
    else:
        upper = rest.upper()
        if upper.startswith("FERRY"):
            service = "N/M"
            remark = rest[5:].strip(" -–—\t")
        elif upper.startswith("BUSINESS"):
            service = "U/H"
            remark = rest[8:].strip(" -–—\t")
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
# 软换行（<w:br/>）拆分成独立段落
# =========================================================
def _collect_all_paragraphs(doc):
    result = []
    for p in doc.paragraphs:
        result.append(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    result.append(p)
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            for np in ncell.paragraphs:
                                result.append(np)
    return result


def _split_paragraph_by_br(p_elem):
    parent = p_elem.getparent()
    if parent is None:
        return
    idx_in_parent = list(parent).index(p_elem)
    pPr = p_elem.find(qn('w:pPr'))

    groups = [[]]
    for child in list(p_elem):
        if child.tag == qn('w:pPr'):
            continue
        if child.tag == qn('w:r'):
            brs = child.findall(qn('w:br'))
            if not brs:
                groups[-1].append(copy.deepcopy(child))
            else:
                current_r_children = []
                for rc in list(child):
                    if rc.tag == qn('w:br'):
                        if current_r_children:
                            new_r = OxmlElement('w:r')
                            for x in current_r_children:
                                new_r.append(copy.deepcopy(x))
                            groups[-1].append(new_r)
                            current_r_children = []
                        groups.append([])
                    else:
                        current_r_children.append(rc)
                if current_r_children:
                    new_r = OxmlElement('w:r')
                    for x in current_r_children:
                        new_r.append(copy.deepcopy(x))
                    groups[-1].append(new_r)
        else:
            groups[-1].append(copy.deepcopy(child))

    if len(groups) <= 1:
        return

    parent.remove(p_elem)
    for i, group in enumerate(groups):
        new_p = OxmlElement('w:p')
        if pPr is not None:
            new_p.append(copy.deepcopy(pPr))
        for child in group:
            new_p.append(child)
        parent.insert(idx_in_parent + i, new_p)


def normalize_soft_breaks(doc):
    for p in _collect_all_paragraphs(doc):
        _split_paragraph_by_br(p._element)


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
            "_idx": len(rows),
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
                        "_idx": len(flights),
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
# 文本覆盖率检查
# =========================================================
def check_text_coverage(excel_rows, text_flights, city_to_icao):
    domestic_rows = [
        r for r in excel_rows
        if (r["dep"].startswith("Z") or r["arr"].startswith("Z"))
        and r["dep"]
        and r["arr"]
    ]

    text_by_reg = {}
    for tf in text_flights:
        text_by_reg.setdefault(tf["reg"], []).append(tf)

    covered = 0
    missing = []
    for r in domestic_rows:
        candidates = text_by_reg.get(r["reg"], [])
        found = False
        for tf in candidates:
            dep_icao = city_to_icao.get(tf["dep_city"])
            arr_icao = city_to_icao.get(tf["arr_city"])
            if dep_icao == r["dep"] and arr_icao == r["arr"]:
                found = True
                break
        if found:
            covered += 1
        else:
            missing.append(r)

    return len(domestic_rows), covered, missing


# =========================================================
# 匹配
# =========================================================
def find_excel_match(approval, excel_rows, used_excel):
    base_candidates = [
        r for r in excel_rows
        if r["_idx"] not in used_excel
        and r["reg"] == approval["reg"]
        and r["dep"] == approval["dep"]
        and r["arr"] == approval["arr"]
    ]
    if not base_candidates:
        return None

    approval_date = approval["dep_dt_bj"].date()
    approval_dep_time = approval["dep_dt_bj"].strftime("%H:%M")

    same_date = [r for r in base_candidates if r["dep_date"] == approval_date]
    if same_date:
        same_date.sort(key=lambda r: time_diff_minutes(r["dep_time"], approval_dep_time))
        matched = same_date[0]
        used_excel.add(matched["_idx"])
        return matched

    close_date = []
    for r in base_candidates:
        if r["dep_date"] is None:
            continue
        delta_days = (approval_date - r["dep_date"]).days
        if abs(delta_days) != 1:
            continue
        real_diff = compute_real_minute_diff(
            approval_date, approval_dep_time,
            r["dep_date"], r["dep_time"]
        )
        if real_diff is not None and abs(real_diff) <= MAX_CROSS_DAY_GAP_MIN:
            close_date.append((abs(real_diff), r))

    if close_date:
        close_date.sort(key=lambda x: x[0])
        matched = close_date[0][1]
        used_excel.add(matched["_idx"])
        return matched

    return None


def find_text_match(approval, text_flights, city_to_icao, used_text):
    candidates = []
    for i, tf in enumerate(text_flights):
        if i in used_text:
            continue
        if tf["reg"] != approval["reg"]:
            continue
        if (city_to_icao.get(tf["dep_city"]) == approval["dep"]
                and city_to_icao.get(tf["arr_city"]) == approval["arr"]):
            candidates.append((i, tf))

    if not candidates:
        return None

    target = approval["dep_dt_bj"].strftime("%H:%M")
    candidates.sort(key=lambda x: time_diff_minutes(x[1]["dep_time"], target))
    idx, matched = candidates[0]
    used_text.add(idx)
    return matched


# =========================================================
# docx 标红/标绿/高亮/追加
# =========================================================
W_R = qn('w:r')
W_RPR = qn('w:rPr')
W_COLOR = qn('w:color')
W_T = qn('w:t')
W_HIGHLIGHT = qn('w:highlight')


def _set_run_color(run_element, color_hex):
    rPr = run_element.find(W_RPR)
    if rPr is None:
        rPr = run_element.makeelement(W_RPR, {})
        run_element.insert(0, rPr)
    for c in rPr.findall(W_COLOR):
        rPr.remove(c)
    color = rPr.makeelement(W_COLOR, {qn('w:val'): color_hex})
    rPr.append(color)


def _set_run_highlight(run_element, color_name=HIGHLIGHT_YELLOW):
    rPr = run_element.find(W_RPR)
    if rPr is None:
        rPr = run_element.makeelement(W_RPR, {})
        run_element.insert(0, rPr)
    for h in rPr.findall(W_HIGHLIGHT):
        rPr.remove(h)
    hl = rPr.makeelement(W_HIGHLIGHT, {qn('w:val'): color_name})
    rPr.append(hl)


def _make_run_like(src_run_elem, text, color_hex=None, highlight=None):
    new_r = copy.deepcopy(src_run_elem)
    for t in new_r.findall(W_T):
        new_r.remove(t)
    t = new_r.makeelement(W_T, {})
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    new_r.append(t)
    if color_hex:
        _set_run_color(new_r, color_hex)
    if highlight:
        _set_run_highlight(new_r, highlight)
    return new_r


def set_paragraph_runs(paragraph, text, color_overrides):
    if not color_overrides:
        return
    runs = list(paragraph.runs)
    if not runs:
        return
    full_text = "".join(r.text for r in runs)
    if not full_text:
        return

    char_color = [None] * len(full_text)
    for color_hex, parts in color_overrides:
        for part in parts:
            if not part:
                continue
            start = 0
            while True:
                idx = full_text.find(part, start)
                if idx == -1:
                    break
                for i in range(idx, idx + len(part)):
                    char_color[i] = color_hex
                start = idx + len(part)

    if not any(c is not None for c in char_color):
        return

    pos = 0
    for run in runs:
        r_text = run.text
        if not r_text:
            continue
        r_start = pos
        r_len = len(r_text)
        run_colors = [char_color[r_start + i] for i in range(r_len)]

        unique_colors = set(run_colors)
        if len(unique_colors) == 1:
            color = run_colors[0]
            if color is None:
                pos += r_len
                continue
            _set_run_color(run._element, color)
            pos += r_len
            continue

        run_elem = run._element
        parent = run_elem.getparent()
        idx_in_parent = list(parent).index(run_elem)

        pieces = []
        i = 0
        while i < r_len:
            c = run_colors[i]
            j = i + 1
            while j < r_len and run_colors[j] == c:
                j += 1
            pieces.append((r_text[i:j], c))
            i = j

        parent.remove(run_elem)
        for k, (seg, color) in enumerate(pieces):
            new_r = _make_run_like(run_elem, seg, color)
            parent.insert(idx_in_parent + k, new_r)

        pos += r_len


def append_red_text(paragraph, text):
    runs = list(paragraph.runs)
    p_elem = paragraph._element
    if runs:
        src = runs[-1]._element
        new_r = _make_run_like(src, text, color_hex=RED, highlight=HIGHLIGHT_YELLOW)
    else:
        new_r = p_elem.makeelement(W_R, {})
        rPr = new_r.makeelement(W_RPR, {})
        new_r.insert(0, rPr)
        color = rPr.makeelement(W_COLOR, {qn('w:val'): RED})
        rPr.append(color)
        hl = rPr.makeelement(W_HIGHLIGHT, {qn('w:val'): HIGHLIGHT_YELLOW})
        rPr.append(hl)
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
    for h in rPr.findall(W_HIGHLIGHT):
        rPr.remove(h)

    color = OxmlElement('w:color')
    color.set(qn('w:val'), RED)
    rPr.append(color)

    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), HIGHLIGHT_YELLOW)
    rPr.append(hl)

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
# 构造"待申请"简洁文本
# =========================================================
def build_approval_text(excel_row, is_domestic, note_kind=""):
    reg = excel_row["reg"]
    dep_date = excel_row["dep_date"]
    if not dep_date:
        return None

    date_str = dep_date.strftime("%d%b").upper()
    parts = [reg, f"{excel_row['dep']}-{excel_row['arr']}", date_str]

    if is_domestic:
        if note_kind == "中国籍":
            parts.append("中国籍")
        elif note_kind == "外籍":
            parts.append("外籍")
        elif note_kind == "机组未定":
            parts.append("机组未定")

    parts.append("待申请")
    return " ".join(parts)


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
    normalize_soft_breaks(doc)

    result_rows = []
    approval_red_map = {}
    approval_green_map = {}
    cancel_paragraphs = set()
    change_paragraphs = set()

    used_excel = set()
    used_text = set()
    unapproved_keys = set()

    # ===== 第一遍：核对 docx 已有批复 =====
    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if not raw_text:
            continue

        approval = parse_approval_line(raw_text)
        if not approval:
            continue

        if not approval["is_domestic"] and approval["flight_no"]:
            unapproved_keys.add(
                (approval["reg"], approval["dep"], approval["arr"])
            )
            result_rows.append({
                "批复": raw_text,
                "飞机号": approval["reg"],
                "航班号": approval["flight_no"],
                "机型": AIRCRAFT_TYPE_MAP.get(approval["reg"], ""),
                "内/外机": "外机",
                "批复起飞(北京时)": approval["dep_dt_bj"].strftime("%H:%M"),
                "批复落地(北京时)": approval["arr_dt_bj"].strftime("%H:%M"),
                "批复日期": approval["dep_dt_bj"].date().isoformat(),
                "Excel 用途": "",
                "文本标记": "",
                "Excel 计划起飞": "",
                "Excel 计划落地": "",
                "Excel 出发地": "",
                "Excel 到达地": "",
                "差异": "外机未批（已申请）",
                "是否一致": "待确认",
                "备注": "未批",
            })
            continue

        red_parts = []
        green_parts = []
        diffs = []
        note_parts = []
        info_note = ""

        excel_row = find_excel_match(approval, excel_rows, used_excel)

        if excel_row is None:
            note_parts.append("待取消")
            red_parts.append(approval["dep"])
            red_parts.append(approval["arr"])
            red_parts.append(approval["date_raw"])

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
                "差异": f"Excel 中无 {approval['dep']}→{approval['arr']}（±1 天）匹配",
                "是否一致": "否",
                "备注": "待取消",
            })
            approval_red_map[raw_text] = red_parts
            cancel_paragraphs.add(raw_text)
            continue

        text_flight = find_text_match(
            approval, text_flights, city_to_icao, used_text
        )

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

        ap_dur = flight_duration_minutes(approval_dep_time, approval_arr_time)
        xl_dur = flight_duration_minutes(excel_row["dep_time"], excel_row["arr_time"])
        dur_diff = None
        if ap_dur is not None and xl_dur is not None:
            dur_diff = ap_dur - xl_dur
            if abs(dur_diff) > 30:
                diffs.append(
                    f"飞行时长：批复 {fmt_duration(ap_dur)} vs 计划 {fmt_duration(xl_dur)}"
                    f"（差 {dur_diff:+d} 分钟）"
                )
                red_parts.append(approval["arr_time_raw"])

        real_dep_diff = compute_real_minute_diff(
            approval["dep_dt_bj"].date(), approval_dep_time,
            excel_row["dep_date"], excel_row["dep_time"]
        )
        dep_ok = False
        if real_dep_diff is None:
            dep_ok = False
        elif real_dep_diff == 0:
            dep_ok = True
        elif -EARLY_GREEN_THRESHOLD_MIN <= real_dep_diff < 0:
            dep_ok = True
            green_parts.append(approval["dep_time_raw"])
        elif real_dep_diff > 0:
            dep_ok = False
            diffs.append(
                f"起飞时间：批复 {approval_dep_time} 晚于计划 {excel_row['dep_time']}，需重新申请"
            )
            red_parts.append(approval["dep_time_raw"])
            change_paragraphs.add(raw_text)
        else:
            dep_ok = False
            diffs.append(
                f"起飞时间：批复 {approval_dep_time} vs 计划 {excel_row['dep_time']}"
                f"（相差 {real_dep_diff} 分钟）"
            )
            red_parts.append(approval["dep_time_raw"])

        xl_arr_date = excel_row["arr_date"] or excel_row["dep_date"]
        real_arr_diff = compute_real_minute_diff(
            approval["arr_dt_bj"].date(), approval_arr_time,
            xl_arr_date, excel_row["arr_time"]
        )
        if real_arr_diff is not None and real_arr_diff != 0:
            if -EARLY_GREEN_THRESHOLD_MIN <= real_arr_diff < 0 and dep_ok:
                green_parts.append(approval["arr_time_raw"])

        excel_ferry = is_ferry_use(excel_row["use"])
        expected_service = "N/M" if excel_ferry else "U/H"

        if approval["service"] != expected_service:
            diffs.append(
                f"用途：批复 {approval['service']} vs 计划 {excel_row['use']}"
                f"（应为 {expected_service}）"
            )
            red_parts.append(approval["service"])

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
                    info_note = "待确认机组"
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
                info_note = "待确认机组"

        if raw_text in change_paragraphs:
            note_parts.append("待变更")

        final_notes = list(note_parts)
        if info_note:
            final_notes.append(info_note)

        if diffs or note_parts:
            consistency = "否"
        elif info_note:
            consistency = "待确认"
        else:
            consistency = "是"

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
            "是否一致": consistency,
            "备注": "；".join(final_notes),
        })

        if red_parts:
            approval_red_map[raw_text] = red_parts
        if green_parts:
            approval_green_map[raw_text] = green_parts

    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        red = approval_red_map.get(raw_text, [])
        green = approval_green_map.get(raw_text, [])
        if red or green:
            set_paragraph_runs(
                p, raw_text,
                [(GREEN, green), (RED, red)]
            )

    for p in iter_doc_paragraphs(doc):
        raw_text = p.text.strip()
        if raw_text in cancel_paragraphs:
            append_red_text(p, "  待取消")
        elif raw_text in change_paragraphs:
            append_red_text(p, "  待变更")

    # ===== 第二遍：待申请 =====
    pending_by_reg = {}
    for row in excel_rows:
        if not row["reg"]:
            continue
        if not (row["dep"].startswith("Z") or row["arr"].startswith("Z")):
            continue
        if row["_idx"] in used_excel:
            continue
        if row["dep_date"] is None:
            continue
        if (row["reg"], row["dep"], row["arr"]) in unapproved_keys:
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
                tf = find_text_match(fake_ap, text_flights, city_to_icao, used_text)
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

            text = build_approval_text(row, is_domestic, note_kind)
            if not text:
                continue

            items.append({
                "dt": dep_dt_bj,
                "text": text,
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
                diff_text = "Excel 有计划，批复汇总表缺失；文本未提供该航段，机组需确认"
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
        "- **软换行自动拆分**：Shift+Enter 产生的软换行会被拆成独立段落\n"
        "- **文本覆盖率检查**：粘贴文本后立即执行，缺失航段可展开复制\n"
        "- **Excel 匹配**：注册号 + 起降机场 + 日期（同日优先；其次差 1 天且真实时差 ≤ 10 小时）\n"
        "- **文本匹配**：注册号 + 起降城市 完全一致（一条只能用一次）\n"
        "- **外机格式**：已批 `注册号 机型 ... FERRY/BUSINESS`；未批 `注册号 航班号(=注册号) ...`\n"
        "- **用途映射**：`FERRY`→`N/M`；`BUSINESS`→`U/H`\n"
        "- **三种结果**：是 / 否 / 待确认\n"
        "- **待取消 / 待变更**：段落末尾追加**红字黄底**文字\n"
        "- **待申请**：简洁格式插入（如 `VPCSZ ZGSZ-ZBAD 07OCT 待申请`），**红字黄底**\n"
        "- **已有批复标红/标绿**：只改字体颜色，不加黄底"
    )

text_input = st.text_area(
    "③ 粘贴文本版航班信息",
    height=320,
    placeholder=(
        "例如：\n"
        "B65AP 16:30 - 17:45\n"
        "香港 - 泉州晋江\n"
        "P057,P039,C046,M035"
    ),
)

text_content = text_input.strip()

# ===== 文本覆盖率检查（粘贴后立即执行，不需要点按钮）=====
if docx_file and excel_file and text_content:
    preview_excel_rows = load_excel_rows_from_bytes(excel_file.getvalue())
    preview_text_flights = load_text_flights(text_content)

    preview_city_to_icao = {}
    for row in preview_excel_rows:
        if row["dep_city"]:
            preview_city_to_icao[row["dep_city"]] = row["dep"]
        if row["arr_city"]:
            preview_city_to_icao[row["arr_city"]] = row["arr"]

    total, covered, missing = check_text_coverage(
        preview_excel_rows, preview_text_flights, preview_city_to_icao
    )

    st.subheader("🔍 文本覆盖率检查")

    if total == 0:
        st.warning("Excel 里没有国内航段（Z 开头机场），无需核对。")
    else:
        coverage = covered / total

        if coverage >= COVERAGE_WARN_THRESHOLD:
            st.success(
                f"✅ 文本计划覆盖率：{covered}/{total}（{coverage*100:.1f}%）"
            )
        elif coverage >= COVERAGE_ERROR_THRESHOLD:
            st.warning(
                f"⚠️ 文本计划覆盖率：{covered}/{total}（{coverage*100:.1f}%），"
                f"缺 {total - covered} 条。可以核对，但建议补全。"
            )
        else:
            st.error(
                f"❌ 文本计划覆盖率过低：{covered}/{total}（{coverage*100:.1f}%），"
                f"缺 {total - covered} 条。请先补全文本。"
            )

        if missing:
            with st.expander(
                f"📋 缺失航段明细（{len(missing)} 条）—— 点开查看 / 复制",
                expanded=(coverage < COVERAGE_WARN_THRESHOLD),
            ):
                missing_sorted = sorted(
                    missing,
                    key=lambda r: (
                        r["dep_date"] or datetime.date.min,
                        r["reg"],
                        r["dep_time"],
                    ),
                )
                lines = []
                for r in missing_sorted:
                    date_str = r["dep_date"].strftime("%m-%d") if r["dep_date"] else "??-??"
                    lines.append(
                        f"{r['reg']}  {date_str}  {r['dep']}-{r['arr']}  "
                        f"{r['dep_time']}-{r['arr_time']}  ({r['dep_city']} → {r['arr_city']})"
                    )
                st.code("\n".join(lines), language=None)

# ===== 正式核对 =====
if st.button(
    "🚀 开始核对",
    type="primary",
    disabled=not (docx_file and excel_file and text_content),
):
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
    total_rows = len(df)
    diff_count = (df["是否一致"] == "否").sum() if total_rows else 0
    pending_count = (df["是否一致"] == "待确认").sum() if total_rows else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("批复条数", total_rows)
    c2.metric("一致", total_rows - diff_count - pending_count)
    c3.metric("有差异", diff_count)
    c4.metric("待确认", pending_count)

    st.subheader("📋 核对结果")
    if total_rows == 0:
        st.warning("未在 docx 中识别到任何批复行。")
    else:
        def highlight(row):
            if row["是否一致"] == "待确认":
                return ["background-color: #fff3cd"] * len(row)
            if row["是否一致"] == "否":
                if "待变更" in str(row["备注"]):
                    return ["background-color: #e5f0ff"] * len(row)
                return ["background-color: #ffe5e5"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(highlight, axis=1),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("🚨 差异明细")
        diffs_df = df[df["是否一致"] != "是"][["批复", "差异", "备注", "是否一致"]]
        if diffs_df.empty:
            st.success("✅ 所有批复与计划一致，未发现差异。")
        else:
            for _, r in diffs_df.iterrows():
                if r["是否一致"] == "待确认":
                    note_html = (
                        " <span style='color:#d97706;font-weight:bold'>"
                        f"【{r['备注']}】</span>"
                    )
                elif "待变更" in str(r["备注"]):
                    note_html = (
                        " <span style='color:#0066cc;font-weight:bold'>"
                        f"【{r['备注']}】</span>"
                    )
                elif r["备注"]:
                    note_html = (
                        f" <span style='color:red;font-weight:bold'>【{r['备注']}】</span>"
                    )
                else:
                    note_html = ""
                st.markdown(f"**`{r['批复']}`**{note_html}", unsafe_allow_html=True)
                if r["差异"] != "无":
                    for line in r["差异"].split("；"):
                        st.markdown(f"- {line}")

    st.subheader("📥 下载标红后的批复汇总表")
    st.download_button(
        label=f"下载 {docx_file.name}",
        data=out_buf.getvalue(),
        file_name=docx_file.name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
