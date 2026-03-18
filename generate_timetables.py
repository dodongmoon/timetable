#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

from PIL import Image, ImageDraw, ImageFilter, ImageFont


DEFAULT_SOURCE = Path("/Users/yangdongmoon/Downloads/3-2반 1학기 시간표.xlsx")
DEFAULT_OUTPUT = Path("dist/3-2-학생별-시간표.html")
FONT_PATH = "/System/Library/Fonts/AppleSDGothicNeo.ttc"

PDF_REL_DIR = Path("assets/pdf")
PDF_ZIP_REL_PATH = Path("assets/3-2-학생별-시간표-pdf.zip")

SPREADSHEET_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

DAYS = ["월", "화", "수", "목", "금"]
CHOICE_LABELS = ["A", "B", "C", "D", "E", "F", "G", "H"]

BASE_TIMETABLE = [
    {"type": "class", "period": "1", "time": "08:40-09:30", "slots": ["E", "F", "C", "B", "H"]},
    {"type": "class", "period": "2", "time": "09:40-10:30", "slots": ["미창", "F", "D", "E", "H"]},
    {"type": "class", "period": "3", "time": "10:40-11:30", "slots": ["A", "E", "미창", "D", "C"]},
    {"type": "class", "period": "4", "time": "11:40-12:30", "slots": ["A", "E", "A", "D", "C"]},
    {"type": "break", "label": "점심시간"},
    {"type": "class", "period": "5", "time": "13:30-14:20", "slots": ["B", "G", "B", "C", "창체"]},
    {"type": "class", "period": "6", "time": "14:30-15:20", "slots": ["진로", "G", "B", "스생", "창체"]},
    {"type": "break", "label": "청소시간"},
    {"type": "class", "period": "7", "time": "15:40-16:30", "slots": ["", "D", "", "A", ""]},
]

COMMON_SUBJECT_STYLES = {
    "미창": "common-art",
    "진로": "common-career",
    "스생": "common-sports",
    "창체": "common-club",
}

CHOICE_COLORS = {
    "A": {"html": "rgba(247, 220, 197, 0.82)", "image": (247, 223, 200)},
    "B": {"html": "rgba(200, 199, 232, 0.76)", "image": (204, 203, 235)},
    "C": {"html": "rgba(215, 214, 216, 0.78)", "image": (220, 219, 220)},
    "D": {"html": "rgba(200, 234, 223, 0.82)", "image": (203, 234, 223)},
    "E": {"html": "rgba(207, 217, 243, 0.8)", "image": (210, 220, 244)},
    "F": {"html": "rgba(224, 212, 234, 0.82)", "image": (228, 217, 237)},
    "G": {"html": "rgba(255, 241, 184, 0.82)", "image": (255, 243, 188)},
    "H": {"html": "rgba(238, 240, 199, 0.88)", "image": (238, 241, 201)},
}

COMMON_COLORS = {
    "미창": {"html": "#f7efe5", "image": (247, 239, 229)},
    "진로": {"html": "#f9efe6", "image": (249, 240, 230)},
    "스생": {"html": "#ece8e4", "image": (236, 232, 228)},
    "창체": {"html": "#f4eadc", "image": (244, 235, 222)},
}

CHOICE_DESTINATIONS = {
    "A": {
        "화법과작문_1반": "1반",
        "화법과작문_2반": "2반",
        "고전읽기": "3반",
        "미적분": "4반",
        "확률과통계": "5반",
        "경제수학": "6반",
        "영어독해와작문": "7반",
        "영미문학읽기": "8반",
        "사회문화": "9반",
        "생활과윤리": "10반",
        "화학Ⅱ": "11반",
        "생명과학Ⅱ": "12반",
    },
    "B": {
        "화법과작문": "1반",
        "언어와매체": "2반",
        "현대문학감상": "3반",
        "미적분": "4반",
        "확률과통계_1반": "5반",
        "확률과통계_2반": "6반",
        "영어독해와작문": "7반",
        "진로영어": "8반",
        "한국지리": "9반",
        "동아시아사": "10반",
        "생활과윤리": "11반",
        "화학Ⅱ": "12반",
        "물리학Ⅱ": "물리지구과학실",
    },
    "C": {
        "화법과작문_1반": "1반",
        "화법과작문_2반": "2반",
        "언어와매체": "3반",
        "확률과통계_1반": "4반",
        "확률과통계_2반": "5반",
        "영어독해와작문": "6반",
        "영미문학읽기": "7반",
        "진로영어": "8반",
        "정치와법": "9반",
        "사회문제탐구": "10반",
        "물리학Ⅱ": "11반",
        "생명과학Ⅱ": "12반",
    },
    "D": {
        "화법과작문_1반": "1반",
        "화법과작문_2반": "2반",
        "언어와매체": "3반",
        "확률과통계": "4반",
        "경제수학": "5반",
        "영어독해와작문": "6반",
        "한국지리": "7반",
        "여행지리": "8반",
        "사회문화": "9반",
        "생활과윤리": "10반",
        "화학Ⅱ": "11반",
        "생활과과학": "12반",
    },
    "E": {
        "화법과작문": "1반",
        "현대문학감상": "2반",
        "미적분": "3반",
        "확률과통계_1반": "4반",
        "확률과통계_2반": "5반",
        "진로영어": "6반",
        "여행지리": "7반",
        "사회문화": "8반",
        "생활과윤리": "9반",
        "고전과윤리": "10반",
        "생명과학Ⅱ": "11반",
        "지구과학Ⅱ": "물리지구과학실",
    },
    "F": {
        "시창청음_1반": "1반",
        "시창청음_2반": "2반",
        "평면조형_1반": "3반",
        "평면조형_2반": "4반",
        "공학일반": "5반",
        "가정과학": "6반",
        "인공지능과피지컬컴퓨팅": "7반",
        "일본어Ⅱ": "8반",
        "중국어Ⅱ": "9반",
        "한문Ⅱ_1반": "10반",
        "한문Ⅱ_2반": "11반",
        "사물인터넷서비스기획": "빅데이터분석실",
    },
    "G": {
        "시창청음_1반": "1반",
        "시창청음_2반": "2반",
        "평면조형_1반": "3반",
        "평면조형_2반": "4반",
        "공학일반": "5반",
        "가정과학": "6반",
        "인공지능과피지컬컴퓨팅": "7반",
        "일본어Ⅱ": "8반",
        "중국어Ⅱ": "9반",
        "한문Ⅱ_1반": "10반",
        "한문Ⅱ_2반": "11반",
        "사물인터넷서비스기획": "빅데이터분석실",
    },
    "H": {
        "시창청음_1반": "1반",
        "시창청음_2반": "2반",
        "평면조형_1반": "3반",
        "평면조형_2반": "4반",
        "공학일반": "5반",
        "가정과학": "6반",
        "사물인터넷서비스기획": "7반",
        "인공지능과피지컬컴퓨팅": "8반",
        "일본어Ⅱ": "9반",
        "중국어Ⅱ": "10반",
        "한문Ⅱ_1반": "11반",
        "한문Ⅱ_2반": "12반",
    },
}

PDF_CANVAS = (2262, 1600)


@dataclass
class Student:
    student_no: str
    name: str
    choices: dict[str, str]


def extract_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("a:si", SPREADSHEET_NS):
        strings.append("".join(text.text or "" for text in item.findall(".//a:t", SPREADSHEET_NS)))
    return strings


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    raw_value = cell.findtext("a:v", default="", namespaces=SPREADSHEET_NS)
    if cell_type == "s":
        return shared_strings[int(raw_value)]
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//a:t", SPREADSHEET_NS))
    return raw_value


def parse_students(source: Path) -> list[Student]:
    with zipfile.ZipFile(source) as archive:
        shared_strings = extract_shared_strings(archive)
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))

    students: list[Student] = []
    for row in sheet.findall(".//a:sheetData/a:row", SPREADSHEET_NS):
        if int(row.attrib["r"]) == 1:
            continue

        values: dict[str, str] = {}
        for cell in row.findall("a:c", SPREADSHEET_NS):
            reference = cell.attrib["r"]
            column = re.sub(r"\d+", "", reference)
            values[column] = read_cell_value(cell, shared_strings).strip()

        choices = {label: values.get(column, "") for label, column in zip(CHOICE_LABELS, "CDEFGHIJ")}
        students.append(Student(student_no=values.get("A", ""), name=values.get("B", ""), choices=choices))

    students.sort(key=lambda student: student.student_no)
    return students


def subject_length_class(subject: str) -> str:
    size = len(subject.replace("_", ""))
    if size >= 16:
        return "subject tightest"
    if size >= 12:
        return "subject tighter"
    if size >= 9:
        return "subject tight"
    return "subject"


def format_subject(subject: str) -> str:
    return "<br>".join(html.escape(piece) for piece in subject.split("_"))


def pdf_href(student: Student) -> str:
    return f"{PDF_REL_DIR.as_posix()}/{student.student_no}.pdf"


def pdf_download_name(student: Student) -> str:
    return f"{student.student_no}_{student.name}_시간표.pdf"


def destination_text(choice_label: str, subject: str) -> str:
    destination = CHOICE_DESTINATIONS.get(choice_label, {}).get(subject, "")
    return f"({destination})" if destination else ""


def render_slot(token: str, student: Student) -> str:
    if token in CHOICE_LABELS:
        subject = student.choices[token]
        room = destination_text(token, subject)
        room_html = f'<div class="subject-room">{html.escape(room)}</div>' if room else ""
        return (
            f'<td class="slot choice choice-{token.lower()}">'
            f'<div class="slot-badge">선택{token}</div>'
            f'<div class="{subject_length_class(subject)}">{format_subject(subject)}</div>'
            f"{room_html}"
            f"</td>"
        )

    if not token:
        return '<td class="slot empty"><div class="empty-mark">-</div></td>'

    style = COMMON_SUBJECT_STYLES.get(token, "common")
    return (
        f'<td class="slot common {style}">'
        f'<div class="slot-badge slot-badge-common">공통</div>'
        f'<div class="subject common-subject">{html.escape(token)}</div>'
        f"</td>"
    )


def render_rows(student: Student) -> str:
    rows: list[str] = []
    for entry in BASE_TIMETABLE:
        if entry["type"] == "break":
            rows.append(
                '<tr class="divider-row">'
                f'<td class="divider-cell" colspan="7">{html.escape(entry["label"])}</td>'
                "</tr>"
            )
            continue

        slots_html = "".join(render_slot(token, student) for token in entry["slots"])
        rows.append(
            "<tr>"
            f'<th class="period" scope="row">{html.escape(entry["period"])}</th>'
            f'<td class="time">{html.escape(entry["time"])}</td>'
            f"{slots_html}"
            "</tr>"
        )
    return "\n".join(rows)


def render_choice_list(student: Student) -> str:
    items = []
    for label in CHOICE_LABELS:
        items.append(
            f'<li class="choice-item choice-{label.lower()}">'
            f'<span class="choice-key">선택{label}</span>'
            f'<span class="choice-name">{html.escape(student.choices[label])}</span>'
            "</li>"
        )
    return "\n".join(items)


def render_page(student: Student) -> str:
    return f"""
    <section class="student-page is-visible" data-name="{html.escape(student.name.lower())}" data-student-no="{html.escape(student.student_no)}">
      <div class="page-accent"></div>
      <header class="page-header">
        <div class="student-heading">
          <div>
            <p class="eyebrow">2026학년도 1학기 3학년 2반 개인 시간표</p>
            <h2 class="student-name">{html.escape(student.name)}</h2>
            <p class="student-meta">학번 {html.escape(student.student_no)}</p>
          </div>
          <div class="page-actions">
            <a class="btn secondary page-download" data-pdf-link href="{html.escape(pdf_href(student))}" download="{html.escape(pdf_download_name(student))}">PDF 다운로드</a>
          </div>
        </div>
        <aside class="choice-panel">
          <p class="panel-title">선택 과목</p>
          <ul class="choice-list">
            {render_choice_list(student)}
          </ul>
        </aside>
      </header>

      <div class="timetable-shell">
        <table class="timetable">
          <thead>
            <tr>
              <th class="period-head">교시</th>
              <th class="time-head">시간</th>
              <th>월</th>
              <th>화</th>
              <th>수</th>
              <th>목</th>
              <th>금</th>
            </tr>
          </thead>
          <tbody>
            {render_rows(student)}
          </tbody>
        </table>
      </div>
    </section>
    """.strip()


def build_document(students: list[Student]) -> str:
    template = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>3학년 2반 학생별 시간표</title>
  <style>
    :root {
      --paper: #fffdf8;
      --ink: #2a2521;
      --muted: #756a61;
      --line: #d8cfc7;
      --card-shadow: 0 24px 70px rgba(80, 58, 37, 0.12);
      --choice-a: __CHOICE_A__;
      --choice-b: __CHOICE_B__;
      --choice-c: __CHOICE_C__;
      --choice-d: __CHOICE_D__;
      --choice-e: __CHOICE_E__;
      --choice-f: __CHOICE_F__;
      --choice-g: __CHOICE_G__;
      --choice-h: __CHOICE_H__;
      --common-art: #f7efe5;
      --common-career: #f9efe6;
      --common-sports: #ece8e4;
      --common-club: #f4eadc;
      --empty: #f4f1ed;
      --accent: #ca8b57;
    }

    * {
      box-sizing: border-box;
    }

    html {
      scroll-behavior: smooth;
    }

    body {
      margin: 0;
      color: var(--ink);
      font-family: "Apple SD Gothic Neo", "Noto Sans KR", sans-serif;
      background:
        radial-gradient(circle at top left, rgba(202, 139, 87, 0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(124, 170, 148, 0.18), transparent 24%),
        linear-gradient(180deg, #f8f3ec 0%, #f1ece6 100%);
    }

    .app {
      width: min(1520px, calc(100vw - 32px));
      margin: 24px auto 64px;
    }

    .toolbar {
      position: sticky;
      top: 16px;
      z-index: 10;
      display: flex;
      gap: 12px;
      align-items: center;
      justify-content: space-between;
      padding: 16px 18px;
      margin-bottom: 20px;
      border: 1px solid rgba(159, 145, 133, 0.24);
      border-radius: 24px;
      backdrop-filter: blur(16px);
      background: rgba(255, 252, 246, 0.82);
      box-shadow: 0 18px 40px rgba(93, 70, 48, 0.08);
    }

    .toolbar-main {
      display: flex;
      gap: 14px;
      align-items: center;
      flex-wrap: wrap;
    }

    .toolbar h1 {
      margin: 0;
      font-size: 1.2rem;
    }

    .toolbar-meta {
      margin: 2px 0 0;
      color: var(--muted);
      font-size: 0.92rem;
    }

    .toolbar-actions {
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }

    .search-input {
      min-width: 240px;
      padding: 12px 14px;
      border: 1px solid rgba(159, 145, 133, 0.55);
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.9);
      color: var(--ink);
      font-size: 0.96rem;
      outline: none;
    }

    .search-input:focus {
      border-color: var(--accent);
      box-shadow: 0 0 0 4px rgba(202, 139, 87, 0.14);
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      text-decoration: none;
      border: 0;
      border-radius: 999px;
      padding: 12px 16px;
      cursor: pointer;
      font-size: 0.94rem;
      font-weight: 700;
      color: white;
      background: linear-gradient(135deg, #b56f3d 0%, #ce9566 100%);
      box-shadow: 0 10px 24px rgba(181, 111, 61, 0.25);
    }

    .btn.secondary {
      color: var(--ink);
      background: white;
      border: 1px solid rgba(159, 145, 133, 0.42);
      box-shadow: none;
    }

    .status-pill {
      padding: 10px 14px;
      border-radius: 999px;
      background: rgba(124, 170, 148, 0.12);
      color: #456a56;
      font-size: 0.9rem;
      font-weight: 700;
    }

    .pages {
      display: flex;
      flex-direction: column;
      gap: 26px;
    }

    .student-page {
      position: relative;
      overflow: hidden;
      padding: 26px 28px 28px;
      border: 1px solid rgba(159, 145, 133, 0.32);
      border-radius: 32px;
      background:
        linear-gradient(180deg, rgba(255, 247, 236, 0.96) 0%, rgba(255, 255, 255, 0.96) 100%),
        var(--paper);
      box-shadow: var(--card-shadow);
    }

    .student-page.is-hidden {
      display: none;
    }

    .page-accent {
      position: absolute;
      inset: 0 auto auto 0;
      width: 100%;
      height: 12px;
      background: linear-gradient(90deg, #ce9566 0%, #e1bc75 34%, #8db69c 68%, #88a8dc 100%);
    }

    .page-header {
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(320px, 0.9fr);
      gap: 18px;
      align-items: start;
      margin-bottom: 20px;
      padding-top: 12px;
    }

    .student-heading {
      display: flex;
      gap: 18px;
      align-items: flex-start;
      justify-content: space-between;
    }

    .page-actions {
      display: flex;
      gap: 10px;
      flex-shrink: 0;
    }

    .page-download {
      min-width: 132px;
    }

    .eyebrow {
      margin: 0 0 10px;
      color: var(--muted);
      font-size: 0.95rem;
      font-weight: 700;
      letter-spacing: 0.02em;
    }

    .student-name {
      margin: 0;
      font-size: clamp(2.1rem, 3.3vw, 3.1rem);
      line-height: 1;
    }

    .student-meta {
      margin: 10px 0 0;
      font-size: 1.08rem;
      color: var(--muted);
      font-weight: 700;
    }

    .choice-panel {
      padding: 18px;
      border: 1px solid rgba(159, 145, 133, 0.25);
      border-radius: 24px;
      background:
        linear-gradient(145deg, rgba(255, 251, 244, 0.95), rgba(252, 246, 237, 0.88));
    }

    .panel-title {
      margin: 0 0 12px;
      font-size: 0.98rem;
      font-weight: 800;
    }

    .choice-list {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin: 0;
      padding: 0;
      list-style: none;
    }

    .choice-item {
      min-width: 0;
      padding: 10px 12px;
      border-radius: 18px;
      border: 1px solid rgba(159, 145, 133, 0.18);
    }

    .choice-key,
    .choice-name {
      display: block;
    }

    .choice-key {
      margin-bottom: 4px;
      font-size: 0.82rem;
      color: var(--muted);
      font-weight: 800;
    }

    .choice-name {
      font-size: 0.92rem;
      font-weight: 700;
      line-height: 1.3;
      word-break: keep-all;
    }

    .choice-a { background: var(--choice-a); }
    .choice-b { background: var(--choice-b); }
    .choice-c { background: var(--choice-c); }
    .choice-d { background: var(--choice-d); }
    .choice-e { background: var(--choice-e); }
    .choice-f { background: var(--choice-f); }
    .choice-g { background: var(--choice-g); }
    .choice-h { background: var(--choice-h); }

    .timetable-shell {
      overflow: hidden;
      border: 1px solid rgba(159, 145, 133, 0.32);
      border-radius: 24px;
      background: rgba(255, 255, 255, 0.8);
    }

    .timetable {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }

    .timetable th,
    .timetable td {
      border: 1px solid rgba(159, 145, 133, 0.44);
      text-align: center;
      padding: 10px 8px;
      vertical-align: middle;
    }

    .timetable thead th {
      padding: 12px 8px;
      background: #f6eee4;
      font-size: 0.97rem;
      font-weight: 800;
    }

    .period-head,
    .period {
      width: 58px;
      background: #fffaf2;
      font-weight: 800;
    }

    .time-head,
    .time {
      width: 120px;
      background: #fffcf7;
      color: var(--muted);
      font-weight: 700;
      font-size: 0.9rem;
    }

    .slot {
      height: 108px;
      padding: 8px;
    }

    .slot-badge {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 62px;
      margin-bottom: 10px;
      padding: 5px 10px;
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.68);
      color: rgba(56, 43, 34, 0.86);
      font-size: 0.78rem;
      font-weight: 800;
    }

    .slot-badge-common {
      background: rgba(255, 255, 255, 0.88);
    }

    .subject {
      font-weight: 800;
      line-height: 1.28;
      word-break: keep-all;
      overflow-wrap: anywhere;
    }

    .subject.tight {
      font-size: 0.95rem;
    }

    .subject.tighter {
      font-size: 0.88rem;
    }

    .subject.tightest {
      font-size: 0.8rem;
    }

    .subject-room {
      margin-top: 6px;
      color: var(--muted);
      font-size: 0.74rem;
      font-weight: 700;
      line-height: 1.2;
    }

    .choice.choice-a { background: var(--choice-a); }
    .choice.choice-b { background: var(--choice-b); }
    .choice.choice-c { background: var(--choice-c); }
    .choice.choice-d { background: var(--choice-d); }
    .choice.choice-e { background: var(--choice-e); }
    .choice.choice-f { background: var(--choice-f); }
    .choice.choice-g { background: var(--choice-g); }
    .choice.choice-h { background: var(--choice-h); }

    .common-art { background: var(--common-art); }
    .common-career { background: var(--common-career); }
    .common-sports { background: var(--common-sports); }
    .common-club { background: var(--common-club); }

    .empty {
      background: var(--empty);
    }

    .empty-mark {
      color: rgba(117, 106, 97, 0.72);
      font-size: 1.2rem;
      font-weight: 700;
    }

    .divider-cell {
      padding: 8px 12px;
      background:
        linear-gradient(90deg, rgba(202, 139, 87, 0.11), rgba(124, 170, 148, 0.11));
      color: #5b4e45;
      font-size: 0.88rem;
      font-weight: 800;
      letter-spacing: 0.06em;
    }

    @media (max-width: 1080px) {
      .app {
        width: min(100vw - 18px, 1520px);
        margin: 12px auto 36px;
      }

      .toolbar {
        position: static;
      }

      .page-header {
        grid-template-columns: 1fr;
      }

      .student-heading {
        flex-direction: column;
        align-items: flex-start;
      }

      .page-actions {
        width: 100%;
      }

      .choice-list {
        grid-template-columns: 1fr;
      }

      .timetable-shell {
        overflow-x: auto;
      }

      .timetable {
        min-width: 860px;
      }
    }
  </style>
</head>
<body>
  <main class="app">
    <section class="toolbar">
      <div class="toolbar-main">
        <div>
          <h1>3학년 2반 학생별 시간표</h1>
          <p class="toolbar-meta">총 __COUNT__명 · 생성 시각 __TIMESTAMP__ · 학생별 PDF를 바로 내려받을 수 있습니다.</p>
        </div>
        <span class="status-pill" id="visibleCount">현재 __COUNT__명 표시 중</span>
      </div>
      <div class="toolbar-actions">
        <input id="searchInput" class="search-input" type="search" placeholder="이름 또는 학번 검색">
        <button id="resetButton" class="btn secondary" type="button">전체 보기</button>
        <button id="downloadVisibleButton" class="btn" type="button">검색한 시간표 PDF</button>
        <a class="btn secondary" href="__ZIP_HREF__" download>전체 PDF ZIP</a>
      </div>
    </section>

    <section class="pages" id="pages">
      __PAGES__
    </section>
  </main>

  <script>
    const searchInput = document.getElementById("searchInput");
    const resetButton = document.getElementById("resetButton");
    const downloadVisibleButton = document.getElementById("downloadVisibleButton");
    const visibleCount = document.getElementById("visibleCount");
    const cards = Array.from(document.querySelectorAll(".student-page"));

    function normalize(value) {
      return value.toLowerCase().trim();
    }

    function getVisibleCards() {
      return cards.filter((card) => !card.classList.contains("is-hidden"));
    }

    function applyFilter() {
      const query = normalize(searchInput.value);
      let count = 0;

      cards.forEach((card) => {
        const haystack = normalize(card.dataset.name) + " " + normalize(card.dataset.studentNo);
        const matched = query === "" || haystack.includes(query);
        card.classList.toggle("is-hidden", !matched);
        card.classList.toggle("is-visible", matched);
        if (matched) count += 1;
      });

      visibleCount.textContent = `현재 ${count}명 표시 중`;
    }

    function downloadVisibleCardPdf() {
      const visibleCards = getVisibleCards();
      if (visibleCards.length === 0) {
        alert("다운로드할 학생이 없습니다.");
        return;
      }

      if (visibleCards.length !== 1) {
        alert("한 명만 보이도록 검색한 뒤 다시 눌러 주세요. 전체 파일이 필요하면 '전체 PDF ZIP'을 사용하면 됩니다.");
        return;
      }

      const link = visibleCards[0].querySelector("[data-pdf-link]");
      if (link) {
        link.click();
      }
    }

    searchInput.addEventListener("input", applyFilter);
    resetButton.addEventListener("click", () => {
      searchInput.value = "";
      applyFilter();
      searchInput.focus();
    });
    downloadVisibleButton.addEventListener("click", downloadVisibleCardPdf);

    applyFilter();
  </script>
</body>
</html>
"""

    pages = "\n".join(render_page(student) for student in students)
    timestamp = html.escape(datetime.now().strftime("%Y-%m-%d %H:%M"))
    replacements = {
        "__COUNT__": str(len(students)),
        "__TIMESTAMP__": timestamp,
        "__PAGES__": pages,
        "__ZIP_HREF__": PDF_ZIP_REL_PATH.as_posix(),
        "__CHOICE_A__": CHOICE_COLORS["A"]["html"],
        "__CHOICE_B__": CHOICE_COLORS["B"]["html"],
        "__CHOICE_C__": CHOICE_COLORS["C"]["html"],
        "__CHOICE_D__": CHOICE_COLORS["D"]["html"],
        "__CHOICE_E__": CHOICE_COLORS["E"]["html"],
        "__CHOICE_F__": CHOICE_COLORS["F"]["html"],
        "__CHOICE_G__": CHOICE_COLORS["G"]["html"],
        "__CHOICE_H__": CHOICE_COLORS["H"]["html"],
    }
    for key, value in replacements.items():
        template = template.replace(key, value)
    return template


@lru_cache(maxsize=None)
def get_font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH, size=size)


def gradient_image(size: tuple[int, int], start: tuple[int, int, int], end: tuple[int, int, int], horizontal: bool) -> Image.Image:
    width, height = size
    image = Image.new("RGB", size, start)
    draw = ImageDraw.Draw(image)
    steps = width if horizontal else height
    if steps <= 1:
        return image

    for step in range(steps):
        ratio = step / (steps - 1)
        color = tuple(int(start[index] + (end[index] - start[index]) * ratio) for index in range(3))
        if horizontal:
            draw.line((step, 0, step, height), fill=color)
        else:
            draw.line((0, step, width, step), fill=color)
    return image


def rounded_mask(size: tuple[int, int], radius: int) -> Image.Image:
    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, size[0], size[1]), radius=radius, fill=255)
    return mask


def paste_rounded_fill(
    base: Image.Image,
    box: tuple[int, int, int, int],
    fill: tuple[int, int, int],
    radius: int,
    outline: tuple[int, int, int] | None = None,
    width: int = 1,
) -> None:
    draw = ImageDraw.Draw(base)
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def paste_rounded_gradient(
    base: Image.Image,
    box: tuple[int, int, int, int],
    start: tuple[int, int, int],
    end: tuple[int, int, int],
    radius: int,
    horizontal: bool = False,
    outline: tuple[int, int, int] | None = None,
    width: int = 1,
) -> None:
    x1, y1, x2, y2 = box
    gradient = gradient_image((x2 - x1, y2 - y1), start, end, horizontal).convert("RGBA")
    mask = rounded_mask((x2 - x1, y2 - y1), radius)
    base.alpha_composite(gradient, (x1, y1))
    if outline is not None:
        ImageDraw.Draw(base).rounded_rectangle(box, radius=radius, outline=outline, width=width)


def draw_shadow(base: Image.Image, box: tuple[int, int, int, int], radius: int) -> None:
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    shadow_box = (box[0], box[1] + 18, box[2], box[3] + 18)
    ImageDraw.Draw(shadow).rounded_rectangle(shadow_box, radius=radius, fill=(96, 73, 49, 52))
    shadow = shadow.filter(ImageFilter.GaussianBlur(28))
    base.alpha_composite(shadow)


def fit_multiline_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    box: tuple[int, int, int, int],
    max_size: int,
    min_size: int,
) -> tuple[ImageFont.FreeTypeFont, int]:
    max_width = box[2] - box[0] - 10
    max_height = box[3] - box[1] - 10
    for size in range(max_size, min_size - 1, -2):
        font = get_font(size)
        spacing = max(4, int(size * 0.18))
        bounds = draw.multiline_textbbox((0, 0), text, font=font, spacing=spacing, align="center")
        width = bounds[2] - bounds[0]
        height = bounds[3] - bounds[1]
        if width <= max_width and height <= max_height:
            return font, spacing
    return get_font(min_size), max(4, int(min_size * 0.18))


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    fill: tuple[int, int, int],
    max_size: int,
    min_size: int,
) -> None:
    font, spacing = fit_multiline_text(draw, text, box, max_size=max_size, min_size=min_size)
    bounds = draw.multiline_textbbox((0, 0), text, font=font, spacing=spacing, align="center")
    width = bounds[2] - bounds[0]
    height = bounds[3] - bounds[1]
    x = box[0] + ((box[2] - box[0]) - width) / 2
    y = box[1] + ((box[3] - box[1]) - height) / 2
    draw.multiline_text((x, y), text, font=font, fill=fill, spacing=spacing, align="center")


def subject_lines(subject: str) -> str:
    return "\n".join(subject.split("_"))


def draw_label_pill(
    draw: ImageDraw.ImageDraw,
    center_x: int,
    y: int,
    text: str,
    font_size: int,
    background: tuple[int, int, int],
    foreground: tuple[int, int, int],
) -> int:
    font = get_font(font_size)
    bounds = draw.textbbox((0, 0), text, font=font)
    width = bounds[2] - bounds[0] + 30
    height = bounds[3] - bounds[1] + 14
    x1 = center_x - width // 2
    y1 = y
    x2 = x1 + width
    y2 = y1 + height
    draw.rounded_rectangle((x1, y1, x2, y2), radius=height // 2, fill=background)
    draw.text((center_x, y1 + height // 2), text, font=font, fill=foreground, anchor="mm")
    return y2


def draw_choice_chip(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    label: str,
    subject: str,
    fill: tuple[int, int, int],
) -> None:
    paste_rounded_fill(image, box, fill, radius=26, outline=(217, 205, 194), width=2)
    x1, y1, x2, y2 = box
    label_font = get_font(18)
    subject_box = (x1 + 18, y1 + 24, x2 - 18, y2 - 10)
    value_font, spacing = fit_multiline_text(draw, subject_lines(subject), subject_box, 23, 15)
    draw.text((x1 + 18, y1 + 10), label, font=label_font, fill=(118, 106, 98))
    draw.multiline_text((x1 + 18, y1 + 28), subject_lines(subject), font=value_font, fill=(48, 42, 36), spacing=spacing, align="left")


def draw_slot_cell(
    image: Image.Image,
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    token: str,
    student: Student,
) -> None:
    if token in CHOICE_LABELS:
        fill = CHOICE_COLORS[token]["image"]
        badge_text = f"선택{token}"
        subject = student.choices[token]
        destination = destination_text(token, subject)
    elif token:
        fill = COMMON_COLORS.get(token, {"image": (244, 239, 233)})["image"]
        badge_text = "공통"
        subject = token
        destination = ""
    else:
        fill = (244, 241, 237)
        badge_text = ""
        subject = "-"
        destination = ""

    ImageDraw.Draw(image).rectangle(box, fill=fill, outline=(211, 200, 190), width=1)
    x1, y1, x2, y2 = box
    center_x = (x1 + x2) // 2

    if token:
        pill_bottom = draw_label_pill(draw, center_x, y1 + 18, badge_text, 18, (255, 251, 247), (116, 103, 95))
        subject_bottom = y2 - 28 if destination else y2 - 12
        draw_centered_text(draw, (x1 + 18, pill_bottom + 10, x2 - 18, subject_bottom), subject_lines(subject), (43, 36, 31), 34, 20)
        if destination:
            draw.text((center_x, y2 - 16), destination, font=get_font(15), fill=(118, 106, 98), anchor="mm")
    else:
        draw.text((center_x, (y1 + y2) // 2), "-", font=get_font(38), fill=(150, 140, 132), anchor="mm")


def build_student_pdf_image(student: Student) -> Image.Image:
    canvas = Image.new("RGBA", PDF_CANVAS, (0, 0, 0, 0))
    background = gradient_image(PDF_CANVAS, (248, 243, 236), (233, 237, 245), horizontal=True).convert("RGBA")
    canvas.alpha_composite(background)

    card = (18, 18, PDF_CANVAS[0] - 18, PDF_CANVAS[1] - 18)
    draw_shadow(canvas, card, radius=42)
    paste_rounded_fill(canvas, card, (255, 249, 241), radius=42, outline=(224, 214, 205), width=2)
    paste_rounded_gradient(canvas, (card[0], card[1], card[2], card[1] + 12), (206, 149, 102), (136, 168, 220), radius=42, horizontal=True)

    draw = ImageDraw.Draw(canvas)
    ink = (42, 37, 33)
    muted = (118, 106, 98)
    line = (214, 203, 193)

    left_x = 58
    top_y = 72

    draw.text((left_x, top_y), "2026학년도 1학기 3학년 2반 개인 시간표", font=get_font(28), fill=muted)
    draw.text((left_x, top_y + 52), student.name, font=get_font(74), fill=ink)
    draw.text((left_x, top_y + 136), f"학번 {student.student_no}", font=get_font(32), fill=muted)

    panel = (1230, 56, PDF_CANVAS[0] - 56, 430)
    paste_rounded_fill(canvas, panel, (255, 250, 244), radius=34, outline=line, width=2)
    draw.text((panel[0] + 26, panel[1] + 22), "선택 과목", font=get_font(32), fill=ink)

    chip_gap_x = 18
    chip_gap_y = 8
    chip_width = ((panel[2] - panel[0]) - 52 - chip_gap_x) // 2
    chip_height = 70
    chip_x1 = panel[0] + 26
    chip_x2 = chip_x1 + chip_width + chip_gap_x
    chip_start_y = panel[1] + 68

    for index, label in enumerate(CHOICE_LABELS):
        row = index // 2
        column = index % 2
        x1 = chip_x1 if column == 0 else chip_x2
        y1 = chip_start_y + row * (chip_height + chip_gap_y)
        x2 = x1 + chip_width
        y2 = y1 + chip_height
        draw_choice_chip(
            canvas,
            draw,
            (x1, y1, x2, y2),
            f"선택{label}",
            student.choices[label],
            CHOICE_COLORS[label]["image"],
        )

    table = (58, 470, PDF_CANVAS[0] - 58, PDF_CANVAS[1] - 56)
    paste_rounded_fill(canvas, table, (255, 252, 247), radius=34, outline=line, width=2)

    table_x1, table_y1, table_x2, _ = table
    period_width = 88
    time_width = 180
    total_day_width = (table_x2 - table_x1) - period_width - time_width
    day_width = total_day_width // 5
    day_widths = [day_width] * 5
    day_widths[-1] += total_day_width - day_width * 5

    columns = [table_x1, table_x1 + period_width, table_x1 + period_width + time_width]
    for width in day_widths:
        columns.append(columns[-1] + width)

    header_height = 58
    class_height = 128
    divider_height = 42
    current_y = table_y1

    header_fill = (246, 238, 228)
    draw.rectangle((columns[0], current_y, columns[1], current_y + header_height), fill=header_fill, outline=line, width=1)
    draw.rectangle((columns[1], current_y, columns[2], current_y + header_height), fill=header_fill, outline=line, width=1)
    draw_centered_text(draw, (columns[0], current_y, columns[1], current_y + header_height), "교시", ink, 28, 20)
    draw_centered_text(draw, (columns[1], current_y, columns[2], current_y + header_height), "시간", ink, 28, 20)
    for day_index, day in enumerate(DAYS):
        x1 = columns[2 + day_index]
        x2 = columns[3 + day_index]
        draw.rectangle((x1, current_y, x2, current_y + header_height), fill=header_fill, outline=line, width=1)
        draw_centered_text(draw, (x1, current_y, x2, current_y + header_height), day, ink, 30, 22)
    current_y += header_height

    for entry in BASE_TIMETABLE:
        if entry["type"] == "break":
            draw.rectangle((table_x1, current_y, table_x2, current_y + divider_height), fill=(248, 244, 238), outline=line, width=1)
            draw_centered_text(draw, (table_x1, current_y, table_x2, current_y + divider_height), entry["label"], (92, 79, 69), 24, 18)
            current_y += divider_height
            continue

        row_bottom = current_y + class_height
        draw.rectangle((columns[0], current_y, columns[1], row_bottom), fill=(255, 250, 242), outline=line, width=1)
        draw.rectangle((columns[1], current_y, columns[2], row_bottom), fill=(255, 252, 247), outline=line, width=1)
        draw_centered_text(draw, (columns[0], current_y, columns[1], row_bottom), entry["period"], ink, 32, 22)
        draw_centered_text(draw, (columns[1], current_y, columns[2], row_bottom), entry["time"], muted, 27, 18)

        for day_index, token in enumerate(entry["slots"]):
            x1 = columns[2 + day_index]
            x2 = columns[3 + day_index]
            draw_slot_cell(canvas, draw, (x1, current_y, x2, row_bottom), token, student)

        current_y = row_bottom

    return canvas.convert("RGB")


def generate_pdf_assets(output_path: Path, students: Iterable[Student]) -> tuple[Path, Path]:
    output_dir = output_path.parent
    assets_dir = output_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir = output_dir / PDF_REL_DIR
    pdf_dir.mkdir(parents=True, exist_ok=True)
    assets_shadow = assets_dir / "._pdf"
    if assets_shadow.exists():
        assets_shadow.unlink()

    for old_path in pdf_dir.iterdir():
        if old_path.is_file() and (old_path.suffix == ".pdf" or old_path.name.startswith("._")):
            old_path.unlink()

    students_list = list(students)
    generated_pdfs: list[tuple[Student, Path]] = []
    for student in students_list:
        pdf_path = pdf_dir / f"{student.student_no}.pdf"
        image = build_student_pdf_image(student)
        image.save(pdf_path, "PDF", resolution=200.0)
        apple_double = pdf_dir / f"._{pdf_path.name}"
        if apple_double.exists():
            apple_double.unlink()
        generated_pdfs.append((student, pdf_path))

    zip_path = output_dir / PDF_ZIP_REL_PATH
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for student, pdf_path in generated_pdfs:
            archive.write(pdf_path, arcname=pdf_download_name(student))
    apple_double = zip_path.parent / f"._{zip_path.name}"
    if apple_double.exists():
        apple_double.unlink()
    assets_shadow = assets_dir / "._pdf"
    if assets_shadow.exists():
        assets_shadow.unlink()

    return pdf_dir, zip_path


def write_outputs(output_path: Path, students: list[Student]) -> tuple[Path, Path, Path]:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data_output = output_path.with_suffix(".json")
    data_output.parent.mkdir(parents=True, exist_ok=True)

    html_document = build_document(students)
    output_path.write_text(html_document, encoding="utf-8")

    data_output.write_text(
        json.dumps(
            [
                {
                    "student_no": student.student_no,
                    "name": student.name,
                    "choices": student.choices,
                    "pdf": pdf_href(student),
                }
                for student in students
            ],
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    _, zip_path = generate_pdf_assets(output_path, students)
    return output_path, data_output, zip_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate student timetable HTML and PDF assets.")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Path to the source XLSX file.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Path to the output HTML file.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.source.exists():
        raise SystemExit(f"Source file not found: {args.source}")

    students = parse_students(args.source)
    html_path, json_path, zip_path = write_outputs(args.output, students)

    print(f"Generated {len(students)} student timetables.")
    print(f"HTML: {html_path.resolve()}")
    print(f"JSON: {json_path.resolve()}")
    print(f"PDF ZIP: {zip_path.resolve()}")


if __name__ == "__main__":
    main()
