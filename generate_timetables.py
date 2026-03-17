#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET


DEFAULT_SOURCE = Path("/Users/yangdongmoon/Downloads/3-2반 1학기 시간표.xlsx")
DEFAULT_OUTPUT = Path("dist/3-2-학생별-시간표.html")

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
        row_number = int(row.attrib["r"])
        if row_number == 1:
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
    pieces = [html.escape(piece) for piece in subject.split("_")]
    return "<br>".join(pieces)


def render_slot(token: str, student: Student) -> str:
    if token in CHOICE_LABELS:
        subject = student.choices[token]
        return (
            f'<td class="slot choice choice-{token.lower()}">'
            f'<div class="slot-badge">선택{token}</div>'
            f'<div class="{subject_length_class(subject)}">{format_subject(subject)}</div>'
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
        subject = student.choices[label]
        items.append(
            f'<li class="choice-item choice-{label.lower()}">'
            f'<span class="choice-key">선택{label}</span>'
            f'<span class="choice-name">{html.escape(subject)}</span>'
            "</li>"
        )
    return "\n".join(items)


def render_page(student: Student) -> str:
    return f"""
    <section class="student-page is-visible" data-name="{html.escape(student.name.lower())}" data-student-no="{html.escape(student.student_no)}">
      <div class="page-accent"></div>
      <header class="page-header">
        <div>
          <p class="eyebrow">2026학년도 1학기 3학년 2반 개인 시간표</p>
          <h2 class="student-name">{html.escape(student.name)}</h2>
          <p class="student-meta">학번 {html.escape(student.student_no)}</p>
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
    pages = "\n".join(render_page(student) for student in students)
    student_count = len(students)
    metadata = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "student_count": student_count,
    }
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>3학년 2반 학생별 시간표</title>
  <style>
    :root {{
      --paper: #fffdf8;
      --ink: #2a2521;
      --muted: #756a61;
      --line: #d8cfc7;
      --line-strong: #9f9185;
      --bg: #f5efe7;
      --card-shadow: 0 24px 70px rgba(80, 58, 37, 0.12);
      --choice-a: #f7dcc5;
      --choice-b: #c8c7e8;
      --choice-c: #d7d6d8;
      --choice-d: #c8eadf;
      --choice-e: #cfd9f3;
      --choice-f: #e0d4ea;
      --choice-g: #fff1b8;
      --choice-h: #eef0c7;
      --common-art: #f7efe5;
      --common-career: #f9efe6;
      --common-sports: #ece8e4;
      --common-club: #f4eadc;
      --empty: #f4f1ed;
      --accent: #ca8b57;
      --accent-2: #7caa94;
    }}

    * {{
      box-sizing: border-box;
    }}

    html {{
      scroll-behavior: smooth;
    }}

    body {{
      margin: 0;
      color: var(--ink);
      font-family: "Apple SD Gothic Neo", "Noto Sans KR", sans-serif;
      background:
        radial-gradient(circle at top left, rgba(202, 139, 87, 0.16), transparent 28%),
        radial-gradient(circle at top right, rgba(124, 170, 148, 0.18), transparent 24%),
        linear-gradient(180deg, #f8f3ec 0%, #f1ece6 100%);
    }}

    .app {{
      width: min(1520px, calc(100vw - 32px));
      margin: 24px auto 64px;
    }}

    .toolbar {{
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
    }}

    .toolbar-main {{
      display: flex;
      gap: 14px;
      align-items: center;
      flex-wrap: wrap;
    }}

    .toolbar h1 {{
      margin: 0;
      font-size: 1.2rem;
    }}

    .toolbar-meta {{
      margin: 2px 0 0;
      color: var(--muted);
      font-size: 0.92rem;
    }}

    .toolbar-actions {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }}

    .search-input {{
      min-width: 240px;
      padding: 12px 14px;
      border: 1px solid rgba(159, 145, 133, 0.55);
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.9);
      color: var(--ink);
      font-size: 0.96rem;
      outline: none;
    }}

    .search-input:focus {{
      border-color: var(--accent);
      box-shadow: 0 0 0 4px rgba(202, 139, 87, 0.14);
    }}

    .btn {{
      border: 0;
      border-radius: 999px;
      padding: 12px 16px;
      cursor: pointer;
      font-size: 0.94rem;
      font-weight: 700;
      color: white;
      background: linear-gradient(135deg, #b56f3d 0%, #ce9566 100%);
      box-shadow: 0 10px 24px rgba(181, 111, 61, 0.25);
    }}

    .btn.secondary {{
      color: var(--ink);
      background: white;
      border: 1px solid rgba(159, 145, 133, 0.42);
      box-shadow: none;
    }}

    .status-pill {{
      padding: 10px 14px;
      border-radius: 999px;
      background: rgba(124, 170, 148, 0.12);
      color: #456a56;
      font-size: 0.9rem;
      font-weight: 700;
    }}

    .pages {{
      display: flex;
      flex-direction: column;
      gap: 26px;
    }}

    .student-page {{
      position: relative;
      overflow: hidden;
      padding: 26px 28px 28px;
      border: 1px solid rgba(159, 145, 133, 0.32);
      border-radius: 32px;
      background:
        linear-gradient(180deg, rgba(255, 247, 236, 0.96) 0%, rgba(255, 255, 255, 0.96) 100%),
        var(--paper);
      box-shadow: var(--card-shadow);
      break-after: page;
    }}

    .student-page.is-hidden {{
      display: none;
    }}

    .page-accent {{
      position: absolute;
      inset: 0 auto auto 0;
      width: 100%;
      height: 12px;
      background: linear-gradient(90deg, #ce9566 0%, #e1bc75 34%, #8db69c 68%, #88a8dc 100%);
    }}

    .page-header {{
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(320px, 0.9fr);
      gap: 18px;
      align-items: start;
      margin-bottom: 20px;
      padding-top: 12px;
    }}

    .eyebrow {{
      margin: 0 0 10px;
      color: var(--muted);
      font-size: 0.95rem;
      font-weight: 700;
      letter-spacing: 0.02em;
    }}

    .student-name {{
      margin: 0;
      font-size: clamp(2.1rem, 3.3vw, 3.1rem);
      line-height: 1;
    }}

    .student-meta {{
      margin: 10px 0 0;
      font-size: 1.08rem;
      color: var(--muted);
      font-weight: 700;
    }}

    .choice-panel {{
      padding: 18px;
      border: 1px solid rgba(159, 145, 133, 0.25);
      border-radius: 24px;
      background:
        linear-gradient(145deg, rgba(255, 251, 244, 0.95), rgba(252, 246, 237, 0.88));
    }}

    .panel-title {{
      margin: 0 0 12px;
      font-size: 0.98rem;
      font-weight: 800;
    }}

    .choice-list {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
      margin: 0;
      padding: 0;
      list-style: none;
    }}

    .choice-item {{
      min-width: 0;
      padding: 10px 12px;
      border-radius: 18px;
      border: 1px solid rgba(159, 145, 133, 0.18);
    }}

    .choice-key,
    .choice-name {{
      display: block;
    }}

    .choice-key {{
      margin-bottom: 4px;
      font-size: 0.82rem;
      color: var(--muted);
      font-weight: 800;
    }}

    .choice-name {{
      font-size: 0.92rem;
      font-weight: 700;
      line-height: 1.3;
      word-break: keep-all;
    }}

    .choice-a {{ background: rgba(247, 220, 197, 0.82); }}
    .choice-b {{ background: rgba(200, 199, 232, 0.72); }}
    .choice-c {{ background: rgba(215, 214, 216, 0.74); }}
    .choice-d {{ background: rgba(200, 234, 223, 0.82); }}
    .choice-e {{ background: rgba(207, 217, 243, 0.78); }}
    .choice-f {{ background: rgba(224, 212, 234, 0.82); }}
    .choice-g {{ background: rgba(255, 241, 184, 0.78); }}
    .choice-h {{ background: rgba(238, 240, 199, 0.88); }}

    .timetable-shell {{
      overflow: hidden;
      border: 1px solid rgba(159, 145, 133, 0.32);
      border-radius: 24px;
      background: rgba(255, 255, 255, 0.8);
    }}

    .timetable {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }}

    .timetable th,
    .timetable td {{
      border: 1px solid rgba(159, 145, 133, 0.44);
      text-align: center;
      padding: 10px 8px;
      vertical-align: middle;
    }}

    .timetable thead th {{
      padding: 12px 8px;
      background: #f6eee4;
      font-size: 0.97rem;
      font-weight: 800;
    }}

    .period-head,
    .period {{
      width: 58px;
      background: #fffaf2;
      font-weight: 800;
    }}

    .time-head,
    .time {{
      width: 120px;
      background: #fffcf7;
      color: var(--muted);
      font-weight: 700;
      font-size: 0.9rem;
    }}

    .slot {{
      height: 108px;
      padding: 8px;
    }}

    .slot-badge {{
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
    }}

    .slot-badge-common {{
      background: rgba(255, 255, 255, 0.88);
    }}

    .subject {{
      font-weight: 800;
      line-height: 1.28;
      word-break: keep-all;
      overflow-wrap: anywhere;
    }}

    .subject.tight {{
      font-size: 0.95rem;
    }}

    .subject.tighter {{
      font-size: 0.88rem;
    }}

    .subject.tightest {{
      font-size: 0.8rem;
    }}

    .choice.choice-a {{ background: var(--choice-a); }}
    .choice.choice-b {{ background: var(--choice-b); }}
    .choice.choice-c {{ background: var(--choice-c); }}
    .choice.choice-d {{ background: var(--choice-d); }}
    .choice.choice-e {{ background: var(--choice-e); }}
    .choice.choice-f {{ background: var(--choice-f); }}
    .choice.choice-g {{ background: var(--choice-g); }}
    .choice.choice-h {{ background: var(--choice-h); }}

    .common-art {{ background: var(--common-art); }}
    .common-career {{ background: var(--common-career); }}
    .common-sports {{ background: var(--common-sports); }}
    .common-club {{ background: var(--common-club); }}

    .empty {{
      background: var(--empty);
    }}

    .empty-mark {{
      color: rgba(117, 106, 97, 0.72);
      font-size: 1.2rem;
      font-weight: 700;
    }}

    .divider-cell {{
      padding: 8px 12px;
      background:
        linear-gradient(90deg, rgba(202, 139, 87, 0.11), rgba(124, 170, 148, 0.11));
      color: #5b4e45;
      font-size: 0.88rem;
      font-weight: 800;
      letter-spacing: 0.06em;
    }}

    @media (max-width: 1080px) {{
      .app {{
        width: min(100vw - 18px, 1520px);
        margin: 12px auto 36px;
      }}

      .toolbar {{
        position: static;
      }}

      .page-header {{
        grid-template-columns: 1fr;
      }}

      .choice-list {{
        grid-template-columns: 1fr;
      }}

      .timetable-shell {{
        overflow-x: auto;
      }}

      .timetable {{
        min-width: 860px;
      }}
    }}

    @media print {{
      @page {{
        size: A4 landscape;
        margin: 9mm;
      }}

      body {{
        background: white;
      }}

      .app {{
        width: auto;
        margin: 0;
      }}

      .toolbar {{
        display: none;
      }}

      .pages {{
        gap: 0;
      }}

      .student-page {{
        min-height: calc(100vh - 2mm);
        margin: 0;
        padding: 18px 18px 16px;
        border: 1px solid rgba(159, 145, 133, 0.4);
        border-radius: 18px;
        box-shadow: none;
      }}

      .student-page:last-child {{
        break-after: auto;
      }}

      .slot {{
        height: 92px;
      }}
    }}
  </style>
</head>
<body>
  <main class="app">
    <section class="toolbar">
      <div class="toolbar-main">
        <div>
          <h1>3학년 2반 학생별 시간표</h1>
          <p class="toolbar-meta">총 {student_count}명 · 생성 시각 {html.escape(timestamp)} · 검색 후 보이는 카드만 인쇄됩니다.</p>
        </div>
        <span class="status-pill" id="visibleCount">현재 {student_count}명 표시 중</span>
      </div>
      <div class="toolbar-actions">
        <input id="searchInput" class="search-input" type="search" placeholder="이름 또는 학번 검색">
        <button id="resetButton" class="btn secondary" type="button">전체 보기</button>
        <button id="printButton" class="btn" type="button">보이는 카드 인쇄</button>
      </div>
    </section>

    <section class="pages" id="pages">
      {pages}
    </section>
  </main>

  <script id="dataset" type="application/json">{json.dumps(metadata, ensure_ascii=False)}</script>
  <script>
    const searchInput = document.getElementById("searchInput");
    const resetButton = document.getElementById("resetButton");
    const printButton = document.getElementById("printButton");
    const visibleCount = document.getElementById("visibleCount");
    const cards = Array.from(document.querySelectorAll(".student-page"));

    function normalize(value) {{
      return value.toLowerCase().trim();
    }}

    function applyFilter() {{
      const query = normalize(searchInput.value);
      let count = 0;

      cards.forEach((card) => {{
        const haystack = normalize(card.dataset.name) + " " + normalize(card.dataset.studentNo);
        const matched = query === "" || haystack.includes(query);
        card.classList.toggle("is-hidden", !matched);
        card.classList.toggle("is-visible", matched);
        if (matched) count += 1;
      }});

      visibleCount.textContent = `현재 ${{count}}명 표시 중`;
    }}

    searchInput.addEventListener("input", applyFilter);
    resetButton.addEventListener("click", () => {{
      searchInput.value = "";
      applyFilter();
      searchInput.focus();
    }});

    printButton.addEventListener("click", () => {{
      window.print();
    }});

    applyFilter();
  </script>
</body>
</html>
"""


def write_outputs(output_path: Path, students: list[Student]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    data_output = output_path.with_suffix(".json")
    data_output.parent.mkdir(parents=True, exist_ok=True)
    rendered = build_document(students)
    output_path.write_text(rendered, encoding="utf-8")

    payload = [
        {
            "student_no": student.student_no,
            "name": student.name,
            "choices": student.choices,
        }
        for student in students
    ]
    data_output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate printable student timetable HTML.")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE, help="Path to the source XLSX file.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Path to the output HTML file.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.source.exists():
        raise SystemExit(f"Source file not found: {args.source}")

    students = parse_students(args.source)
    write_outputs(args.output, students)

    print(f"Generated {len(students)} student timetables.")
    print(f"HTML: {args.output.resolve()}")
    print(f"JSON: {args.output.with_suffix('.json').resolve()}")


if __name__ == "__main__":
    main()
