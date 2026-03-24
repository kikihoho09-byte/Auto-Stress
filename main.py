"""
Excel Part Number Matcher
- 엑셀1 파일을 드래그앤드롭으로 불러옴
- Database1.xlsx와 Part Number 비교
- 일치하는 spec 값을 엑셀1의 특정 위치에 기입
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import sys
import shutil
import threading
import queue
from datetime import datetime
import re
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ─────────────────────────────────────────────
# 경로 설정
# ─────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR   = os.path.join(BASE_DIR, "DB")
DB_FILE  = os.path.join(DB_DIR, "Database.xlsx")

# ─────────────────────────────────────────────
# Database 로드 (ZIP/XML 직접 파싱 - 속도 및 정확도 최대화)
# ─────────────────────────────────────────────
def load_database():
    import zipfile
    import xml.etree.ElementTree as ET

    if not os.path.exists(DB_FILE):
        raise FileNotFoundError(f"Database 파일을 찾을 수 없습니다:\n{DB_FILE}")

    records = []
    
    def find_cell_index(cell_ref):
        col_str = "".join(filter(str.isalpha, cell_ref))
        idx = 0
        for char in col_str:
            idx = idx * 26 + (ord(char.upper()) - ord('A') + 1)
        return idx - 1

    try:
        with zipfile.ZipFile(DB_FILE, 'r') as zf:
            # 1. 공통 문자열 로드
            strings = []
            try:
                with zf.open('xl/sharedStrings.xml') as f:
                    root = ET.parse(f).getroot()
                    ns = {'m': root.tag.split('}')[0].strip('{')}
                    for t in root.findall('.//m:t', ns):
                        strings.append(t.text if t.text else "")
            except: pass

            # 2. 시트 리스트 로드
            sheet_paths = {}
            rid_map = {}
            try:
                with zf.open('xl/_rels/workbook.xml.rels') as f:
                    root = ET.parse(f).getroot()
                    ns = {'m': root.tag.split('}')[0].strip('{')}
                    for r in root.findall('.//m:Relationship', ns):
                        rid_map[r.attrib.get('Id')] = r.attrib.get('Target')
                
                with zf.open('xl/workbook.xml') as f:
                    root = ET.parse(f).getroot()
                    ns = {'m': root.tag.split('}')[0].strip('{')}
                    for s in root.findall('.//m:sheet', ns):
                        name = s.attrib.get('name', '').upper().strip()
                        rid = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        path = rid_map.get(rid)
                        if path:
                            if not path.startswith('worksheets/'): path = f"worksheets/{os.path.basename(path)}"
                            sheet_paths[name] = f"xl/{path}"
            except: pass

            # 모든 시트: C열(index 2)이 부품번호, D열(index 3)부터가 스펙
            categories = {
                "IC":     ["IC", 1],      # D
                "MOSFET": ["MOSFET", 3],  # D, E, F
                "DIODE":  ["RECTIFIER(DIODE)", 3], # D, E, F
                "CAP":    ["CAP", 2],     # D, E
                "TR":     ["TR", 4]       # D, E, F, G
            }

            for key, (cat_name, spec_count) in categories.items():
                xpath = None
                for s_name in sheet_paths:
                    if key in s_name:
                        xpath = sheet_paths[s_name]
                        break
                
                if not xpath: continue
                
                try:
                    with zf.open(xpath) as f:
                        root = ET.parse(f).getroot()
                        ns = {'m': root.tag.split('}')[0].strip('{')}
                        count = 0
                        for row_node in root.findall('.//m:row', ns):
                            r_idx = int(row_node.attrib.get('r', 0))
                            if r_idx < 3: continue 
                            
                            row_vals = {}
                            for c in row_node.findall('m:c', ns):
                                ref = c.attrib.get('r')
                                idx = find_cell_index(ref)
                                v = c.find('m:v', ns)
                                val = v.text if v is not None else ""
                                if c.attrib.get('t') == 's' and val.isdigit():
                                    try: val = strings[int(val)]
                                    except: pass
                                row_vals[idx] = val
                            
                            # C열(index 2)이 부품번호!!
                            pn = row_vals.get(2)
                            if pn and str(pn).strip() and str(pn).strip().upper() not in ["IC", "MOSFET", "DIODE", "CAP", "TR", "PART NUMBER"]:
                                # D열(index 3)부터 스펙 시작
                                specs = [str(row_vals.get(i, "")).strip() for i in range(3, 3 + spec_count)]
                                records.append({
                                    "category": cat_name,
                                    "part_number": str(pn).strip(),
                                    "specs": specs
                                })
                                count += 1
                        print(f"[DB 로드] {key} 연관 시트: {count}개 로드 완료")
                except: pass

    except Exception as e:
        print(f"[오류] DB ZIP 파싱 중 실패: {e}")

    return records


# ─────────────────────────────────────────────
# Spec 숫자 파싱
# ─────────────────────────────────────────────
def parse_spec_numbers(spec_str):
    clean = str(spec_str).replace('\n', '/').replace('\r', '')
    parts = clean.split('/')
    numbers = []
    
    pattern = re.compile(r'(\d+(?:\.\d+)?)\s*([a-zA-Z]+)?')
    
    for part in parts:
        m = pattern.search(part.strip())
        if m:
            val = float(m.group(1))
            unit = m.group(2)
            if unit:
                unit = unit.strip().lower()
                if unit == 'mv':
                    val = val / 1000.0
                    unit = 'v'
                elif unit == 'v':
                    unit = 'v'
                elif unit == 'ma':
                    val = val / 1000.0
                    unit = 'a'
                elif unit == 'a':
                    unit = 'a'
                else:
                    pass
            numbers.append((val, unit))
    return numbers


# ─────────────────────────────────────────────
# 측정값 색상 처리
# ─────────────────────────────────────────────
def apply_measurement_color(ws, match_row, category, spec_str, log_func):
    meas_row   = match_row + 6
    col_start  = 28
    col_end    = 91

    spec_limits = parse_spec_numbers(spec_str)
    if not spec_limits:
        log_func(f"  [색상 건너뜀] spec 숫자 없음: {spec_str}")
        return

    log_func(f"  [색상 검사] 행{meas_row} {get_column_letter(col_start)}:{get_column_letter(col_end)} "
             f"| spec 한계(정규화)={spec_limits}")

    red_count = 0
    font_red   = InlineFont(color='FF0000')
    font_black = InlineFont(color='000000')

    pattern = re.compile(r'(\d+(?:\.\d+)?)\s*([a-zA-Z]+)?')

    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=meas_row, column=col)
        val_str = str(cell.value) if cell.value is not None else ""
        if not val_str.strip() or val_str == "None":
            continue

        parts = val_str.split('/')
        has_red = False
        rich_blocks = []

        for i, part in enumerate(parts):
            m = pattern.search(part)
            font_to_use = font_black
            
            if m:
                meas_val = float(m.group(1))
                unit = m.group(2)
                
                if unit:
                    unit = unit.strip().lower()
                    if unit == 'mv':
                        meas_val = meas_val / 1000.0
                    elif unit == 'ma':
                        meas_val = meas_val / 1000.0
                
                limit_val, _ = spec_limits[i % len(spec_limits)]
                
                if meas_val > limit_val:
                    font_to_use = font_red
                    has_red = True
            
            rich_blocks.append(TextBlock(font_to_use, part))
            
            if i < len(parts) - 1:
                rich_blocks.append(TextBlock(font_black, "/"))

        if has_red:
            cell.value = CellRichText(*rich_blocks)
            red_count += 1
            log_func(f"  [빨간색] {get_column_letter(col)}{meas_row} '{val_str}' -> 부분 빨간색 적용", color="red")

    if red_count:
        log_func(f"  -> {red_count}개 셀 빨간색 처리 완료", color="yellow")
    else:
        log_func(f"  -> 한계 초과 없음")


# ─────────────────────────────────────────────
# 엑셀1 처리
# ─────────────────────────────────────────────
def backup_excel(excel_path, log_func):
    backup_dir = os.path.join(BASE_DIR, "backup")
    os.makedirs(backup_dir, exist_ok=True)

    fname      = os.path.basename(excel_path)
    name, ext  = os.path.splitext(fname)
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{name}_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)

    shutil.copy2(excel_path, backup_path)
    log_func(f"[백업] backup\\{backup_name} 저장 완료")
    return backup_path


def process_excel(excel_path, db_records, log_func):
    backup_excel(excel_path, log_func)

    wb = load_workbook(excel_path)
    ws = wb.active

    matched           = 0
    checked           = 0
    unmatched         = []
    row_step          = 7
    row               = 14

    def find_real_max_row(ws, start_row=14):
        # 엑셀의 가짜 max_row(100만 행) 문제를 해결하기 위해 
        # 상단부터 10000행까지만 데이터가 있는지 빠르게 확인
        last_found = start_row
        max_to_check = min(ws.max_row, 10000) # 현실적으로 10000행 이상 데이터가 있지는 않음
        
        # iter_rows가 ws.cell()보다 수천배 빠릅니다.
        for r_idx, row_cells in enumerate(ws.iter_rows(min_row=start_row, max_row=max_to_check, min_col=13, max_col=18, values_only=True), start=start_row):
            if any(cell for cell in row_cells):
                last_found = r_idx
        return last_found

    print(f"[알림] 엑셀 파일 분석 중... 잠시만 기다려 주세요.")
    real_max = find_real_max_row(ws)
    log_func(f"[진행] 실제 데이터 탐지 완료: 행 {real_max}까지 처리를 시작합니다.")

    while row <= real_max:
        # 진행 상황 표시 (20행 단위)
        if (row - 14) % 21 == 0:
            print(f"[진행] 행 {row} 처리 중...")
        
        # 디버그용 DB 레코드 덤프 (최초 1회만)
        if row == 14:
            try:
                with open("db_debug_dump.txt", "w", encoding="utf-8") as f:
                    for r in db_records:
                        f.write(f"{r}\n")
                log_func("[시스템] DB 덤프 완료 (db_debug_dump.txt 확인 가능)")
            except: pass

        m_values = []
        for col_idx in range(13, 19): # M, N, O, P, Q, R (13~18)
            val = ws.cell(row=row, column=col_idx).value
            if val: m_values.append(str(val).strip())

        if not m_values:
            row += 1
            continue

        db_match = None
        current_m_str = " / ".join(m_values) 
        
        # 디버그: 현재 행에서 찾은 문자열들 출력
        # log_func(f"[디버그] 행{row} 검색어: {m_values}")

        for rec in db_records:
            pn = rec["part_number"]
            if not pn: continue
            
            match_found = False
            for m_str in m_values:
                # 1. 길이 4자 미만은 정확히 일치해야 함 (예: 'IC'가 'ICM801'에 매칭되는 것 방지)
                if len(pn) < 4:
                    if pn.lower() == m_str.lower():
                        match_found = True
                        break
                    # 셀 내부 단어와 정확히 일치하는지 확인
                    words = m_str.replace('\n', ' ').replace('\r', ' ').replace('/', ' ').split()
                    if any(pn.lower() == w.strip().lower() for w in words):
                        match_found = True
                        break
                else:
                    # 4자 이상은 기존처럼 포함 여부 확인
                    if pn.lower() in m_str.lower():
                        match_found = True
                        break
                    
                    # 줄바꿈이나 공백, 특수기호로 쪼개서 개별 단어와 일치하는지 확인
                    clean_str = m_str.replace('\n', ' ').replace('\r', ' ').replace('/', ' ').replace('(', ' ').replace(')', ' ')
                    words = clean_str.split()
                    if any(pn.lower() == w.strip().lower() for w in words):
                        match_found = True
                        break
            
            if match_found:
                db_match = rec
                break

        if db_match:
            checked += 1
            specs = db_match["specs"]
            valid_specs = [str(s).strip() for s in specs if s is not None and str(s).strip() != ""]
            combined_spec = " / ".join(valid_specs)
            
            # V열(22번)에 입력
            ws.cell(row=row, column=22).value = combined_spec
            
            log_func(
                f"[{db_match['category']}] 행{row} 매칭성공: '{db_match['part_number']}' "
                f"-> 입력스펙: {combined_spec}", color="blue"
            )
            matched += 1
            apply_measurement_color(ws, row, db_match["category"], combined_spec, log_func)
            row += row_step
        else:
            # 특정 부품(예: MBRF2080CTP)이 안 나오는 이유를 찾기 위한 임시 디버그
            if any("MBRF2080CTP" in s.upper() for s in m_values):
                log_func(f"[디버그] 행{row}에서 MBRF2080CTP 발견했으나 DB 매칭 실패!!", color="yellow")
            
            unmatched.append((row, current_m_str))
            row += 1

    result_dir  = os.path.join(BASE_DIR, "Result")
    os.makedirs(result_dir, exist_ok=True)

    fname       = os.path.basename(excel_path)
    name, ext   = os.path.splitext(fname)
    timestamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_name = f"{name}_result_{timestamp}{ext}"
    result_path = os.path.join(result_dir, result_name)

    wb.save(result_path)
    wb.close()
    log_func(f"[저장] Result\\{result_name}")
    
    unmatched_count = len(unmatched)
    log_func(f"\n완료: 총 {checked}개 확인, {matched}개 일치, {unmatched_count}개 미일치")
    
    return matched, checked, result_path, unmatched

# ─────────────────────────────────────────────
# v0.dev 스타일 모던 UI 클래스 컨ポー넌트
# ─────────────────────────────────────────────

BG_MAIN = "#09090b"
BG_CARD = "#18181b"
BORDER  = "#27272a"
PRIMARY = "#2dd4bf"
PRIMARY_HOV = "#14b8a6"
TEXT_PRI = "#f4f4f5"
TEXT_SEC = "#a1a1aa"
SUCCESS = "#22c55e"
ERROR_COL = "#ef4444"
BTN_DARK = "#27272a"
BTN_DARK_HOV = "#3f3f46"
WARN_BG = "#450a0a"
WARN_HOV = "#7f1d1d"

def create_rounded_rect(canvas, x1, y1, x2, y2, r, **kwargs):
    return canvas.create_polygon(
        x1+r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2, x2-r, y2, 
        x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1,
        smooth=True, **kwargs
    )

class ModernRoundedButton(tk.Canvas):
    def __init__(self, parent, text, command=None, height=45, radius=10, 
                 bg_color=PRIMARY, hover_color=PRIMARY_HOV, text_color="#000000", icon=None):
        super().__init__(parent, height=height, bg=parent["bg"], highlightthickness=0)
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.radius = radius
        self.state = "normal"
        self.text_val = text
        self.text_color = text_color
        self.icon = icon

        self.bind("<Configure>", self._on_resize)
        self.bind("<Button-1>", self._on_click)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_resize(self, event):
        self._draw(self.bg_color if self.state == "normal" else BORDER)
        
    def _draw(self, color):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10: return
        create_rounded_rect(self, 1, 1, w-1, h-1, self.radius, fill=color, outline="")
        
        display_text = f"{self.icon}   {self.text_val}" if self.icon else self.text_val
        self.text_id = self.create_text(w/2, h/2, text=display_text, fill=self.text_color if self.state == "normal" else TEXT_SEC, font=("Segoe UI", 11, "bold"), justify="center")

    def _on_enter(self, e):
        if self.state == "normal": self._draw(self.hover_color)
    def _on_leave(self, e):
        if self.state == "normal": self._draw(self.bg_color)
    def _on_click(self, e): pass
    def _on_release(self, e):
        if self.state == "normal" and self.command:
            self._draw(self.bg_color)
            self.command()

    def configure(self, **kwargs):
        if "state" in kwargs:
            self.state = kwargs["state"]
            self._draw(self.bg_color if self.state == "normal" else BORDER)

class DbCard(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent, height=80, bg=parent["bg"], highlightthickness=0)
        self.bind("<Configure>", self._on_resize)
        self.status = "⏳ 대기 중..."
        self.desc = "데이터 대기 중"
        self.color = TEXT_SEC
        
    def _on_resize(self, e): self._draw()
        
    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        create_rounded_rect(self, 2, 2, w-2, h-2, 10, fill=BG_CARD, outline=BORDER, width=1.5)
        self.create_oval(20, h/2-14, 48, h/2+14, fill="#042f2e", outline=PRIMARY, width=1.5)
        self.create_text(34, h/2, text="🗄", fill=PRIMARY, font=("Segoe UI", 12))
        self.create_text(65, h/2-10, text="Database", fill=TEXT_PRI, font=("Segoe UI", 11, "bold"), anchor="w")
        self.create_text(145, h/2-10, text=self.status, fill=self.color, font=("Segoe UI", 9), anchor="w")
        self.create_text(65, h/2+10, text=self.desc, fill=TEXT_SEC, font=("Segoe UI", 9), anchor="w")

    def set_loaded(self, count):
        self.status = "✅ 로드 완료"
        self.desc = f"{count}개 항목이 로드되었습니다"
        self.color = PRIMARY
        self._draw()
        
    def set_error(self, err):
        self.status = "❌ 로드 실패"
        self.desc = str(err)
        self.color = ERROR_COL
        self._draw()

class DashedUploadDropZone(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent, bg=parent["bg"], highlightthickness=0)
        self.bind("<Configure>", self._on_resize)
        self.state_text = "엑셀 파일을 드래그하여 업로드"
        self.sub_text = "또는 아래 부분을 클릭하세요"
        self.is_selected = False
        self.fname = ""
        
    def _on_resize(self, e): self._draw()

    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        if self.is_selected:
            create_rounded_rect(self, 2, 2, w-2, h-2, 12, fill="#042f2e", outline=PRIMARY, width=2)
            self.create_text(w/2, h/2 - 15, text="✅ 파일 선택 완료", fill=PRIMARY, font=("Segoe UI", 14, "bold"))
            self.create_text(w/2, h/2 + 20, text=self.fname, fill=TEXT_PRI, font=("Segoe UI", 11))
        else:
            create_rounded_rect(self, 2, 2, w-2, h-2, 12, fill=BG_CARD, outline=BORDER, dash=(7,7), width=2)
            self.create_rectangle(w/2-25, h/2-40, w/2+25, h/2-4, fill=BG_CARD, outline=BORDER, width=2)
            self.create_line(w/2, h/2-30, w/2, h/2-10, fill=TEXT_SEC, width=2)
            self.create_line(w/2-8, h/2-22, w/2, h/2-30, w/2+8, h/2-22, fill=TEXT_SEC, width=2)
            self.create_text(w/2, h/2 + 20, text=self.state_text, fill=TEXT_PRI, font=("Segoe UI", 12, "bold"))
            self.create_text(w/2, h/2 + 45, text=self.sub_text, fill=TEXT_SEC, font=("Segoe UI", 10))

    def set_selected(self, fname):
        self.is_selected = True
        self.fname = fname
        self._draw()

class LogCard(tk.Canvas):
    def __init__(self, parent, raw_msg, color_type):
        super().__init__(parent, height=65, bg=BG_MAIN, highlightthickness=0)
        self.raw_msg = raw_msg
        self.tag = "INF"
        self.content = raw_msg
        
        m = re.match(r'^\[(.*?)\] (.*)', raw_msg)
        if m:
            self.tag = m.group(1).upper()
            self.content = m.group(2)

        self.accent_color = BORDER
        self.label_color = TEXT_SEC
        
        if color_type == "red" or "오류" in self.tag or "실패" in self.tag:
            self.accent_color = ERROR_COL
            self.label_color = ERROR_COL
        elif color_type == "yellow" or "미일치" in self.tag:
            self.accent_color = "#eab308"
            self.label_color = "#eab308"
        elif "CAP" in self.tag or "IC" in self.tag or "MOSFET" in self.tag or "DIODE" in self.tag or "RECT" in self.tag or "일치" in self.tag or color_type == "green":
            self.accent_color = PRIMARY
            self.label_color = PRIMARY

        self.bind("<Configure>", self._on_resize)
        
    def _on_resize(self, e):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        create_rounded_rect(self, 2, 2, w-2, h-2, 8, fill=BG_CARD, outline=BORDER, width=1.5)
        self.create_rectangle(2, 10, 6, h-10, fill=self.accent_color, outline="")
        self.create_text(20, 22, text=self.tag, fill=self.label_color, font=("Segoe UI", 9, "bold"), anchor="w")
        self.create_text(20, 42, text=self.content[:95] + ('...' if len(self.content)>95 else ''), fill=TEXT_PRI, font=("Consolas", 10), anchor="w")

class ScrollableLogFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG_MAIN)
        self.canvas = tk.Canvas(self, bg=BG_MAIN, highlightthickness=0)
        style = ttk.Style()
        style.configure("Dark.Vertical.TScrollbar", background=BG_CARD, troughcolor=BG_MAIN, bordercolor=BG_MAIN, arrowcolor=PRIMARY)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview, style="Dark.Vertical.TScrollbar")
        self.scrollable_frame = tk.Frame(self.canvas, bg=BG_MAIN)
        self.window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y", padx=(5,0))
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.window_id, width=e.width))
        self.log_cards = []

    def add_log(self, text, color_type=None):
        card = LogCard(self.scrollable_frame, text, color_type)
        card.pack(fill="x", padx=2, pady=(0, 8))
        self.log_cards.append(card)
        if len(self.log_cards) > 500:
            old = self.log_cards.pop(0)
            old.destroy()
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1)

        return len(self.log_cards)

class FlatStatusBar(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent, height=45, bg="#000000", highlightthickness=0)
        self.bind("<Configure>", self._on_resize)
        self.indicator_color = TEXT_SEC
        self.status_text = "준비"
        self.pos = 0
        self.is_running = False
        self.bar_width = 150
        
    def _on_resize(self, e): self._draw()
        
    def _draw(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10: return
        self.create_line(0, 0, w, 0, fill=BORDER)
        self.indicator_id = self.create_oval(30, h/2-4, 38, h/2+4, fill=self.indicator_color, outline="")
        self.text_id = self.create_text(50, h/2, text=self.status_text, fill=TEXT_SEC, font=("Segoe UI", 10), anchor="w")
        self.create_text(w-30, h/2, text="Auto Spec Program v2.0", fill="#52525b", font=("Segoe UI", 9, "bold"), anchor="e")
        
        if self.is_running:
            track_w = 200
            track_x = w/2 - track_w/2
            create_rounded_rect(self, track_x, h/2-3, track_x+track_w, h/2+3, 3, fill=BG_CARD, outline=BORDER)
            bar_x = track_x + self.pos
            create_rounded_rect(self, bar_x, h/2-3, min(bar_x+self.bar_width, track_x+track_w), h/2+3, 3, fill=PRIMARY, outline="")

    def set_status(self, msg, color=TEXT_SEC):
        self.status_text = msg
        self.indicator_color = color
        self._draw()

    def start_progress(self):
        self.is_running = True
        self.pos = 0
        self.set_status("처리 중...", "#eab308")
        self._animate()

    def _animate(self):
        if not self.is_running: return
        self.pos += 4
        if self.pos > 200: self.pos = -self.bar_width
        self._draw()
        self.after(20, self._animate)
        
    def stop_progress(self, msg, success=True):
        self.is_running = False
        self.set_status(msg, SUCCESS if success else ERROR_COL)

# ─────────────────────────────────────────────
# 메인 어플리케이션
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto Spec Program")
        self.state('zoomed')
        self.minsize(1024, 768)
        self.configure(bg=BG_MAIN)

        self.log_queue = queue.Queue()
        self.db_records = []
        self.excel_path = None
        self.last_unmatched = []

        self._build_ui()
        self._setup_dnd()
        self._process_log_queue()
        self._load_db_on_start()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=4)
        self.grid_columnconfigure(1, weight=6)
        self.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self, bg=BG_MAIN)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=40, pady=(25, 10))
        
        logo_loaded = False
        if HAS_PIL:
            logo_path = os.path.join(BASE_DIR, "image", "logo.ico")
            if os.path.exists(logo_path):
                try:
                    img = Image.open(logo_path)
                    w, h = img.size
                    new_h = 42
                    new_w = int(w * (new_h / h))
                    try:
                        resample = Image.Resampling.LANCZOS
                    except AttributeError:
                        resample = Image.LANCZOS
                    img = img.resize((new_w, new_h), resample)
                    self.logo_img = ImageTk.PhotoImage(img) # 참조 유지
                    logo_lbl = tk.Label(header, image=self.logo_img, bg=BG_MAIN, bd=0)
                    logo_lbl.pack(side="left", padx=(0, 15))
                    logo_loaded = True
                except Exception:
                    pass

        if not logo_loaded:
            logo_lbl = tk.Label(header, text="⚡", fg=PRIMARY, bg="#042f2e", font=("Segoe UI", 20, "bold"), width=3)
            logo_lbl.pack(side="left", padx=(0, 10))
            
        title_box = tk.Frame(header, bg=BG_MAIN)
        title_box.pack(side="left")
        tk.Label(title_box, text="Auto Spec Program", fg=TEXT_PRI, bg=BG_MAIN, font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tk.Label(title_box, text="데이터베이스 기준으로 엑셀 측정값을 자동 검사하고 보고서를 생성합니다", fg=TEXT_SEC, bg=BG_MAIN, font=("Segoe UI", 10)).pack(anchor="w")
        tk.Label(header, text="Ver 2.0", fg=TEXT_SEC, bg=BG_CARD, font=("Segoe UI", 10), padx=15, pady=6).pack(side="right")

        left_panel = tk.Frame(self, bg=BG_MAIN)
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(40, 15), pady=(0, 20))
        left_panel.grid_rowconfigure(1, weight=1)

        self.db_card = DbCard(left_panel)
        self.db_card.grid(row=0, column=0, sticky="ew", pady=(0, 20))

        self.upload_zone = DashedUploadDropZone(left_panel)
        self.upload_zone.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
        self.upload_zone.bind("<Button-1>", lambda e: self._select_file())

        btn_grid = tk.Frame(left_panel, bg=BG_MAIN)
        btn_grid.grid(row=2, column=0, sticky="ew")
        btn_grid.grid_columnconfigure(0, weight=1)
        btn_grid.grid_columnconfigure(1, weight=1)

        self.btn_run = ModernRoundedButton(btn_grid, "처리 시작", self._run, bg_color=PRIMARY, hover_color=PRIMARY_HOV, text_color="#000", icon="▷")
        self.btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.btn_open = ModernRoundedButton(btn_grid, "결과 폴더", self._open_result_folder, bg_color=BTN_DARK, hover_color=BTN_DARK_HOV, text_color=TEXT_PRI, icon="📁")
        self.btn_open.grid(row=0, column=1, sticky="ew")

        self.btn_unmatched = ModernRoundedButton(left_panel, "미일치 항목 보기", self._show_unmatched, height=40, bg_color=WARN_BG, hover_color=WARN_HOV, text_color="#fca5a5", icon="⚠")
        self.btn_unmatched.grid(row=3, column=0, sticky="ew", pady=(15, 0))
        self.btn_unmatched.configure(state="disabled")

        right_panel = tk.Frame(self, bg=BG_MAIN)
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(15, 40), pady=(0, 20))
        rf_header = tk.Frame(right_panel, bg=BG_MAIN)
        rf_header.pack(fill="x", pady=(0, 10))
        tk.Label(rf_header, text=">_ 터미널 로그", fg=PRIMARY, bg=BG_MAIN, font=("Segoe UI", 12, "bold")).pack(side="left")
        self.lbl_log_count = tk.Label(rf_header, text="0 entries", fg=TEXT_SEC, bg=BG_CARD, font=("Segoe UI", 9), padx=10, pady=4)
        self.lbl_log_count.pack(side="right")

        self.log_container = ScrollableLogFrame(right_panel)
        self.log_container.pack(fill="both", expand=True)

        self.status_bar = FlatStatusBar(self)
        self.status_bar.grid(row=2, column=0, columnspan=2, sticky="ew")

    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            self.upload_zone.drop_target_register(DND_FILES)
            self.upload_zone.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        path = event.data.strip().strip("{}")
        self._set_file(path)

    def _select_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        if path: self._set_file(path)

    def _set_file(self, path):
        self.excel_path = path
        fname = os.path.basename(path)
        self.upload_zone.set_selected(fname)
        self._log(f"[안내] 엑셀 파일 선택 완료: {fname}", "green")

    def _log(self, msg, color=None):
        self.log_queue.put((msg, color))

    def _process_log_queue(self):
        try:
            while True:
                msg, color = self.log_queue.get_nowait()
                c_cnt = self.log_container.add_log(msg, color)
                self.lbl_log_count.configure(text=f"{c_cnt} entries")
        except queue.Empty: pass
        self.after(50, self._process_log_queue)

    def _load_db_on_start(self):
        def task():
            try:
                self.db_records = load_database()
                self.after(0, lambda: self.db_card.set_loaded(len(self.db_records)))
                self._log(f"[DB] 총 {len(self.db_records)}개 로드 완료 (IC/MOSFET/DIODE/CAP/TR)", "green")
            except Exception as e:
                self.after(0, lambda: self.db_card.set_error(e))
                self._log(f"[오류] DB 연결 실패: {e}", "red")
        threading.Thread(target=task, daemon=True).start()

    def _run(self):
        if not self.excel_path:
            messagebox.showwarning("파일 없음", "먼저 엑셀 파일을 업로드하세요.")
            return
        if not self.db_records:
            messagebox.showerror("DB 오류", "Database가 준비되지 않았습니다.")
            return

        self.btn_run.configure(state="disabled")
        self.status_bar.start_progress()

        def task():
            try:
                self._log(f"[System] 프로세스를 시작합니다: {self.excel_path}")
                matched, checked, result_path, unmatched = process_excel(
                    self.excel_path, self.db_records, self._log
                )
                self.last_unmatched = unmatched
                unmatched_count = len(unmatched)
                
                def update_ui_success():
                    self.status_bar.stop_progress(f"완료 — {checked}개 확인 / {matched}개 일치 / {unmatched_count}개 실패", True)
                    if unmatched_count > 0:
                        self.btn_unmatched.configure(state="normal")
                    self._open_result_folder()
                self.after(0, update_ui_success)
            except Exception as e:
                self._log(f"[오류] 실패: {e}", "red")
                self.after(0, lambda: self.status_bar.stop_progress(f"런타임 오류 발생: {e}", False))
                self.after(0, lambda: messagebox.showerror("처리 오류", str(e)))
            finally:
                self.after(0, lambda: self.btn_run.configure(state="normal"))

        threading.Thread(target=task, daemon=True).start()

    def _open_result_folder(self):
        result_dir = os.path.join(BASE_DIR, "Result")
        if os.path.exists(result_dir):
            try: os.startfile(result_dir)
            except: pass
        else:
            messagebox.showinfo("안내", "아직 Result 폴더가 없습니다.")

    def _show_unmatched(self):
        if not self.last_unmatched: return
        win = tk.Toplevel(self)
        win.title("미일치 항목")
        win.geometry("500x400")
        win.configure(bg=BG_MAIN)
        tk.Label(win, text=f"⚠ 미일치 항목 ({len(self.last_unmatched)}개)", bg=BG_MAIN, fg=ERROR_COL, font=("Segoe UI", 12, "bold")).pack(pady=15, padx=15, anchor="w")
        tb = scrolledtext.ScrolledText(win, bg=BG_CARD, fg=TEXT_PRI, font=("Consolas", 10), relief="flat")
        tb.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        for i, (r_num, pn) in enumerate(self.last_unmatched, 1):
            tb.insert("end", f"{i:3d}. 행{r_num:4d}: {pn}\n")
        tb.configure(state="disabled")

def main():
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        class AppDnD(TkinterDnD.Tk):
            def __init__(self):
                super().__init__()
                self.title("Auto Spec Program")
                self.state('zoomed')
                self.minsize(1024, 768)
                self.configure(bg=BG_MAIN)
                self.log_queue = queue.Queue()
                self.db_records = []
                self.excel_path = None
                self.last_unmatched = []

                App._build_ui(self)
                
                self.upload_zone.drop_target_register(DND_FILES)
                self.upload_zone.dnd_bind("<<Drop>>", App._on_drop.__get__(self))
                
                App._process_log_queue(self)
                App._load_db_on_start(self)

            _select_file = App._select_file
            _set_file = App._set_file
            _log = App._log
            _run = App._run
            _open_result_folder = App._open_result_folder
            _show_unmatched = App._show_unmatched
            _process_log_queue = App._process_log_queue
            _load_db_on_start = App._load_db_on_start

        app = AppDnD()
    except ImportError:
        app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
