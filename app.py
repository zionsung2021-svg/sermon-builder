import io
import os
import re
import json
import sqlite3
import urllib.request
import urllib.error
import base64
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

# =============================================================
# 0. 디자인 상수 정의
# =============================================================
FONT_NAME = "경기천년제목 Medium"

COLOR_BG = RGBColor(0, 32, 96)
COLOR_GOLD = RGBColor(255, 192, 0)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_YELLOW = RGBColor(255, 255, 0)

def set_slide_background(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_BG

def format_run(run, text, font_size_pt, color_rgb, bold=False, underline=False):
    run.text = text
    run.font.name = FONT_NAME
    run.font.size = Pt(font_size_pt)
    run.font.color.rgb = color_rgb
    run.font.bold = bold
    run.font.underline = underline
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = OxmlElement('a:ea')
        rPr.append(ea)
    ea.set('typeface', FONT_NAME)

class BibleManager:
    BOOK_ALIASES = {
        "창세기": "창세기", "창": "창세기", "출애굽기": "출애굽기", "출": "출애굽기",
        "레위기": "레위기", "레": "레위기", "민수기": "민수기", "민": "민수기",
        "신명기": "신명기", "신": "신명기", "여호수아": "여호수아", "수": "여호수아",
        "사사기": "사사기", "삿": "사사기", "룻기": "룻기", "룻": "룻기",
        "사무엘상": "사무엘상", "삼상": "사무엘상", "사무엘하": "사무엘하", "삼하": "사무엘하",
        "열왕기상": "열왕기상", "왕상": "열왕기상", "열왕기하": "열왕기하", "왕하": "열왕기하",
        "역대상": "역대상", "대상": "역대상", "역대하": "역대하", "대하": "역대하",
        "에스라": "에스라", "스": "에스라", "느헤미야": "느헤미야", "느": "느헤미야",
        "에스더": "에스더", "에": "에스더", "욥기": "욥기", "욥": "욥기",
        "시편": "시편", "시": "시편", "잠언": "잠언", "잠": "잠언",
        "전도서": "전도서", "전": "전도서", "아가": "아가", "아": "아가",
        "이사야": "이사야", "사": "이사야", "예레미야": "예레미야", "렘": "예레미야",
        "예레미야애가": "예레미야애가", "애": "예레미야애가", "에스겔": "에스겔", "겔": "에스겔",
        "다니엘": "다니엘", "단": "다니엘", "호세아": "호세아", "호": "호세아",
        "요엘": "요엘", "욜": "요엘", "아모스": "아모스", "암": "아모스",
        "오바댜": "오바댜", "옵": "오바댜", "요나": "요나", "욘": "요나",
        "미가": "미가", "미": "미가", "나훔": "나훔", "나": "나훔",
        "하박국": "하박국", "합": "하박국", "스바냐": "스바냐", "습": "스바냐",
        "학개": "학개", "학": "학개", "스가랴": "스가랴", "슥": "스가랴",
        "말라기": "말라기", "말": "말라기",
        "마태복음": "마태복음", "마": "마태복음", "마가복음": "마가복음", "막": "마가복음",
        "누가복음": "누가복음", "눅": "누가복음", "요한복음": "요한복음", "요": "요한복음",
        "사도행전": "사도행전", "행": "사도행전", "로마서": "로마서", "롬": "로마서",
        "고린도전서": "고린도전서", "고전": "고린도전서", "고린도후서": "고린도후서", "고후": "고린도후서",
        "갈라디아서": "갈라디아서", "갈": "갈라디아서", "에베소서": "에베소서", "엡": "에베소서",
        "빌립보서": "빌립보서", "빌": "빌립보서", "골로새서": "골로새서", "골": "골로새서",
        "데살로니가전서": "데살로니가전서", "살전": "데살로니가전서", "데살로니가후서": "데살로니가후서", "살후": "데살로니가후서",
        "디모데전서": "디모데전서", "딤전": "디모데전서", "디모데후서": "디모데후서", "딤후": "디모데후서",
        "디도서": "디도서", "딛": "디도서", "빌레몬서": "빌레몬서", "몬": "빌레몬서",
        "히브리서": "히브리서", "히": "히브리서", "야고보서": "야고보서", "약": "야고보서",
        "베드로전서": "베드로전서", "벧전": "베드로전서", "베드로후서": "베드로후서", "벧후": "베드로후서",
        "요한일서": "요한일서", "요일": "요한일서", "요한이서": "요한이서", "요이": "요한이서",
        "요한삼서": "요한삼서", "요삼": "요한삼서", "유다서": "유다서", "유": "유다서",
        "요한계시록": "요한계시록", "계": "요한계시록"
    }

    BUILTIN_VERSES = {
        ("요한복음", 6, 15): "그러므로 예수께서 그들이 와서 자기를 **억지로** 붙들어 **임금**으로 삼으려는 줄 아시고 다시 혼자 **산**으로 떠나 가시니라",
        ("요한복음", 6, 24): "무리가 거기에 예수도 안 계시고 제자들도 없음을 보고 곧 배들을 타고 예수를 찾으러 가버나움으로 가서",
        ("요한복음", 6, 26): "예수께서 대답하여 이르시되 내가 진실로 진실로 너희에게 이르노니 너희가 나를 찾는 것은 표적을 본 까닭이 아니요 **떡을 먹고 배부른 까닭**이로다",
        ("요한복음", 6, 28): "그들이 묻되 우리가 어떻게 하여야 하나님의 일을 하오리이까",
        ("요한복음", 6, 29): "예수께서 대답하여 이르시되 하나님께서 보내신 이를 **믿는 것**이 하나님의 일이니라 하시니",
        ("요한복음", 6, 30): "그들이 묻되 그러면 우리가 보고 당신을 믿도록 행하시는 표적이 무엇이니이까, 하시는 일이 무엇이니이까",
        ("요한복음", 6, 33): "하나님의 떡은 하늘에서 내려 세상에 **생명을 주는 것**이니라",
        ("요한복음", 6, 34): "그들이 이르되 주여 이 떡을 항상 우리에게 주소서",
        ("요한복음", 6, 35): "예수께서 이르시되 **나는 생명의 떡이니** 내게 오는 자는 결코 주리지 아니할 터이요 나를 믿는 자는 영원히 목마르지 아니하리라",
        ("마태복음", 16, 26): "사람이 만일 온 천하를 얻고도 **제 목숨을 잃으면** 무엇이 유익하리요 사람이 무엇을 주고 제 목숨과 바꾸겠느냐",
        ("마가복음", 9, 22): "귀신이 그를 죽이려고 불과 물에 자주 던졌나이다 그러나 무엇을 하실 수 있거든 우리를 불쌍히 여기사 도와 주옵소서",
        ("마가복음", 9, 23): "예수께서 이르시되 할 수 있거든이 무슨 말이냐 **믿는 자에게는 능히 하지 못할 일이 없느니라** 하시니",
        ("마가복음", 9, 24): "곧 그 아이의 아버지가 소리를 질러 이르되 내가 믿나이다 **나의 믿음 없는 것을 도와 주소서** 하더라",
        ("로마서", 10, 17): "그러므로 **믿음은 들음에서 나며** 들음은 그리스도의 말씀으로 말미암았느니라"
    }

    def __init__(self, db_filename="bible.db"):
        # Streamlit 클라우드 환경에서도 정확히 찾도록 절대 경로 지정
        base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
        self.db_path = os.path.join(base_dir, db_filename)
        self.has_sqlite = os.path.exists(self.db_path)

    def normalize_book_name(self, raw_book: str) -> str:
        clean = raw_book.replace(" ", "").strip()
        return self.BOOK_ALIASES.get(clean, raw_book.strip())

    def get_verse_text(self, book: str, chapter: int, verse_start: int, verse_end: int = None) -> str:
        canonical_book = self.normalize_book_name(book)
        end_v = verse_end if verse_end else verse_start

        if self.has_sqlite:
            try:
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT verse, content FROM bible WHERE book=? AND chapter=? AND verse BETWEEN ? AND ? AND version='개역개정' ORDER BY verse",
                    (canonical_book, chapter, verse_start, end_v)
                )
                rows = cursor.fetchall()
                conn.close()
                if rows:
                    return "\n".join([f"({r[0]}) {r[1]}" if len(rows) > 1 else r[1] for r in rows])
                else:
                    return f"[⚠️ 오류: DB에 '{canonical_book} {chapter}장 {verse_start}절' (개역개정) 데이터가 없습니다.]"
            except Exception as e:
                return f"[⚠️ DB 시스템 에러: {str(e)}]"

        result_texts = []
        for v in range(verse_start, end_v + 1):
            key = (canonical_book, chapter, v)
            if key in self.BUILTIN_VERSES:
                txt = self.BUILTIN_VERSES[key]
                if verse_end and verse_end > verse_start:
                    result_texts.append(f"({v}) {txt}")
                else:
                    result_texts.append(txt)
        if result_texts:
            return "\n".join(result_texts)
        
        # 만약 여기까지 왔다면 DB 파일이 없거나 쿼리에 실패한 것임
        if not self.has_sqlite:
            return f"[⚠️ 오류: 'bible.db' 파일을 찾을 수 없습니다. GitHub에 함께 업로드(Commit) 되었는지 확인해주세요.]"
        else:
            return f"[⚠️ 오류: DB에서 '{canonical_book} {chapter}장 {verse_start}절'을 찾을 수 없습니다.]"

bible_service = BibleManager()

@st.cache_data(ttl=300, show_spinner=False)
def fetch_supported_models(api_key: str):
    if not api_key: return ["gemini-2.0-flash"]
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
    try:
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            models = [m.get("name", "").replace("models/", "") for m in data.get("models", []) if "generateContent" in m.get("supportedGenerationMethods", []) and m.get("name")]
            
            # Sort models to prioritize 2.5-flash or flash-latest
            priority = ["gemini-2.5-flash", "gemini-flash-latest", "gemini-2.5-pro"]
            for p in reversed(priority):
                if p in models:
                    models.remove(p)
                    models.insert(0, p)
            return models if models else ["gemini-2.5-flash"]
    except Exception:
        return ["gemini-2.5-flash", "gemini-pro-latest"]

def extract_text_from_handwriting(image_bytes: bytes, mime_type: str, api_key: str, model_name: str) -> str:
    clean_model = model_name.replace("models/", "").strip()
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{clean_model}:generateContent?key={api_key}"
    b64_data = base64.b64encode(image_bytes).decode("utf-8")
    prompt = "첨부된 이미지는 목사님의 손글씨(또는 인쇄된) 설교 원고입니다. 이 이미지에 적힌 글씨를 정확하게 읽어서 텍스트로 변환해 주세요. 가독성이 떨어지는 부분은 기독교 설교 문맥에 맞게 자연스럽게 보정하고, 원래의 줄바꿈과 단락을 최대한 유지해 주세요. 설명이나 마크다운 없이 순수 텍스트 원문만 출력하세요."
    payload = {
        "contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": mime_type, "data": b64_data}}]}],
        "generationConfig": {"temperature": 0.1}
    }
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            return res_data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception as e:
        raise RuntimeError(f"손글씨 인식 실패: {str(e)}")

def analyze_sermon_with_gemini(script_text: str, api_key: str, model_name: str = "gemini-2.5-flash") -> dict:
    import time
    clean_model = model_name.replace("models/", "").strip()
    system_instruction = """당신은 기독교 설교학(Homiletics)과 교회 강단 프레젠테이션 디자인 전문가입니다.
설교 원고를 분석하여 강단 송출용 슬라이드 뼈대를 JSON으로 추출하십시오.
성경 본문 전체를 길게 출력하지 말고, 성경 장·절과 '강조할 핵심 단어 1~3개'만 간결하게 추출하십시오.

[반드시 준수할 JSON 규격]:
{
  "title": "설교 제목",
  "scripture_range": "본문 성경 범위 (예: 요한복음 6장 22-36절)",
  "slides": [
    {"type": "TITLE", "title": "설교 제목", "subtitle": "(본문 범위)"},
    {"type": "MAIN_POINT", "point_header": "1. 대지 명제"},
    {"type": "SCRIPTURE", "reference": "요한복음 6장 15절", "highlights": ["억지로", "임금", "산"]}
  ]
}
- 규칙: 인용 성구는 본문 텍스트 없이 'reference'와 'highlights' 목록만 작성하십시오."""
    prompt = f"{system_instruction}\n\n[설교 원고]:\n{script_text}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1, "thinkingConfig": {"thinkingBudget": 0}}
    }
    last_error = None
    for current_model in [clean_model, "gemini-1.5-flash"]:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{current_model}:generateContent?key={api_key}"
        req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers={"Content-Type": "application/json"}, method="POST")
        for attempt in range(2):
            try:
                with urllib.request.urlopen(req, timeout=30) as response:
                    res_data = json.loads(response.read().decode("utf-8"))
                    parsed_json = json.loads(res_data["candidates"][0]["content"]["parts"][0]["text"])
                    for s in parsed_json.get("slides", []):
                        if s.get("type") == "SCRIPTURE":
                            ref = s.get("reference", "")
                            highlights = s.get("highlights", [])
                            v_match = re.search(r'([가-힣\s]+?)\s*(\d+)\s*(?:장|:)\s*(\d+)(?:[-~](\d+))?', ref)
                            if v_match:
                                b, c = v_match.group(1).strip(), int(v_match.group(2))
                                sv, ev = int(v_match.group(3)), int(v_match.group(4)) if v_match.group(4) else int(v_match.group(3))
                                canon = bible_service.normalize_book_name(b)
                                body_txt = bible_service.get_verse_text(canon, c, sv, ev)
                                for kw in highlights:
                                    if kw and kw in body_txt: body_txt = body_txt.replace(kw, f"**{kw}**")
                                s["text"] = body_txt
                            else: s["text"] = s.get("text", "")
                    return parsed_json
            except urllib.error.HTTPError as e:
                last_error = e
                if e.code in [503, 429]: time.sleep(2); continue
                break
            except Exception as e:
                last_error = e
                break
    raise RuntimeError(f"AI 서버 응답 지연: {str(last_error)}")

def parse_smart_sermon_script(text: str, auto_fill_bible: bool = True, standardize_points: bool = True) -> dict:
    raw_lines = [l.strip() for l in text.strip().split('\n') if l.strip()]
    if not raw_lines: return {"title": "설교 제목", "scripture": "", "slides": []}
    title, scripture_range, body_start_idx = "", "", 0
    for i, line in enumerate(raw_lines[:6]):
        t_m = re.search(r'^(?:[#*\s\-]*)?(?:제목|설교\s*제목)\s*[:：]\s*([^\n\r]+)', line)
        s_m = re.search(r'^(?:[#*\s\-]*)?(?:본문|성경|성경\s*말씀|말씀\s*본문)\s*[:：]\s*([^\n\r]+)', line)
        if t_m: title, body_start_idx = t_m.group(1).strip(), max(body_start_idx, i + 1)
        if s_m: scripture_range, body_start_idx = s_m.group(1).strip(), max(body_start_idx, i + 1)
    if not title:
        first_m = re.search(r'^(.*?)\s*[\(\[](.*?)[\)\]]$', raw_lines[0])
        if first_m: title, scripture_range, body_start_idx = first_m.group(1).strip(), first_m.group(2).strip(), 1
        else: title, body_start_idx = raw_lines[0], 1
    scripture_range = re.sub(r'^[\(\[]|[\)\]]$', '', scripture_range).strip()
    slides = [{"type": "TITLE", "title": title, "subtitle": f"({scripture_range})" if scripture_range else ""}]
    book_regex = "|".join(re.escape(b) for b in sorted(bible_service.BOOK_ALIASES.keys(), key=lambda x: -len(x)))
    inline_verse_regex = re.compile(r'(?<![가-힣])(' + book_regex + r')\s*(\d+)\s*(?:장|:)\s*(\d+)(?:\s*절)?(?:\s*(?:[-~]|에서|부터|까지)\s*(\d+)\s*절?)?')
    point_prefix_regex = re.compile(r'^(?:[#*\s\-]*)?(\d+[\.\)]|\(\d+\)|(?:첫|두|둘|세|셋|네|넷|다섯|여섯|일곱|여덟|아홉|열)\s*(?:번째|째)(?:로)?[,.\s:]*|[제\s]*[12345일이삼사오]\s*대지[:\s]*|대지\s*[12345일이삼사오][:\s]*|[IVXLCDM]+\.)\s*(.*)')
    ordinal_dict = {'첫': 1, '두': 2, '둘': 2, '세': 3, '셋': 3, '네': 4, '넷': 4, '다섯': 5, '여섯': 6, '일곱': 7, '여덟': 8, '아홉': 9, '열': 10}
    body_lines, point_counter, i = raw_lines[body_start_idx:], 1, 0
    while i < len(body_lines):
        line = body_lines[i]
        p_match = point_prefix_regex.match(line)
        if p_match:
            p_prefix, p_body = p_match.group(1).strip(), p_match.group(2).strip()
            if standardize_points:
                num = next((int(m.group(0)) for m in [re.search(r'\d+', p_prefix)] if m), next((v for k, v in ordinal_dict.items() if p_prefix.startswith(k)), point_counter))
                clean_point_header = f"{num}. {p_body}"
            else: clean_point_header = f"{p_prefix} {p_body}".strip()
            slides.append({"type": "MAIN_POINT", "point_header": clean_point_header})
            point_counter, i = point_counter + 1, i + 1
            continue
        v_matches = list(inline_verse_regex.finditer(line))
        if v_matches:
            for vm in v_matches:
                raw_book, ch, sv, ev = vm.group(1).strip(), int(vm.group(2)), int(vm.group(3)), int(vm.group(4)) if vm.group(4) else int(vm.group(3))
                canon_book = bible_service.normalize_book_name(raw_book)
                ref_str = f"{canon_book} {ch}장 {sv}-{ev}절" if ev > sv else f"{canon_book} {ch}장 {sv}절"
                verse_text = ""
                curr_quotes = re.findall(r'["“](.*?)["”]', line)
                if curr_quotes: verse_text = curr_quotes[0]
                elif i + 1 < len(body_lines):
                    next_line = body_lines[i + 1]
                    if not point_prefix_regex.match(next_line) and not inline_verse_regex.search(next_line):
                        next_quotes = re.findall(r'["“](.*?)["”]', next_line)
                        if next_quotes: verse_text, i = next_quotes[0], i + 1
                        elif len(next_line) < 180 and any(k in next_line for k in ["하셨습니다", "가시니라", "이르시되", "하더라", "말미암았느니라", "**"]): verse_text, i = next_line, i + 1
                if not verse_text and auto_fill_bible: verse_text = bible_service.get_verse_text(canon_book, ch, sv, ev)
                if verse_text:
                    if len(verse_text) > 160:
                        sentences, chunk, part = re.split(r'(?<=[.?!])\s+', verse_text), "", 1
                        for s in sentences:
                            if len(chunk) + len(s) > 150: slides.append({"type": "SCRIPTURE", "reference": f"{ref_str} ({part})", "text": chunk.strip()}); chunk, part = s + " ", part + 1
                            else: chunk += s + " "
                        if chunk.strip(): slides.append({"type": "SCRIPTURE", "reference": f"{ref_str} ({part})" if part > 1 else ref_str, "text": chunk.strip()})
                    else: slides.append({"type": "SCRIPTURE", "reference": ref_str, "text": verse_text})
            i += 1
            continue
        i += 1
    return {"title": title, "scripture": scripture_range, "slides": slides}

def build_sermon_pptx(parsed_data: dict) -> io.BytesIO:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    for item in parsed_data["slides"]:
        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide)
        stype = item["type"]
        if stype == "TITLE":
            box = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.733), Inches(3.5))
            p1 = box.text_frame.paragraphs[0]
            p1.alignment = PP_ALIGN.CENTER
            format_run(p1.add_run(), item["title"], 60, COLOR_GOLD)
            if item.get("subtitle"):
                p2 = box.text_frame.add_paragraph()
                p2.alignment, p2.space_before = PP_ALIGN.CENTER, Pt(20)
                format_run(p2.add_run(), item["subtitle"], 48, COLOR_WHITE)
        elif stype == "MAIN_POINT":
            box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.733), Inches(2.5))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            format_run(p.add_run(), item["point_header"], 64, COLOR_GOLD)
        elif stype == "SCRIPTURE":
            rbox = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.733), Inches(1.2))
            rp = rbox.text_frame.paragraphs[0]
            rp.alignment = PP_ALIGN.CENTER
            format_run(rp.add_run(), item["reference"], 50, COLOR_GOLD, True)
            bbox = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.733), Inches(4.6))
            btf = bbox.text_frame
            btf.word_wrap = True  # 텍스트 박스 경계선 자동 줄바꿈(Word Wrap) 강제 활성화
            bp = btf.paragraphs[0]
            bp.alignment = PP_ALIGN.CENTER
            bp.line_spacing = 1.35
            raw_text = item.get("text", "")
            
            # [의미 단위 스마트 줄바꿈]: 여러 절이 있을 경우 절 번호 뒤에서 띄어쓰기 및 줄바꿈
            raw_text = re.sub(r' (\(\d+\)) ', r'\n\n ', raw_text)
            
            # 주어(은,는,이,가)와 서술부 분리를 위한 가독성 줄바꿈 적용
            chunked_lines = []
            for line in raw_text.split("\n"):
                words = line.split(" ")
                current_chunk = ""
                for word in words:
                    current_chunk += word + " "
                    # 어절이 어느정도 쌓였고(8자 이상), 조사가 주격/연결어미로 끝나면 의미 단위로 줄바꿈
                    if len(current_chunk) >= 8 and re.search(r'(은|는|이|가|고|며)$', word.replace("**", "")):
                        chunked_lines.append(current_chunk.strip())
                        current_chunk = ""
                if current_chunk:
                    chunked_lines.append(current_chunk.strip())
            
            # 정제된 텍스트로 덮어쓰기
            raw_text = "\n".join(chunked_lines)
            tokens = re.split(r'(\*\*.*?\*\*)', raw_text)
            for token in tokens:
                if not token: continue
                run = bp.add_run()
                if token.startswith("**") and token.endswith("**"): format_run(run, token[2:-2], 44, COLOR_YELLOW, underline=True)
                else: format_run(run, token, 44, COLOR_WHITE)
        elif stype == "MESSAGE":
            box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.733), Inches(2.5))
            p = box.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            format_run(p.add_run(), item.get("text", ""), 50, COLOR_WHITE)
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

st.set_page_config(page_title="AI 설교-PPT 전자동 빌더", layout="wide", page_icon="📖")
st.title("📖 AI 강단 송출용 설교 PPT 전자동 빌더")
st.caption("손글씨 원고 촬영본을 AI가 정밀 판독하고, 대지 구조화, 성구 복원, **노란색 핵심 강조**까지 완벽한 PPTX로 빌드합니다.")

unstructured_sample_script = """2026년 10월 25일 주일예배 설교
제목: 나의 관심은 어디에 있는가?
본문: 요한복음 6장 22절-36절

오늘 우리는 요한복음 6장 15절 말씀을 통해 우리의 신앙을 돌아보고자 합니다.
"그러므로 예수께서 그들이 와서 자기를 **억지로** 붙들어 **임금**으로 삼으려는 줄 아시고 다시 혼자 **산**으로 떠나 가시니라"
첫째로, 왜 예수님을 찾습니까?
무리들은 예수님을 만나기 위해 배를 타고 가버나움까지 찾아왔습니다 (요 6:24).
"""

if "script_input" not in st.session_state:
    st.session_state["script_input"] = unstructured_sample_script

with st.sidebar:
    st.header("⚙️ 분석 엔진 설정")
    engine_choice = st.radio("분석 모드 선택", ["🤖 AI 심층 분석 모드 (Gemini)", "⚡ 스마트 규칙 기반 모드 (오프라인)"])
    api_key, selected_model = "", "gemini-2.5-flash"
    if "Gemini" in engine_choice:
        api_key = st.text_input("Gemini API Key", value=os.environ.get("GEMINI_API_KEY", ""), type="password")
        
        # API Key 입력 시 지원 모델 자동 조회
        fetched_models = fetch_supported_models(api_key.strip())
        
        # 필수 안정화 모델 목록 강제 추가 + 중복 제거 + 직접 입력
        base_models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
        model_options = list(dict.fromkeys(base_models + fetched_models + ["직접 입력"]))
        
        model_selection = st.selectbox("사용 모델", model_options, index=0)
        
        if model_selection == "직접 입력":
            selected_model = st.text_input("모델명 직접 입력 (예: gemini-1.5-flash)", value="gemini-1.5-flash")
        else:
            selected_model = model_selection

left_col, right_col = st.columns([1, 1])

with left_col:
    st.subheader("1. 설교 원고 입력")
    tab1, tab2 = st.tabs(["✍️ 원고 직접 입력", "📸 손글씨 사진 자동 판독"])
    with tab1:
        st.info("💡 강조할 단어를 `**단어**`로 묶으면 노란색 밑줄로 서식화됩니다.")
        user_script = st.text_area("비정형 설교 원고", value=st.session_state["script_input"], height=480, key="editor")
        st.session_state["script_input"] = user_script
        execute_btn = st.button("🚀 AI 분석 및 강단 PPTX 생성", type="primary", use_container_width=True)
    with tab2:
        st.write("목사님이 쓰신 **손글씨 설교 노트**나 인쇄된 원고 사진을 올리면 AI가 정확히 읽어 텍스트로 변환합니다.")
        uploaded_image = st.file_uploader("이미지 업로드 (JPG, PNG)", type=["jpg", "jpeg", "png"])
        if uploaded_image and st.button("🔍 이미지에서 원고 텍스트 추출하기", use_container_width=True):
            if not api_key: st.error("이미지 인식 기능을 사용하려면 사이드바에 API Key를 입력해야 합니다.")
            else:
                with st.spinner("AI가 목사님의 손글씨를 정밀 판독하고 있습니다..."):
                    try:
                        extracted_text = extract_text_from_handwriting(uploaded_image.read(), "image/jpeg" if uploaded_image.name.lower().endswith(("jpg","jpeg")) else "image/png", api_key, selected_model)
                        st.session_state["script_input"] = extracted_text
                        st.success("✅ 손글씨 판독 완료! [✍️ 원고 직접 입력] 탭에서 확인하세요.")
                    except Exception as e: st.error(f"판독 실패: {e}")

with right_col:
    st.subheader("2. 분석 결과 및 PPTX 다운로드")
    if execute_btn and st.session_state["script_input"]:
        if "Gemini" in engine_choice:
            if not api_key: st.error("⚠️ Gemini API Key를 입력해 주세요.")
            else:
                with st.spinner(f"🤖 AI가 초고속으로 구조 분석 중입니다 (3~5초 소요)..."):
                    try: st.session_state["parsed_result"] = analyze_sermon_with_gemini(st.session_state["script_input"], api_key, selected_model)
                    except Exception as e: st.error(f"AI 분석 실패: {e}")
        else:
            st.session_state["parsed_result"] = parse_smart_sermon_script(st.session_state["script_input"])

    if "parsed_result" in st.session_state:
        res = st.session_state["parsed_result"]
        st.success(f"✅ 총 **{len(res.get('slides', []))}장** 슬라이드 생성 완료 (제목: {res.get('title')})")
        st.download_button("📥 완성된 PPTX 다운로드 (경기천년제목 & 네이비)", data=build_sermon_pptx(res), file_name=f"{res.get('title') or '설교자료'}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
        with st.container(height=420):
            for idx, s in enumerate(res.get("slides", []), 1):
                stype = s.get("type")
                if stype == "TITLE": st.markdown(f"**[{idx}p] 표지**: `{s.get('title')}` {s.get('subtitle', '')}")
                elif stype == "MAIN_POINT": st.warning(f"**[{idx}p] 대지 선포**: {s.get('point_header')}")
                elif stype == "SCRIPTURE":
                    with st.expander(f"**[{idx}p] 성구**: {s.get('reference')}", expanded=True): st.markdown(s.get("text", ""))
