import io
import os
import json
import tempfile
import anthropic

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="AI PPT Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def root():
    return FileResponse("static/index.html")


# ── 테마 정의 ────────────────────────────────────────────────────
THEMES = {
    "dark":      {"bg": "1E2761", "header": "0D1B5E", "title": "FFFFFF", "body": "CADCFC", "accent": "4FC3F7"},
    "corporate": {"bg": "1565C0", "header": "0D47A1", "title": "FFFFFF", "body": "E3F2FD", "accent": "FFD54F"},
    "modern":    {"bg": "4A148C", "header": "38006B", "title": "FFFFFF", "body": "EDE7F6", "accent": "CE93D8"},
    "warm":      {"bg": "B85042", "header": "8D2B1A", "title": "FFFFFF", "body": "ECE2D0", "accent": "F9E795"},
}

def hex_to_rgb(hex_str: str) -> RGBColor:
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return RGBColor(r, g, b)

def add_rect(slide, x, y, w, h, color_hex):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(color_hex)
    shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, color_hex, font_size, bold=False, align=PP_ALIGN.LEFT, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = hex_to_rgb(color_hex)
    run.font.name = font_name
    return txBox


# ── 파일 텍스트 추출 ──────────────────────────────────────────────
def extract_text_from_file(content_bytes: bytes, filename: str) -> str:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    if ext == "pdf":
        try:
            import pdfplumber
            with pdfplumber.open(io.BytesIO(content_bytes)) as pdf:
                texts = []
                for page in pdf.pages[:20]:
                    t = page.extract_text()
                    if t:
                        texts.append(t)
            return "\n\n".join(texts) if texts else f"[{filename}: PDF에서 텍스트를 추출할 수 없습니다]"
        except Exception as e:
            return f"[PDF 파싱 실패: {e}]"

    elif ext == "docx":
        try:
            from docx import Document
            doc = Document(io.BytesIO(content_bytes))
            parts = [p.text for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    parts.append(" | ".join(cell.text for cell in row.cells))
            return "\n".join(parts)
        except Exception as e:
            return f"[DOCX 파싱 실패: {e}]"

    elif ext == "csv":
        try:
            import csv
            text = content_bytes.decode("utf-8", errors="replace")
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)
            return "\n".join([" | ".join(row) for row in rows[:150]])
        except Exception as e:
            return f"[CSV 파싱 실패: {e}]"

    elif ext == "json":
        try:
            text = content_bytes.decode("utf-8", errors="replace")
            data = json.loads(text)
            return json.dumps(data, ensure_ascii=False, indent=2)[:8000]
        except Exception:
            try:
                return content_bytes.decode("utf-8", errors="replace")
            except Exception as e:
                return f"[JSON 파싱 실패: {e}]"

    else:
        for enc in ("utf-8", "cp949", "euc-kr"):
            try:
                return content_bytes.decode(enc)
            except UnicodeDecodeError:
                continue
        return f"[파일명: {filename}] 바이너리 파일 - 파일 이름 기반으로 프레젠테이션을 생성합니다."


# ── 차트 슬라이드 ─────────────────────────────────────────────────
CHART_TYPE_MAP = {
    "bar":     XL_CHART_TYPE.BAR_CLUSTERED,
    "column":  XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line":    XL_CHART_TYPE.LINE,
    "pie":     XL_CHART_TYPE.PIE,
    "area":    XL_CHART_TYPE.AREA,
}

def add_chart_slide(slide, slide_info, t):
    add_rect(slide, 0, 0, 10, 0.9, t["header"])
    add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)

    categories = slide_info.get("categories", [])
    series_list = slide_info.get("series", [])

    if not categories or not series_list:
        add_text(slide, "차트 데이터가 없습니다", 0.5, 2.5, 9.0, 1.0, t["body"], 14, align=PP_ALIGN.CENTER)
        return

    cd = ChartData()
    cd.categories = [str(c) for c in categories]
    for s in series_list:
        values = tuple(float(v) if v is not None else 0.0 for v in s.get("values", []))
        cd.add_series(s.get("name", ""), values)

    xl_type = CHART_TYPE_MAP.get(slide_info.get("chart_type", "column").lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

    try:
        chart_shape = slide.shapes.add_chart(
            xl_type,
            Inches(0.5), Inches(1.05),
            Inches(9.0), Inches(4.3),
            cd
        )
        chart = chart_shape.chart

        try:
            chart.chart_area.format.fill.solid()
            chart.chart_area.format.fill.fore_color.rgb = hex_to_rgb(t["bg"])
        except Exception:
            pass
        try:
            chart.plot_area.format.fill.solid()
            chart.plot_area.format.fill.fore_color.rgb = hex_to_rgb(t["bg"])
        except Exception:
            pass

        series_colors = [t["accent"], "FFFFFF", t["body"], "FF8C00", "00CED1"]
        try:
            for i, series in enumerate(chart.series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = hex_to_rgb(series_colors[i % len(series_colors)])
        except Exception:
            pass

        try:
            if xl_type == XL_CHART_TYPE.PIE or len(series_list) > 1:
                chart.has_legend = True
                chart.legend.font.size = Pt(11)
                chart.legend.font.color.rgb = hex_to_rgb(t["body"])
            else:
                chart.has_legend = False
        except Exception:
            pass

        try:
            chart.category_axis.tick_labels.font.color.rgb = hex_to_rgb(t["body"])
            chart.value_axis.tick_labels.font.color.rgb = hex_to_rgb(t["body"])
        except Exception:
            pass  # 파이 차트는 축 없음

    except Exception as e:
        add_text(slide, f"차트 생성 오류: {str(e)[:80]}", 0.5, 2.5, 9.0, 1.0, t["body"], 11)


# ── 테이블 슬라이드 ───────────────────────────────────────────────
def add_table_slide(slide, slide_info, t):
    add_rect(slide, 0, 0, 10, 0.9, t["header"])
    add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)

    headers = slide_info.get("headers", [])
    rows = slide_info.get("rows", [])

    if not headers:
        return

    cols = len(headers)
    total_rows = len(rows) + 1

    table_shape = slide.shapes.add_table(
        total_rows, cols,
        Inches(0.4), Inches(1.1),
        Inches(9.2), Inches(4.1)
    )
    tbl = table_shape.table

    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = hex_to_rgb(t["header"])
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(h)
        run.font.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = hex_to_rgb(t["accent"])
        run.font.name = "Calibri"

    for i, row_data in enumerate(rows):
        for j in range(cols):
            val = row_data[j] if j < len(row_data) else ""
            cell = tbl.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_to_rgb(t["bg"] if i % 2 == 0 else t["header"])
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(val)
            run.font.size = Pt(11)
            run.font.color.rgb = hex_to_rgb(t["body"])
            run.font.name = "Calibri"


# ── 프로세스 슬라이드 ─────────────────────────────────────────────
def add_process_slide(slide, slide_info, t):
    add_rect(slide, 0, 0, 10, 0.9, t["header"])
    add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)

    steps = slide_info.get("steps", [])[:5]
    if not steps:
        return

    n = len(steps)
    step_w = 9.0 / n
    start_x = 0.5

    for i, step in enumerate(steps):
        x = start_x + i * step_w
        box_w = step_w - 0.25

        add_rect(slide, x, 1.5, box_w, 0.65, t["accent"])
        add_text(slide, step.get("number", f"{i+1:02d}"), x, 1.5, box_w, 0.65,
                 t["bg"], 18, bold=True, align=PP_ALIGN.CENTER)

        add_text(slide, step.get("title", ""), x, 2.3, box_w, 0.55,
                 t["title"], 12, bold=True, align=PP_ALIGN.CENTER)

        add_text(slide, step.get("desc", ""), x, 2.95, box_w, 1.5,
                 t["body"], 10, align=PP_ALIGN.CENTER)

        if i < n - 1:
            add_text(slide, "▶", x + box_w + 0.02, 1.6, 0.22, 0.5,
                     t["accent"], 14, align=PP_ALIGN.CENTER)


# ── 통계/KPI 슬라이드 ─────────────────────────────────────────────
def add_stats_slide(slide, slide_info, t):
    add_rect(slide, 0, 0, 10, 0.9, t["header"])
    add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)

    stats = slide_info.get("stats", [])[:4]
    if not stats:
        return

    n = len(stats)
    card_w = 9.0 / n
    start_x = 0.5

    for i, stat in enumerate(stats):
        x = start_x + i * card_w
        w = card_w - 0.2

        add_rect(slide, x, 1.2, w, 3.5, t["header"])
        add_rect(slide, x, 1.2, w, 0.08, t["accent"])

        add_text(slide, stat.get("value", ""), x + 0.1, 1.5, w - 0.2, 1.3,
                 t["accent"], 36, bold=True, align=PP_ALIGN.CENTER)

        add_text(slide, stat.get("label", ""), x + 0.1, 2.9, w - 0.2, 0.6,
                 t["title"], 13, bold=True, align=PP_ALIGN.CENTER)

        if stat.get("desc"):
            add_text(slide, stat["desc"], x + 0.1, 3.55, w - 0.2, 0.7,
                     t["body"], 10, align=PP_ALIGN.CENTER)


# ── PPT 빌더 ─────────────────────────────────────────────────────
def build_pptx(slides_data: dict, theme_key: str) -> str:
    t = THEMES.get(theme_key, THEMES["dark"])
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(5.625)

    blank_layout = prs.slide_layouts[6]

    for i, slide_info in enumerate(slides_data.get("slides", [])):
        slide = prs.slides.add_slide(blank_layout)
        stype = slide_info.get("type", "content")

        # 새 슬라이드 타입 처리
        if stype == "chart":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_chart_slide(slide, slide_info, t)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

        elif stype == "table":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_table_slide(slide, slide_info, t)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

        elif stype == "process":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_process_slide(slide, slide_info, t)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

        elif stype == "stats":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_stats_slide(slide, slide_info, t)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

        elif stype == "title":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_rect(slide, 0.4, 2.0, 0.08, 1.5, t["accent"])
            add_text(slide, slide_info.get("title", slides_data.get("title", "")),
                     0.6, 1.9, 8.8, 1.0, t["title"], 36, bold=True)
            add_text(slide, slide_info.get("subtitle", slides_data.get("subtitle", "")),
                     0.6, 3.0, 8.8, 0.7, t["body"], 16)
            add_rect(slide, 0, 4.5, 10, 0.6, t["header"])
            add_text(slide, slides_data.get("author", ""),
                     0.4, 4.55, 9.2, 0.5, t["body"], 11)

        elif stype == "closing":
            add_rect(slide, 0, 0, 10, 5.625, t["header"])
            add_text(slide, slide_info.get("title", "감사합니다"),
                     0.5, 1.4, 9.0, 1.2, t["title"], 40, bold=True, align=PP_ALIGN.CENTER)
            if slide_info.get("message"):
                add_text(slide, slide_info["message"],
                         0.5, 2.9, 9.0, 0.8, t["accent"], 18, align=PP_ALIGN.CENTER)
            if slide_info.get("contact"):
                add_text(slide, slide_info["contact"],
                         0.5, 3.9, 9.0, 0.5, t["body"], 13, align=PP_ALIGN.CENTER)

        elif stype == "two-column":
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_rect(slide, 0, 0, 10, 0.9, t["header"])
            add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)
            add_rect(slide, 4.95, 1.0, 0.08, 4.0, t["accent"])
            left = slide_info.get("left", {})
            add_text(slide, left.get("heading", ""), 0.3, 1.0, 4.4, 0.5, t["accent"], 14, bold=True)
            for j, pt in enumerate(left.get("points", [])):
                add_text(slide, "• " + pt, 0.3, 1.6 + j * 0.6, 4.4, 0.55, t["body"], 13)
            right = slide_info.get("right", {})
            add_text(slide, right.get("heading", ""), 5.2, 1.0, 4.4, 0.5, t["accent"], 14, bold=True)
            for j, pt in enumerate(right.get("points", [])):
                add_text(slide, "• " + pt, 5.2, 1.6 + j * 0.6, 4.4, 0.55, t["body"], 13)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

        else:  # content
            add_rect(slide, 0, 0, 10, 5.625, t["bg"])
            add_rect(slide, 0, 0, 10, 0.9, t["header"])
            add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)
            add_rect(slide, 0.4, 1.0, 1.2, 0.05, t["accent"])
            for j, bullet in enumerate(slide_info.get("bullets", [])):
                add_rect(slide, 0.4, 1.2 + j * 0.78, 0.06, 0.38, t["accent"])
                add_text(slide, bullet, 0.65, 1.2 + j * 0.78, 9.0, 0.65, t["body"], 14)
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    tmp.close()
    return tmp.name


# ── Claude API 호출 ───────────────────────────────────────────────
SYSTEM_PROMPT = """당신은 전문 프레젠테이션 디자이너입니다. 파일 내용을 분석하여 시각적으로 풍부하고 설득력 있는 프레젠테이션을 만들어주세요.

사용 가능한 슬라이드 타입:
- "title": 표지 슬라이드 (항상 첫 장)
- "content": 글머리 기호 목록 (일반 내용)
- "two-column": 두 열 비교 슬라이드
- "chart": 차트/그래프 (수치·통계 데이터가 있을 때 필수 사용)
- "table": 표 슬라이드 (비교표, 데이터표)
- "process": 단계별 프로세스 다이어그램 (절차·단계 설명)
- "stats": KPI/통계 카드 (핵심 수치·지표)
- "closing": 마무리 슬라이드 (항상 마지막 장)

핵심 규칙:
1. 수치 데이터나 통계가 있으면 반드시 chart 또는 stats 슬라이드를 포함
2. 단계별 프로세스·절차가 있으면 process 슬라이드 사용
3. 비교 데이터가 있으면 table 또는 two-column 사용
4. 단순 텍스트만 있어도 chart/stats/process 타입을 창의적으로 추론하여 삽입
5. content 타입만 나열하는 것을 피하고 다양한 타입을 혼합할 것

순수 JSON만 응답하세요. 마크다운 코드블록 없이."""


def call_claude(file_content: str, file_name: str, slide_count: int, language: str, purpose: str) -> dict:
    api_key = os.getenv("Anthropic_API_KEY_Win001")
    if not api_key:
        raise HTTPException(status_code=500, detail="Anthropic_API_KEY_Win001가 .env에 없습니다.")

    client = anthropic.Anthropic(api_key=api_key)

    purpose_map = {
        "general":   "일반 발표",
        "business":  "비즈니스 제안서",
        "education": "교육/강의 자료",
        "research":  "연구 보고서",
    }

    user_prompt = f"""파일명: {file_name}
발표 목적: {purpose_map.get(purpose, "일반 발표")}
슬라이드 수: {slide_count}장
언어: {language}

파일 내용:
{file_content[:6000]}

아래 JSON 형식으로 응답하세요. 총 {slide_count}장 (title 1 + 내용 여러 장 + closing 1):

{{
  "title": "프레젠테이션 제목",
  "subtitle": "부제목",
  "author": "작성자 또는 출처",
  "slides": [
    {{"type":"title","title":"발표 제목","subtitle":"부제목"}},
    {{"type":"content","title":"슬라이드 제목","bullets":["핵심 내용 1","핵심 내용 2","핵심 내용 3"]}},
    {{"type":"chart","title":"차트 제목","chart_type":"column","categories":["1분기","2분기","3분기","4분기"],"series":[{{"name":"매출(억)","values":[120,145,160,180]}}]}},
    {{"type":"table","title":"비교표","headers":["항목","A안","B안","비고"],"rows":[["비용","100만","150만","A 유리"],["기간","3개월","2개월","B 유리"]]}},
    {{"type":"process","title":"프로세스","steps":[{{"number":"01","title":"기획","desc":"요구사항 수집"}},{{"number":"02","title":"설계","desc":"시스템 설계"}},{{"number":"03","title":"개발","desc":"구현 및 테스트"}},{{"number":"04","title":"배포","desc":"서비스 출시"}}]}},
    {{"type":"stats","title":"핵심 지표","stats":[{{"value":"98%","label":"고객 만족도","desc":"+3% YoY"}},{{"value":"2.4M","label":"월간 사용자","desc":"+40% YoY"}},{{"value":"150+","label":"파트너사","desc":"글로벌 확장"}}]}},
    {{"type":"two-column","title":"비교","left":{{"heading":"장점","points":["항목1","항목2"]}},"right":{{"heading":"단점","points":["항목1","항목2"]}}}},
    {{"type":"closing","title":"감사합니다","message":"핵심 메시지","contact":"Q&A"}}
  ]
}}"""

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=8192,
        system=[{
            "type": "text",
            "text": SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"}
        }],
        messages=[{"role": "user", "content": user_prompt}]
    )

    raw = message.content[0].text.strip()
    # JSON 펜스 제거 (혹시 포함된 경우)
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # 응답이 잘렸을 경우: closing 슬라이드를 강제 삽입하고 닫기 시도
        # 마지막 완전한 슬라이드 객체까지만 잘라내기
        last_brace = raw.rfind("},")
        if last_brace == -1:
            last_brace = raw.rfind("}")
        if last_brace != -1:
            trimmed = raw[:last_brace + 1]
            # slides 배열 닫기 + 최상위 객체 닫기
            closing = ',{"type":"closing","title":"감사합니다","message":"","contact":"Q&A"}]}'
            try:
                return json.loads(trimmed + closing)
            except json.JSONDecodeError:
                pass
        raise


# ── 엔드포인트 ────────────────────────────────────────────────────
@app.post("/generate")
async def generate(
    file: UploadFile = File(...),
    slide_count: int = Form(8),
    language: str = Form("Korean"),
    theme: str = Form("dark"),
    purpose: str = Form("general"),
):
    content_bytes = await file.read()
    file_content = extract_text_from_file(content_bytes, file.filename)

    try:
        slides_data = call_claude(file_content, file.filename, slide_count, language, purpose)
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=500, detail=f"AI 응답 파싱 오류: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    pptx_path = build_pptx(slides_data, theme)
    safe_title = slides_data.get("title", "presentation").replace(" ", "_")[:40]
    from urllib.parse import quote
    encoded_title = quote(safe_title)

    return FileResponse(
        pptx_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{encoded_title}.pptx",
        headers={
            "X-Slide-Count": str(len(slides_data.get("slides", []))),
            "X-Presentation-Title": encoded_title,
        }
    )


@app.get("/health")
def health():
    key = os.getenv("Anthropic_API_KEY_Win001", "")
    return {"status": "ok", "api_key_set": bool(key)}
