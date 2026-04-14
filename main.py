import os
import json
import tempfile
import anthropic

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="AI PPT Generator")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 정적 파일 (frontend) ──────────────────────────────────────────
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
    from pptx.util import Inches
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

        # 전체 배경
        add_rect(slide, 0, 0, 10, 5.625, t["bg"])

        if stype == "title":
            # 왼쪽 accent bar
            add_rect(slide, 0.4, 2.0, 0.08, 1.5, t["accent"])
            add_text(slide, slide_info.get("title", slides_data.get("title", "")),
                     0.6, 1.9, 8.8, 1.0, t["title"], 36, bold=True)
            add_text(slide, slide_info.get("subtitle", slides_data.get("subtitle", "")),
                     0.6, 3.0, 8.8, 0.7, t["body"], 16)
            # 하단 바
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
            add_rect(slide, 0, 0, 10, 0.9, t["header"])
            add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)
            # 구분선
            add_rect(slide, 4.95, 1.0, 0.08, 4.0, t["accent"])
            # 왼쪽
            left = slide_info.get("left", {})
            add_text(slide, left.get("heading", ""), 0.3, 1.0, 4.4, 0.5, t["accent"], 14, bold=True)
            for j, pt in enumerate(left.get("points", [])):
                add_text(slide, "• " + pt, 0.3, 1.6 + j * 0.6, 4.4, 0.55, t["body"], 13)
            # 오른쪽
            right = slide_info.get("right", {})
            add_text(slide, right.get("heading", ""), 5.2, 1.0, 4.4, 0.5, t["accent"], 14, bold=True)
            for j, pt in enumerate(right.get("points", [])):
                add_text(slide, "• " + pt, 5.2, 1.6 + j * 0.6, 4.4, 0.55, t["body"], 13)

        else:  # content
            add_rect(slide, 0, 0, 10, 0.9, t["header"])
            add_text(slide, slide_info.get("title", ""), 0.4, 0.1, 9.2, 0.7, t["title"], 22, bold=True)
            add_rect(slide, 0.4, 1.0, 1.2, 0.05, t["accent"])
            for j, bullet in enumerate(slide_info.get("bullets", [])):
                add_rect(slide, 0.4, 1.2 + j * 0.78, 0.06, 0.38, t["accent"])
                add_text(slide, bullet, 0.65, 1.2 + j * 0.78, 9.0, 0.65, t["body"], 14)
            # 슬라이드 번호
            add_text(slide, str(i + 1), 9.0, 4.9, 0.6, 0.35, t["accent"], 10, align=PP_ALIGN.RIGHT)

    # 임시 파일로 저장
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(tmp.name)
    tmp.close()
    return tmp.name


# ── Claude API 호출 ───────────────────────────────────────────────
def call_claude(file_content: str, file_name: str, slide_count: int, language: str, purpose: str) -> dict:
    api_key = os.getenv("Anthropic_API_KEY_Win001")
    if not api_key:
        raise HTTPException(status_code=500, detail="Anthropic_API_KEY_Win001가 .env에 없습니다.")

    client = anthropic.Anthropic(api_key=api_key)

    purpose_map = {
        "general": "일반 발표",
        "business": "비즈니스 제안서",
        "education": "교육/강의 자료",
        "research": "연구 보고서",
    }

    prompt = f"""당신은 전문 프레젠테이션 디자이너입니다.
아래 파일 내용을 분석하여 {language}로 된 {slide_count}장의 프레젠테이션 슬라이드 구성을 만들어주세요.

파일명: {file_name}
발표 목적: {purpose_map.get(purpose, "일반 발표")}
슬라이드 수: {slide_count}장

파일 내용:
{file_content[:4000]}

다음 JSON 형식으로만 응답하세요 (다른 텍스트 없이 순수 JSON만):
{{
  "title": "프레젠테이션 제목",
  "subtitle": "부제목",
  "author": "작성자",
  "slides": [
    {{
      "type": "title",
      "title": "발표 제목",
      "subtitle": "부제목"
    }},
    {{
      "type": "content",
      "title": "슬라이드 제목",
      "bullets": ["핵심 내용 1", "핵심 내용 2", "핵심 내용 3"]
    }},
    {{
      "type": "two-column",
      "title": "비교 슬라이드",
      "left": {{"heading": "왼쪽 제목", "points": ["항목 1", "항목 2"]}},
      "right": {{"heading": "오른쪽 제목", "points": ["항목 1", "항목 2"]}}
    }},
    {{
      "type": "closing",
      "title": "감사합니다",
      "message": "핵심 메시지",
      "contact": "Q&A"
    }}
  ]
}}

총 {slide_count}장: title 1장 + content/two-column 여러 장 + closing 1장으로 구성하세요."""

    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    # JSON 펜스 제거
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()
    return json.loads(raw)


# ── 엔드포인트 ────────────────────────────────────────────────────
@app.post("/generate")
async def generate(
    file: UploadFile = File(...),
    slide_count: int = Form(8),
    language: str = Form("Korean"),
    theme: str = Form("dark"),
    purpose: str = Form("general"),
):
    # 파일 읽기
    content_bytes = await file.read()
    try:
        file_content = content_bytes.decode("utf-8")
    except UnicodeDecodeError:
        file_content = f"[파일명: {file.filename}] 바이너리 파일입니다. 파일 이름 기반으로 프레젠테이션을 만들어주세요."

    # Claude 호출
    try:
        slides_data = call_claude(file_content, file.filename, slide_count, language, purpose)
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=500, detail=f"AI 응답 파싱 오류: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    # PPTX 생성
    pptx_path = build_pptx(slides_data, theme)
    safe_title = slides_data.get("title", "presentation").replace(" ", "_")[:40]
    from urllib.parse import quote
    encoded_title = quote(safe_title)

    return FileResponse(
        pptx_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{encoded_title}.pptx",
        headers={"X-Slide-Count": str(len(slides_data.get("slides", []))),
                 "X-Presentation-Title": encoded_title}
    )


@app.get("/health")
def health():
    key = os.getenv("Anthropic_API_KEY_Win001", "")
    return {"status": "ok", "api_key_set": bool(key)}
