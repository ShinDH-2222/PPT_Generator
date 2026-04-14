# AI PPT Generator 🎯

파일을 업로드하면 Claude AI가 자동으로 PPT를 만들어주는 웹 앱입니다.

## 📁 프로젝트 구조

```
ppt-generator/
├── main.py              ← FastAPI 백엔드 서버
├── requirements.txt     ← Python 패키지 목록
├── .env                 ← API 키 (직접 만들어야 함)
├── .env.example         ← .env 샘플
└── static/
    └── index.html       ← 프론트엔드
```

---

## 🚀 로컬 실행 방법

### 1. .env 파일 만들기
```bash
cp .env.example .env
```
`.env` 파일을 열어서 API 키 입력:
```
ANTHROPIC_API_KEY=sk-ant-여기에_실제_키_입력
```

### 2. 패키지 설치
```bash
pip install -r requirements.txt
```

### 3. 서버 실행
```bash
uvicorn main:app --reload --port 8000
```

### 4. 브라우저에서 접속
```
http://localhost:8000
```

---

## ☁️ Railway 배포 방법 (외부 공개)

1. [railway.app](https://railway.app) 가입
2. New Project → Deploy from GitHub
3. 이 폴더를 GitHub에 올린 뒤 연결
4. Variables 탭에서 환경변수 추가:
   ```
   ANTHROPIC_API_KEY = sk-ant-...
   ```
5. 자동으로 배포되고 URL이 생성됩니다

---

## 🌐 Render 배포 방법 (무료 티어)

1. [render.com](https://render.com) 가입
2. New Web Service → GitHub 연결
3. Build Command: `pip install -r requirements.txt`
4. Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Environment Variables에 `ANTHROPIC_API_KEY` 추가

---

## ⚙️ API 엔드포인트

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/` | 프론트엔드 페이지 |
| POST | `/generate` | PPT 생성 |
| GET | `/health` | 서버 상태 확인 |

### POST /generate 파라미터
| 파라미터 | 타입 | 설명 |
|----------|------|------|
| file | File | 업로드할 파일 (txt, md, pdf 등) |
| slide_count | int | 슬라이드 수 (기본: 8) |
| language | string | Korean / English / Japanese |
| theme | string | dark / corporate / modern / warm |
| purpose | string | general / business / education / research |
