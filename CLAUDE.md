# Claude Code 프로젝트 지침

> 이 파일은 Claude Code가 참조하는 글로벌 코딩 표준입니다.
> ~/.claude/CLAUDE.md에 복사하면 모든 프로젝트에서 자동 적용됩니다.
>
> **상세 규칙**: `CODING_STANDARDS.md` 참조

---

## 핵심 원칙

1. **한국어 우선**: 주석, 커밋 메시지, 문서는 한국어
2. **CSS는 CSS로**: 인라인 스타일 금지, 클래스만 토글
3. **설정은 한 곳에**: JS는 `CONFIG`, Python은 `settings`
4. **src/ 구조 강제**: Python은 항상 `src/` 패키지 구조

---

## CSS 규칙

```
필수:
- 인라인 스타일 절대 금지
- CSS 변수 사용 (variables.css)
- 클래스 기반 스타일링

파일 구조:
css/
├── variables.css   # 토큰만 (색상, 간격, 폰트)
├── base.css        # 리셋, 기본 스타일
├── components.css  # 버튼, 카드, 토스트 등
└── utilities.css   # .hidden, .text-*, .flex 등
```

### 색상 팔레트
```css
--color-primary: #667eea;
--color-success: #10b981;
--color-warning: #f59e0b;
--color-danger: #ef4444;
```

---

## JavaScript 규칙

```javascript
// 설정은 CONFIG 객체로 (freeze 필수)
const CONFIG = {
  API: { BASE_URL: '...', TIMEOUT: 10000 },
  UI: { TOAST_DURATION: 3000 }
};
Object.freeze(CONFIG);

// 로그는 디버그 모드에서만
if (CONFIG.FEATURES.DEBUG) console.log(...);

// UI는 클래스 토글만
element.classList.add('toast', 'toast-success');  // ✓
element.style.backgroundColor = 'red';             // ✗ 금지
```

---

## Python 규칙

```
프로젝트 구조 (필수):
project/
├── src/
│   ├── __init__.py
│   ├── main.py      # FastAPI 앱
│   ├── config.py    # pydantic-settings
│   └── api/
├── tests/
├── requirements.txt
├── requirements-dev.txt
├── pyproject.toml   # ruff, black 설정
└── .env.example     # 환경 변수 예시

실행:
uvicorn src.main:app --reload
```

### FastAPI 필수 패턴
```python
# 예외 핸들러는 JSONResponse 사용
from fastapi.responses import JSONResponse

@app.exception_handler(404)
async def not_found(request: Request, exc) -> JSONResponse:
    return JSONResponse(status_code=404, content={...})
```

### pydantic-settings (v2)
```python
from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env")
```

---

## 보안 규칙

```
절대 금지:
- .env 파일 커밋
- API 키 코드에 하드코딩
- SECRET_KEY 기본값 그대로 배포

필수:
- .env.example 생성
- .gitignore에 .env 포함
- 배포 전 SECRET_KEY 변경
```

---

## 린트/포맷터 (기본 적용)

```bash
# Python
ruff check .        # 린트
black .             # 포맷

# JS (선택)
eslint .
prettier --write .
```

---

## Git 커밋

```
<타입>: <제목> (한국어)

타입: feat, fix, docs, style, refactor, test, chore
```

---

## 체크리스트

새 프로젝트 시작 시:
- [ ] `src/` 구조 생성 (Python)
- [ ] CSS 파일 분리 (인라인 0)
- [ ] `.env.example` 생성
- [ ] `pyproject.toml` 설정 (ruff, black)
- [ ] `.gitignore` 확인
