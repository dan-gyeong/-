# 자기소개서 도우미

경험 메모로 자기소개서 **초안을 작성**하고, 작성한 글을 **첨삭**하는 웹앱 (Flask + Claude API).

## 실행
```
cd coverletter
pip install -r requirements.txt
export ANTHROPIC_API_KEY=발급받은_키
python app.py        # http://127.0.0.1:5000
```
- 모델 변경: `CL_MODEL` 환경변수 (기본 `claude-sonnet-5-5`)
- API 키는 서버에서만 사용되며 브라우저로 전달되지 않음

## 동작 원칙 (시스템 프롬프트에 반영)
- 입력된 경험·수치만 사용, 근거 부족 시 `[확인 필요: ...]`로 표시
- 이름·연락처 등 직접 식별정보와 임의 법령·제도 인용 금지
- 글자수(공백 포함/제외)와 제한 초과 여부 표시

## API
- `POST /api/draft` — `{job, question, limit, memo}` → `{draft, missing[], length, length_no_space}`
- `POST /api/review` — `{job, question, limit, text}` → `{summary, strengths[], issues[], revised, ...}`
