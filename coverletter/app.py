"""자기소개서 초안 작성 및 첨삭 웹앱.

실행:
    export ANTHROPIC_API_KEY=...
    pip install -r requirements.txt
    python app.py   # http://127.0.0.1:5000
"""
import json
import os
import re

import anthropic
from flask import Flask, jsonify, request, send_from_directory

MODEL = os.environ.get("CL_MODEL", "claude-sonnet-5-5")
MAX_INPUT_CHARS = 8000

app = Flask(__name__, static_folder="static", static_url_path="/static")
_client = None


def client():
    global _client
    if _client is None:
        _client = anthropic.Anthropic()  # ANTHROPIC_API_KEY 환경변수 사용
    return _client


COMMON_RULES = """\
[공통 원칙]
- 사용자가 제공한 경험·수치·사실만 사용한다. 없는 경력, 성과, 수치, 자격, 기관명을 지어내지 않는다.
- 근거가 부족한 부분은 임의로 채우지 말고 [확인 필요: 무엇이 필요한지] 형태로 표시한다.
- 실명, 주민등록번호, 전화번호, 주소, 이메일 등 직접 식별정보는 쓰지 않으며, 입력에 있어도 출력에서 제외한다.
- 특정 법령·정부 지침·제도 기준을 임의로 인용하지 않는다.
- 한국어 자기소개서 문체(~입니다/~했습니다)로 쓰고, 과장·상투 표현을 피하고 구체적 사례 중심으로 쓴다."""

DRAFT_SYSTEM = f"""너는 한국어 자기소개서 작성 코치다.
사용자의 경험 메모를 바탕으로 문항에 맞는 자기소개서 초안을 작성한다.
구조는 '결론(핵심 메시지) → 상황·과제 → 행동 → 결과 → 지원 직무와의 연결' 흐름을 따른다.
글자수 제한이 있으면 공백 포함 기준으로 제한의 90~100% 안에 맞춘다.

{COMMON_RULES}

출력은 아래 JSON 하나만 낸다(코드블록 없이).
{{"draft": "초안 본문", "missing": ["보완하면 좋을 정보 질문", ...]}}"""

REVIEW_SYSTEM = f"""너는 한국어 자기소개서 첨삭 전문가다.
사용자가 쓴 글의 의도와 사실관계는 유지하면서 구조, 구체성, 문장력, 직무 적합성을 점검한다.
사실을 새로 추가하지 않는다. 보강이 필요하면 revised 안에 [확인 필요: ...]로 표시한다.

{COMMON_RULES}

출력은 아래 JSON 하나만 낸다(코드블록 없이).
{{
  "summary": "총평 2~3문장",
  "strengths": ["잘된 점", ...],
  "issues": [{{"quote": "원문 일부", "problem": "문제점", "suggestion": "수정 방향"}}],
  "revised": "첨삭본 전문"
}}"""


def clip(value):
    return (value or "").strip()[:MAX_INPUT_CHARS]


def ask(system, user_text):
    resp = client().messages.create(
        model=MODEL,
        max_tokens=4000,
        system=system,
        messages=[{"role": "user", "content": user_text}],
    )
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"raw": text}


def with_length(result, key):
    body = result.get(key)
    if isinstance(body, str):
        result["length"] = len(body)
        result["length_no_space"] = len(re.sub(r"\s", "", body))
    return result


@app.get("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.post("/api/draft")
def draft():
    d = request.get_json(silent=True) or {}
    question, memo = clip(d.get("question")), clip(d.get("memo"))
    if not question or not memo:
        return jsonify(error="문항과 경험 메모를 입력해 주세요."), 400
    prompt = (
        f"[지원 분야/직무] {clip(d.get('job')) or '미입력'}\n"
        f"[문항] {question}\n"
        f"[글자수 제한(공백 포함)] {clip(d.get('limit')) or '없음'}\n"
        f"[경험 메모]\n{memo}"
    )
    try:
        return jsonify(with_length(ask(DRAFT_SYSTEM, prompt), "draft"))
    except anthropic.APIError as e:
        return jsonify(error=f"API 오류: {e}"), 502


@app.post("/api/review")
def review():
    d = request.get_json(silent=True) or {}
    text = clip(d.get("text"))
    if not text:
        return jsonify(error="첨삭할 글을 입력해 주세요."), 400
    prompt = (
        f"[지원 분야/직무] {clip(d.get('job')) or '미입력'}\n"
        f"[문항] {clip(d.get('question')) or '미입력'}\n"
        f"[글자수 제한(공백 포함)] {clip(d.get('limit')) or '없음'}\n"
        f"[원문]\n{text}"
    )
    try:
        result = ask(REVIEW_SYSTEM, prompt)
    except anthropic.APIError as e:
        return jsonify(error=f"API 오류: {e}"), 502
    result["original_length"] = len(text)
    return jsonify(with_length(result, "revised"))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5000)))
