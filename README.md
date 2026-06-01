# DeckSpec OOXML PPTX Renderer

DeckSpec JSON을 입력받아 PowerPoint 내부 언어인 OOXML 파일들을 직접 만들고, 이를 `.pptx`로 패키징하는 Python 스크립트입니다.

LLM은 발표 내용과 디자인 의도를 `DeckSpec JSON`으로 만들고, 이 스크립트는 JSON을 안정적인 OOXML/PPTX로 렌더링하는 역할을 합니다.

## Files

```text
deckspec-ooxml-pptx/
├── make_ppt.py
├── base_ooxml/
├── template_profiles.json
├── README.md
├── 프롬프트.txt
├── requirements.txt
└── test/
    ├── deckspec_template.json
    └── ai_harness_deckspec.json
```

## Requirements

Python 3.9 이상을 권장합니다.

외부 패키지는 필요 없습니다. `requirements.txt`는 의존성이 없음을 명시하기 위해 포함되어 있습니다.

```bash
python3 --version
```

## Runtime Template

`base_ooxml/`은 필수 런타임 템플릿입니다.

이 폴더에는 PowerPoint가 안정적으로 여는 기본 OOXML 골격이 들어 있습니다.

```text
base_ooxml/
└── ppt/
    ├── theme/
    ├── slideLayouts/
    ├── slideMasters/
    ├── notesMasters/
    ├── presProps.xml
    ├── viewProps.xml
    └── tableStyles.xml
```

`make_ppt.py`는 매번 새 슬라이드 XML을 생성한 뒤, `base_ooxml/`의 안정적인 마스터/레이아웃/테마/프레젠테이션 속성 파일과 함께 `.pptx`로 패키징합니다.

`base_ooxml/`을 삭제하면 스크립트의 최소 OOXML fallback이 동작할 수 있지만, PowerPoint에서 복구 경고가 뜰 가능성이 커집니다. 실사용 배포에는 `base_ooxml/`을 반드시 포함하세요.

### Does `base_ooxml/` Lock The Design?

아니요. `base_ooxml/`은 디자인을 고정하는 용도라기보다, PowerPoint가 신뢰하는 문서 골격을 제공하는 용도입니다.

실제 슬라이드의 디자인과 패턴은 주로 아래에서 결정됩니다.

```text
DeckSpec JSON
├── design.theme
├── slide.layout
├── components
└── bullets / metrics / chart data

make_ppt.py
├── theme presets
├── background motifs
├── metric card renderer
├── bar chart renderer
└── slide layout renderer
```

즉 `base_ooxml/`은 문서의 뼈대이고, 슬라이드별 배경 도형, 카드, 막대그래프, 색상, 텍스트, 레이아웃은 `make_ppt.py`가 생성하는 `ppt/slides/slideN.xml`에서 계속 바뀝니다.

다만 `base_ooxml/` 안의 `theme1.xml`, `slideMaster1.xml`, `slideLayout1.xml`이 기본 폰트/테마/마스터 정보를 제공하므로, 아주 깊은 수준의 PowerPoint 기본 스타일은 어느 정도 영향을 줄 수 있습니다. 더 강한 브랜드 스타일이 필요하면 `base_ooxml/`을 여러 개 두고 선택하게 만들 수 있습니다.

예:

```text
base_ooxml/
base_ooxml_modern/
base_ooxml_editorial/
base_ooxml_corporate/
```

그리고 DeckSpec에서 이렇게 고르게 만들 수 있습니다.

```json
{
  "design": {
    "base_template": "base_ooxml_modern",
    "theme": "canva_modern_pitch"
  }
}
```

현재 버전은 단일 안정 템플릿 `base_ooxml/`을 사용하고, 시각적 차이는 `design.theme`, `slide.layout`, `components`로 만듭니다.

## Style Template Profiles

`template_profiles.json`은 스타일 템플릿 후보 목록입니다.

이 파일은 PPTX 원본을 한 번 분석해서 얻은 대표 색상, 배경 모티프, bullet 스타일, 레이아웃 변형 정보를 저장합니다. 런타임에는 PPTX 원본 파일이 필요하지 않습니다.

기본 선택 방식은 `auto_random`입니다.

```json
{
  "design": {
    "theme": "data_report",
    "template": "auto_random"
  }
}
```

동작 방식:

```text
template_profiles.json
├── 대표 색상
├── 배경 모티프
├── bullet 스타일
├── layout_variants
└── 현재 DeckSpec 슬라이드에 적용
```

`layout_variants`는 같은 DeckSpec 레이아웃을 템플릿마다 다르게 그리도록 만듭니다. 예를 들어 `metric_dashboard`는 기본 카드형, 스코어보드형, 세로 테이블형, 스태거 카드형으로 달라질 수 있고, `bar_comparison`은 가로 막대, 세로 컬럼, lollipop 차트로 달라질 수 있습니다.

특정 템플릿을 고정하고 싶으면 프로필의 `name` 값을 넣습니다.

```json
{
  "design": {
    "theme": "data_report",
    "template": "파워포인트-템플릿-원본-파일-다운로드-받기-free-ppt-template-2028"
  }
}
```

템플릿 후보 사용을 끄고 싶으면 아래처럼 둡니다.

```json
{
  "design": {
    "theme": "data_report",
    "template": "none"
  }
}
```

## Quick Start

템플릿 JSON으로 PPTX를 생성합니다.

```bash
python3 make_ppt.py test/deckspec_template.json -o output.pptx
```

생성되는 내부 OOXML 파일도 함께 보고 싶으면 `--dump-ooxml`을 사용합니다.

```bash
python3 make_ppt.py test/deckspec_template.json -o output.pptx --dump-ooxml output_ooxml
```

## Input Methods

### 1. JSON File

```bash
python3 make_ppt.py test/deckspec_template.json -o output.pptx
```

### 2. stdin

```bash
cat test/deckspec_template.json | python3 make_ppt.py - -o output.pptx
```

### 3. JSON String

```bash
python3 make_ppt.py --json-string '{"deck_title":"Demo","design":{"theme":"canva_fresh_startup"},"slides":[{"layout":"takeaway","kicker":"Demo","title":"JSON 문자열 입력","bullets":["파일 없이 바로 생성","OOXML을 직접 만든 뒤 PPTX로 패키징"],"speaker_note":"demo"}]}' -o output.pptx
```

## Write A Fresh Template

내장 샘플 DeckSpec을 새 JSON 템플릿 파일로 씁니다.

```bash
python3 make_ppt.py --write-template test/deckspec_template.json
```

## Validate The PPTX Package

생성된 PPTX가 zip 패키지로 정상인지 확인합니다.

```bash
unzip -t output.pptx
```

macOS에서 QuickLook 썸네일 렌더링을 확인할 수도 있습니다.

```bash
qlmanage -t -s 1200 -o . output.pptx
```

## DeckSpec Shape

기본 구조는 다음과 같습니다.

```json
{
  "deck_title": "Deck title",
  "subtitle": "Optional subtitle",
  "design": {
    "theme": "policy_brief",
    "template": "auto_random",
    "tone": "calm_analytical",
    "density": "medium",
    "visual_style": "clean_data_brief"
  },
  "slides": [
    {
      "layout": "title_summary",
      "kicker": "Section label",
      "title": "Slide title",
      "bullets": ["Short bullet", "Short bullet"],
      "speaker_note": "Optional speaker note"
    }
  ]
}
```

## Supported Themes

Use one of these values in `design.theme`.

```text
auto
policy_brief
executive_summary
data_report
canva_modern_pitch
canva_warm_editorial
canva_fresh_startup
crisis_brief
tech_architecture
strategy_board
weather_risk
```

The `canva_*` themes are not copied from Canva templates. They are Canva-like presentation style presets implemented with OOXML shapes, colors, and background motifs.

`auto` lets the renderer choose a theme from deck content. For example, incident/safety text tends to select `crisis_brief`, weather-risk text tends to select `weather_risk`, AI/system architecture text tends to select `tech_architecture`, and CEO/board text tends to select `executive_summary`.

## Supported Slide Layouts

Use one of these values in each slide's `layout`.

```text
title_cover
title_summary
bullets
metric_dashboard
bar_comparison
line_trend
timeline
process_flow
comparison
risk_matrix
cause_effect
architecture_map
callout_focus
takeaway
```

## Supported Components

### Metric Card

Used with `metric_dashboard`.

```json
{
  "type": "metric_card",
  "label": "출생아 수",
  "value": "2만6916명",
  "delta": "+11.7%",
  "context": "전년 동월 대비",
  "emphasis": "primary"
}
```

### Bar Chart

Used with `bar_comparison`.

```json
{
  "type": "bar_chart",
  "title": "전년 동월 대비 증가율",
  "unit": "%",
  "data": [
    { "label": "출생아 수", "value": 11.7 },
    { "label": "혼인 건수", "value": 12.4 }
  ],
  "emphasis": "accent"
}
```

### Line Chart

Used with `line_trend`.

```json
{
  "type": "line_chart",
  "title": "사망자 수 기준 사고 이력",
  "unit": "명",
  "data": [
    { "label": "2018", "value": 5 },
    { "label": "2019", "value": 3 },
    { "label": "2026", "value": 5 }
  ],
  "emphasis": "warning"
}
```

### Timeline

Used with `timeline`.

```json
{
  "type": "timeline",
  "events": [
    { "time": "10:59", "label": "폭발 발생", "detail": "56동 세척 공실" },
    { "time": "11:49", "label": "초진", "detail": "50분 만에 진압" }
  ]
}
```

### Process Flow

Used with `process_flow`.

```json
{
  "type": "process_flow",
  "steps": [
    { "label": "현장 통제", "detail": "붕괴 위험과 접근 제한" },
    { "label": "원인 조사", "detail": "세척 공정 폭발 추정" }
  ]
}
```

### Risk Matrix

Used with `risk_matrix`. `likelihood` and `impact` use `0` to `1`.

```json
{
  "type": "risk_matrix",
  "items": [
    { "label": "세척공정", "likelihood": 0.82, "impact": 0.9, "emphasis": "warning" },
    { "label": "점검사각", "likelihood": 0.68, "impact": 0.78, "emphasis": "accent" }
  ]
}
```

### Cause Effect

Used with `cause_effect`.

```json
{
  "type": "cause_effect",
  "causes": ["화약 묻은 공구 세척", "작은 건물 점검 사각"],
  "events": ["56동 세척 공실 폭발", "화재와 구조 지연"],
  "effects": ["근로자 7명 사상", "공정 안전 재검토 필요"]
}
```

### Architecture Map

Used with `architecture_map`.

```json
{
  "type": "architecture_map",
  "nodes": [
    { "id": "user", "label": "User" },
    { "id": "harness", "label": "Harness", "emphasis": "primary" },
    { "id": "model", "label": "Model" },
    { "id": "tools", "label": "Tools" }
  ],
  "edges": [
    { "from": "user", "to": "harness" },
    { "from": "harness", "to": "model" },
    { "from": "harness", "to": "tools" }
  ]
}
```

### Callout

Used with `callout_focus`.

```json
{
  "type": "callout",
  "headline": "복구보다 중요한 것은 반복 패턴의 차단",
  "body": "핵심 메시지를 한 문장 중심으로 크게 배치합니다."
}
```

## LLM Prompt For DeckSpec JSON

다른 세션에서 LLM에게 DeckSpec JSON을 만들게 할 때는 아래 프롬프트를 사용하세요.

LLM은 원문을 요약하고, 슬라이드 장수와 디자인을 판단한 뒤, 이 스크립트가 렌더링할 수 있는 JSON만 반환해야 합니다.

```text
You are a senior presentation strategist and information designer.

Convert the source text into a DeckSpec JSON for the provided Python OOXML PPTX renderer.

Return valid JSON only.
Do not return markdown.
Do not explain your choices outside the JSON.

Rules:
- Create 3 to 6 slides unless the source text clearly requires fewer or more.
- Preserve important source numbers exactly.
- Do not invent facts, sources, dates, numbers, company names, or claims.
- Use concise Korean business presentation language.
- Keep each bullet short enough for a slide.
- Choose the most suitable design.theme from the allowed theme list. Use "auto" when the source text gives enough clues for content-aware theme selection.
- Choose each slide.layout from the allowed layout list.
- Use metric_card components for important headline numbers.
- Use bar_chart components when comparing numeric values.
- Use line_chart components for time-series or repeated incident history.
- Use timeline components when exact event order matters.
- Use process_flow components for operational steps.
- Use risk_matrix components when likelihood and impact should be compared.
- Use cause_effect components when causes, event, and effects should be separated.
- Use architecture_map components for systems, agents, tools, model routing, or harness/A2A explanations.
- Use callout components for a strong single-message slide.
- Use bullets when the slide is mostly explanatory.
- Use takeaway for the final executive conclusion.
- If a layout requires components, include components.
- If a slide has components, it may also include bullets for supporting context.
- Keep slide titles under 34 Korean characters when possible.
- Keep bullets short. The renderer will cap overflowing bullets, but the JSON should still be slide-friendly.

Allowed design.theme values:
- auto: content-aware theme selection by renderer
- policy_brief: sober public-policy or research brief
- executive_summary: CEO/board-level concise business summary
- data_report: analytical report with numbers and comparisons
- canva_modern_pitch: dark modern pitch deck style
- canva_warm_editorial: warm editorial/report style
- canva_fresh_startup: fresh startup/product style
- crisis_brief: incident, safety, risk, accident, emergency, or audit brief
- tech_architecture: AI, software architecture, harness, A2A, agent systems
- strategy_board: strategy, roadmap, investment, transformation, board memo
- weather_risk: typhoon, heavy rain, heat wave, forecast, weather warning brief

Allowed design.template values:
- auto_random: randomly select one profile from template_profiles.json
- none: do not use template profiles
- a profile name from template_profiles.json: use that template profile

Allowed slide.layout values:
- title_cover
- title_summary
- bullets
- metric_dashboard
- bar_comparison
- line_trend
- timeline
- process_flow
- comparison
- risk_matrix
- cause_effect
- architecture_map
- callout_focus
- takeaway

Supported component types:
- metric_card
- bar_chart
- line_chart
- timeline
- process_flow
- comparison
- risk_matrix
- cause_effect
- architecture_map
- callout

DeckSpec JSON schema:
{
  "deck_title": "string",
  "subtitle": "string",
  "design": {
    "theme": "auto | policy_brief | executive_summary | data_report | canva_modern_pitch | canva_warm_editorial | canva_fresh_startup | crisis_brief | tech_architecture | strategy_board | weather_risk",
    "template": "auto_random | none | template profile name",
    "tone": "string",
    "density": "low | medium | high",
    "visual_style": "string"
  },
  "slides": [
    {
      "layout": "title_cover | title_summary | bullets | metric_dashboard | bar_comparison | line_trend | timeline | process_flow | comparison | risk_matrix | cause_effect | architecture_map | callout_focus | takeaway",
      "kicker": "string",
      "title": "string",
      "bullets": ["string"],
      "components": [
        {
          "type": "metric_card",
          "label": "string",
          "value": "string",
          "delta": "string",
          "context": "string",
          "emphasis": "primary | accent | warning"
        },
        {
          "type": "bar_chart",
          "title": "string",
          "unit": "string",
          "data": [
            { "label": "string", "value": 0 }
          ],
          "emphasis": "primary | accent | warning"
        },
        {
          "type": "line_chart | timeline | process_flow | comparison | risk_matrix | cause_effect | architecture_map | callout",
          "data": [],
          "items": [],
          "events": [],
          "steps": [],
          "nodes": [],
          "edges": []
        }
      ],
      "speaker_note": "string"
    }
  ]
}

Layout guidance:
- Put a title_cover slide as the first item in slides when the deck needs a cover.
- Use title_summary for the first content slide after the cover.
- Use metric_dashboard when the source has 2-3 important numbers.
- Use bar_comparison when two or more numeric values should be compared.
- Use line_trend for time-series, repeated incidents, or year-by-year values.
- Use timeline for chronological event handling.
- Use process_flow for operational or investigation steps.
- Use comparison for two-sided decisions, before/after, or immediate/structural split.
- Use risk_matrix for likelihood-impact prioritization.
- Use cause_effect for root-cause framing.
- Use architecture_map for AI systems, harnesses, agents, tools, memory, or model routing.
- Use callout_focus when a single conclusion should dominate the slide.
- Use bullets for compact explanation when no visual pattern fits.
- Use takeaway for the final conclusion.

Source text:
<<<
PASTE_SOURCE_TEXT_HERE
>>>
```

## Theme Selection Guide

LLM에게 테마를 고르게 할 때는 아래 기준을 쓰면 됩니다.

```text
policy_brief:
정부, 정책, 리서치, 공공 데이터, 차분한 보고서

executive_summary:
CEO 보고, 이사회 보고, 경영진 의사결정 자료

data_report:
숫자, 비교, 지표, 성과 분석 중심 보고서

canva_modern_pitch:
스타트업 피치, AI/테크, 강한 첫인상, 어두운 모던 톤

canva_warm_editorial:
브랜드 스토리, 교육, 사회 이슈, 따뜻한 리포트 톤

canva_fresh_startup:
제품 소개, 성장 전략, 젊고 밝은 스타트업 톤

crisis_brief:
사고, 안전, 인명 피해, 리스크, 감사, 수습, 조사 보고

tech_architecture:
AI 에이전트, 하네스, A2A, 모델 라우팅, 소프트웨어 구조

strategy_board:
전략, 로드맵, 투자, 전환 과제, 임원 의사결정

weather_risk:
태풍, 폭우, 폭염, 기상 특보, 지역별 날씨 위험 브리프
```

## Example LLM Output

```json
{
  "deck_title": "AI 하네스 엔지니어링의 부상",
  "subtitle": "AI 모델보다 운용 체계 설계가 성과를 좌우하는 시대로 전환",
  "design": {
    "theme": "executive_summary",
    "tone": "calm_analytical",
    "density": "medium",
    "visual_style": "board_brief"
  },
  "slides": [
    {
      "layout": "title_summary",
      "kicker": "핵심 개념",
      "title": "하네스는 AI 에이전트의 운용 체계",
      "bullets": [
        "AI 에이전트를 감싸는 제어 구조",
        "규칙·도구·샌드박스·피드백을 포함",
        "모델이 아닌 일하는 환경 설계가 핵심"
      ],
      "speaker_note": "하네스를 AI 에이전트가 안전하고 예측 가능하게 작동하도록 만드는 운용 인프라로 정의한다."
    },
    {
      "layout": "metric_dashboard",
      "kicker": "성과 신호",
      "title": "하네스 설계가 생산성 차이를 만든다",
      "components": [
        {
          "type": "metric_card",
          "label": "제품 코드",
          "value": "100만 줄",
          "delta": "5개월",
          "context": "오픈AI 사례",
          "emphasis": "primary"
        },
        {
          "type": "metric_card",
          "label": "PR 처리",
          "value": "1500개",
          "delta": "5개월간",
          "context": "하네스 기반 개발 흐름",
          "emphasis": "accent"
        },
        {
          "type": "metric_card",
          "label": "개발자 통합률",
          "value": "60%",
          "delta": "이미 업무 통합",
          "context": "AI 활용 확산",
          "emphasis": "warning"
        }
      ],
      "speaker_note": "중요한 숫자를 카드형 지표로 강조한다."
    }
  ]
}
```

## Notes

- This renderer does not call an LLM.
- This renderer does not use `python-pptx`.
- It creates OOXML strings directly, writes files such as `ppt/slides/slide1.xml`, and packages them with Python's standard `zipfile`.
- For production use, open the generated PPTX in PowerPoint, Keynote, or LibreOffice as a final compatibility check.
