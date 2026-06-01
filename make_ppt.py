import argparse
import html
import json
import os
import random
import re
import sys
import zipfile
from pathlib import Path


DEFAULT_OUTPUT = "deck_from_deckspec_ooxml.pptx"
BASE_OOXML_DIR = Path(__file__).with_name("base_ooxml")
TEMPLATE_PROFILES_PATH = Path(__file__).with_name("template_profiles.json")

EMU_PER_INCH = 914400
SLIDE_W = 12192000
SLIDE_H = 6858000

THEMES = {
    "policy_brief": {
        "background": "FFFFFF",
        "surface": "F8FAFC",
        "surface_alt": "EEF2FF",
        "text": "111827",
        "muted": "6B7280",
        "line": "E5E7EB",
        "primary": "2563EB",
        "accent": "0F766E",
        "warning": "F59E0B",
        "motif": "top_band",
        "bullet_style": "cards",
    },
    "executive_summary": {
        "background": "FFFFFF",
        "surface": "F9FAFB",
        "surface_alt": "ECFDF5",
        "text": "0F172A",
        "muted": "64748B",
        "line": "CBD5E1",
        "primary": "0F766E",
        "accent": "2563EB",
        "warning": "D97706",
        "motif": "side_panel",
        "bullet_style": "numbered",
    },
    "data_report": {
        "background": "FFFFFF",
        "surface": "F8FAFC",
        "surface_alt": "F0F9FF",
        "text": "111827",
        "muted": "64748B",
        "line": "D1D5DB",
        "primary": "0284C7",
        "accent": "7C3AED",
        "warning": "EA580C",
        "motif": "data_blocks",
        "bullet_style": "data_rows",
    },
    "canva_modern_pitch": {
        "background": "0F172A",
        "surface": "1E293B",
        "surface_alt": "312E81",
        "text": "F8FAFC",
        "muted": "CBD5E1",
        "line": "334155",
        "primary": "38BDF8",
        "accent": "A78BFA",
        "warning": "FBBF24",
        "motif": "diagonal_blocks",
        "bullet_style": "outline",
    },
    "canva_warm_editorial": {
        "background": "FFF7ED",
        "surface": "FFFFFF",
        "surface_alt": "FFEDD5",
        "text": "1C1917",
        "muted": "78716C",
        "line": "FED7AA",
        "primary": "C2410C",
        "accent": "0F766E",
        "warning": "CA8A04",
        "motif": "editorial_frame",
        "bullet_style": "editorial",
    },
    "canva_fresh_startup": {
        "background": "F0FDFA",
        "surface": "FFFFFF",
        "surface_alt": "CCFBF1",
        "text": "134E4A",
        "muted": "0F766E",
        "line": "99F6E4",
        "primary": "0D9488",
        "accent": "2563EB",
        "warning": "F59E0B",
        "motif": "soft_circles",
        "bullet_style": "cards",
    },
    "crisis_brief": {
        "background": "FFFFFF",
        "surface": "F8FAFC",
        "surface_alt": "FEF2F2",
        "text": "111827",
        "muted": "6B7280",
        "line": "D1D5DB",
        "primary": "B91C1C",
        "accent": "EA580C",
        "warning": "DC2626",
        "motif": "alert_band",
        "bullet_style": "alert",
    },
    "tech_architecture": {
        "background": "08111F",
        "surface": "111827",
        "surface_alt": "172554",
        "text": "F8FAFC",
        "muted": "CBD5E1",
        "line": "334155",
        "primary": "38BDF8",
        "accent": "22C55E",
        "warning": "FBBF24",
        "motif": "blueprint",
        "bullet_style": "outline",
    },
    "strategy_board": {
        "background": "FFFFFF",
        "surface": "F8FAFC",
        "surface_alt": "F5F3FF",
        "text": "111827",
        "muted": "4B5563",
        "line": "D1D5DB",
        "primary": "312E81",
        "accent": "B45309",
        "warning": "B91C1C",
        "motif": "executive_frame",
        "bullet_style": "numbered",
    },
    "weather_risk": {
        "background": "F8FBFF",
        "surface": "FFFFFF",
        "surface_alt": "E0F2FE",
        "text": "0F172A",
        "muted": "475569",
        "line": "BAE6FD",
        "primary": "0284C7",
        "accent": "0F766E",
        "warning": "F97316",
        "motif": "weather_front",
        "bullet_style": "weather",
    },
}

TEST_DECKSPEC = {
    "deck_title": "1월 출생아·혼인 지표 반등",
    "subtitle": "출생아 수와 합계출산율이 개선되고 혼인 건수도 최장 증가세를 기록",
    "design": {
        "theme": "policy_brief",
        "tone": "calm_analytical",
        "density": "medium",
        "visual_style": "clean_data_brief",
    },
    "slides": [
        {
            "layout": "title_summary",
            "kicker": "인구동향 요약",
            "title": "출생아 수와 혼인이 동시에 반등",
            "bullets": [
                "1월 출생아 수는 2만6916명으로 7년 만에 최대",
                "합계출산율은 0.99명으로 월간 집계 이후 최고",
                "혼인 건수는 22개월 연속 증가",
            ],
            "speaker_note": "출생과 혼인이 함께 개선된 점을 첫 장에서 요약한다.",
        },
        {
            "layout": "metric_dashboard",
            "kicker": "핵심 수치",
            "title": "출생·혼인 지표가 함께 개선",
            "components": [
                {
                    "type": "metric_card",
                    "label": "출생아 수",
                    "value": "2만6916명",
                    "delta": "+11.7%",
                    "context": "전년 동월 대비",
                    "emphasis": "primary",
                },
                {
                    "type": "metric_card",
                    "label": "합계출산율",
                    "value": "0.99명",
                    "delta": "월간 최대",
                    "context": "2024년 1월 이후",
                    "emphasis": "accent",
                },
                {
                    "type": "metric_card",
                    "label": "혼인 건수",
                    "value": "2만2640건",
                    "delta": "+12.4%",
                    "context": "22개월 연속 증가",
                    "emphasis": "primary",
                },
            ],
            "speaker_note": "핵심 지표를 카드형 컴포넌트로 요약한다.",
        },
        {
            "layout": "bar_comparison",
            "kicker": "증가율 비교",
            "title": "혼인 증가가 출생 흐름의 선행 신호",
            "components": [
                {
                    "type": "bar_chart",
                    "title": "전년 동월 대비 증가율",
                    "unit": "%",
                    "data": [
                        {"label": "출생아 수", "value": 11.7},
                        {"label": "혼인 건수", "value": 12.4},
                    ],
                    "emphasis": "accent",
                }
            ],
            "bullets": ["출생아 수는 19개월 연속 증가", "혼인 건수는 22개월 연속 증가"],
            "speaker_note": "수치 비교는 막대 그래프로 보여준다.",
        },
        {
            "layout": "takeaway",
            "kicker": "Executive Takeaway",
            "title": "반등 신호는 분명하지만 지속성 확인이 핵심",
            "bullets": [
                "출생아 수, 합계출산율, 혼인 건수가 동시에 개선",
                "2차 에코붐 세대와 정책 효과가 겹친 국면",
                "구조적 회복 여부는 혼인 증가세 지속이 관건",
            ],
            "speaker_note": "결론은 낙관 신호와 추세 확인 필요성을 함께 담는다.",
        },
    ],
}


def emu(inches):
    return int(inches * EMU_PER_INCH)


def xml_escape(value):
    return html.escape(str(value), quote=True)


def safe_text(value, default=""):
    if value is None:
        return default
    return str(value)


def clamp_text(value, limit):
    text = safe_text(value).strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def load_template_profiles():
    if not TEMPLATE_PROFILES_PATH.exists():
        return []
    data = json.loads(TEMPLATE_PROFILES_PATH.read_text(encoding="utf-8"))
    profiles = data.get("profiles", []) if isinstance(data, dict) else []
    return [profile for profile in profiles if isinstance(profile, dict)]


def select_template_profile(deckspec):
    profiles = load_template_profiles()
    if not profiles:
        return None
    design = deckspec.get("design", {}) if isinstance(deckspec, dict) else {}
    request = safe_text(design.get("template", "auto_random")).strip()
    if request in {"", "none", "off", "false"}:
        return None
    if request in {"auto", "auto_random", "random"}:
        return random.choice(profiles)
    for profile in profiles:
        if request in {safe_text(profile.get("name")), safe_text(profile.get("source_file"))}:
            return profile
    return None


def deck_text(deckspec):
    parts = [
        safe_text(deckspec.get("deck_title", "")),
        safe_text(deckspec.get("subtitle", "")),
        safe_text(deckspec.get("design", {}).get("visual_style", "")),
        safe_text(deckspec.get("design", {}).get("tone", "")),
    ]
    for slide in deckspec.get("slides", []):
        parts.extend([safe_text(slide.get("kicker", "")), safe_text(slide.get("title", ""))])
        parts.extend(safe_text(b) for b in slide.get("bullets", []))
        for component in slide.get("components", []):
            parts.append(json.dumps(component, ensure_ascii=False))
    return " ".join(parts).lower()


def infer_theme_name(deckspec):
    text = deck_text(deckspec)
    weather_terms = ["태풍", "기상", "폭우", "폭염", "강수", "비구름", "풍랑", "날씨", "기온", "제주", "남부"]
    incident_terms = ["사고", "폭발", "사망", "부상", "피해", "화재", "수습", "조사", "위험"]
    startup_terms = ["제품", "성장", "사용자", "스타트업", "런칭", "고객", "시장"]
    pitch_terms = ["ai", "테크", "플랫폼", "투자", "피치", "전략", "엔지니어링"]
    policy_terms = ["정부", "정책", "공공", "인구", "통계", "출생", "혼인", "데이터"]
    executive_terms = ["ceo", "경영진", "이사회", "의사결정", "운용 체계", "리스크"]

    if any(term in text for term in weather_terms):
        return "weather_risk"
    if any(term in text for term in incident_terms):
        return "crisis_brief"
    if any(term in text for term in startup_terms):
        return "canva_fresh_startup"
    if any(term in text for term in pitch_terms):
        return "tech_architecture"
    if any(term in text for term in executive_terms):
        return "executive_summary"
    if any(term in text for term in policy_terms):
        return "policy_brief"
    return "executive_summary"


def theme_for(deckspec):
    design = deckspec.get("design", {}) if isinstance(deckspec, dict) else {}
    theme_name = design.get("theme", "policy_brief")
    visual_style = safe_text(design.get("visual_style", "")).lower()
    if theme_name == "policy_brief" and any(term in visual_style for term in ["weather", "기상", "typhoon", "risk_brief"]):
        theme_name = "weather_risk"
    if theme_name in ("auto", "content_aware", "content-aware", ""):
        theme_name = infer_theme_name(deckspec)
    theme = dict(THEMES.get(theme_name, THEMES["policy_brief"]))
    profile = select_template_profile(deckspec)
    if profile:
        style_keys = {"primary", "accent", "warning", "surface_alt", "line", "motif", "bullet_style", "layout_variants"}
        theme.update({k: v for k, v in profile.items() if k in style_keys})
        design["selected_template"] = profile["name"]
        deckspec["design"] = design
    return theme


def theme_color(theme, key):
    return theme.get(key, theme["primary"])


def layout_variant(theme, layout, default="default"):
    variants = theme.get("layout_variants", {})
    if isinstance(variants, dict):
        return variants.get(layout, default)
    return default


def srgb_fill_xml(color, alpha=None):
    if alpha is None:
        return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    return f'<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>'


def paragraph_xml(text, size=2200, color="374151", bold=False):
    bold_attr = ' b="1"' if bold else ""
    return (
        "<a:p>"
        '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>'
        "<a:r>"
        f'<a:rPr lang="ko-KR" sz="{size}"{bold_attr}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        '<a:latin typeface="Aptos"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Aptos"/>'
        "</a:rPr>"
        f"<a:t>{xml_escape(text)}</a:t>"
        "</a:r>"
        f'<a:endParaRPr lang="ko-KR" sz="{size}"/>'
        "</a:p>"
    )


def text_box_xml(shape_id, name, x, y, w, h, paragraphs, fill=None, line=None, radius=False):
    fill_xml = '<a:noFill/>' if fill is None else srgb_fill_xml(fill)
    line_xml = "<a:ln><a:noFill/></a:ln>" if line is None else f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    geom = (
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 9000"/></a:avLst></a:prstGeom>'
        if radius
        else '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    )
    if not paragraphs:
        return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          {geom}
          {fill_xml}
          {line_xml}
        </p:spPr>
      </p:sp>"""
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          {geom}
          {fill_xml}
          {line_xml}
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>
          <a:lstStyle/>
          {''.join(paragraphs)}
        </p:txBody>
      </p:sp>"""


def shape_xml(shape_id, name, x, y, w, h, fill, line=None, shape="rect"):
    line_xml = "<a:ln><a:noFill/></a:ln>" if line is None else f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          <a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>
          {srgb_fill_xml(fill)}
          {line_xml}
        </p:spPr>
      </p:sp>"""


def translucent_shape_xml(shape_id, name, x, y, w, h, fill, alpha=18000, shape="rect"):
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          <a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>
          {srgb_fill_xml(fill, alpha)}
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
      </p:sp>"""


def line_xml(shape_id, x, y, w, color="D1D5DB"):
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="Divider"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="0"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>"""


def line_segment_xml(shape_id, name, x1, y1, x2, y2, color, width=19050, dash="solid"):
    x = min(x1, x2)
    y = min(y1, y2)
    w = max(abs(x2 - x1), 1)
    h = max(abs(y2 - y1), 1)
    flip_h = ' flipH="1"' if x2 < x1 else ""
    flip_v = ' flipV="1"' if y2 < y1 else ""
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm{flip_h}{flip_v}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:noFill/>
          <a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:prstDash val="{dash}"/></a:ln>
        </p:spPr>
      </p:sp>"""


def as_number(value, default=0):
    if isinstance(value, (int, float)):
        return float(value)
    match = re.search(r"-?\d+(?:\.\d+)?", safe_text(value))
    return float(match.group(0)) if match else default


def add_background_motif(shapes, theme, shape_id):
    motif = theme.get("motif", "top_band")
    if motif == "side_panel":
        shapes.append(translucent_shape_xml(shape_id, "Side Panel", emu(0), emu(0), emu(1.15), SLIDE_H, theme["primary"], 17000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Corner Block", emu(10.9), emu(0), emu(2.45), emu(1.45), theme["accent"], 15000))
        shape_id += 1
    elif motif == "data_blocks":
        shapes.append(translucent_shape_xml(shape_id, "Data Block A", emu(9.95), emu(0.35), emu(2.35), emu(0.72), theme["primary"], 15000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Data Block B", emu(10.75), emu(1.18), emu(1.55), emu(0.48), theme["accent"], 15000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Data Block C", emu(0.0), emu(6.95), SLIDE_W, emu(0.18), theme["primary"], 10000))
        shape_id += 1
    elif motif == "diagonal_blocks":
        shapes.append(translucent_shape_xml(shape_id, "Diagonal Block A", emu(8.8), emu(0), emu(4.6), emu(2.1), theme["primary"], 18000, "parallelogram"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Diagonal Block B", emu(9.75), emu(1.05), emu(3.3), emu(1.65), theme["accent"], 16000, "parallelogram"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Bottom Band", emu(0), emu(6.88), SLIDE_W, emu(0.22), theme["primary"], 18000))
        shape_id += 1
    elif motif == "editorial_frame":
        shapes.append(translucent_shape_xml(shape_id, "Editorial Top", emu(0), emu(0), SLIDE_W, emu(0.22), theme["primary"], 22000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Editorial Side", emu(12.95), emu(0), emu(0.22), SLIDE_H, theme["accent"], 22000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Editorial Wash", emu(0.65), emu(5.72), emu(4.2), emu(1.25), theme["surface_alt"], 70000, "roundRect"))
        shape_id += 1
    elif motif == "soft_circles":
        shapes.append(translucent_shape_xml(shape_id, "Circle A", emu(10.2), emu(0.2), emu(2.2), emu(2.2), theme["primary"], 14000, "ellipse"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Circle B", emu(11.15), emu(1.25), emu(1.55), emu(1.55), theme["accent"], 15000, "ellipse"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Circle C", emu(-0.55), emu(6.0), emu(1.75), emu(1.75), theme["primary"], 12000, "ellipse"))
        shape_id += 1
    elif motif == "alert_band":
        shapes.append(translucent_shape_xml(shape_id, "Alert Top Band", emu(0), emu(0), SLIDE_W, emu(0.26), theme["warning"], 26000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Alert Side Rail", emu(0), emu(0), emu(0.18), SLIDE_H, theme["primary"], 18000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Alert Corner", emu(10.65), emu(0.2), emu(1.85), emu(1.85), theme["warning"], 12000, "triangle"))
        shape_id += 1
    elif motif == "blueprint":
        for x in [1.2, 3.6, 6.0, 8.4, 10.8]:
            shapes.append(line_segment_xml(shape_id, "Blueprint Grid V", emu(x), emu(0.3), emu(x), emu(6.9), theme["line"], 3175))
            shape_id += 1
        for y in [1.2, 2.4, 3.6, 4.8, 6.0]:
            shapes.append(line_segment_xml(shape_id, "Blueprint Grid H", emu(0.3), emu(y), emu(12.9), emu(y), theme["line"], 3175))
            shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Blueprint Glow", emu(10.4), emu(0.35), emu(1.8), emu(1.8), theme["primary"], 12000, "ellipse"))
        shape_id += 1
    elif motif == "executive_frame":
        shapes.append(translucent_shape_xml(shape_id, "Executive Top", emu(0), emu(0), SLIDE_W, emu(0.18), theme["primary"], 18000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Executive Accent", emu(0.65), emu(6.82), emu(3.8), emu(0.18), theme["accent"], 22000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Executive Corner", emu(11.5), emu(0.35), emu(0.8), emu(0.8), theme["accent"], 15000, "roundRect"))
        shape_id += 1
    elif motif == "weather_front":
        shapes.append(translucent_shape_xml(shape_id, "Sky Wash", emu(0), emu(0), SLIDE_W, emu(0.34), theme["primary"], 17000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Rain Mass", emu(8.95), emu(0), emu(4.6), emu(1.7), theme["surface_alt"], 42000, "rect"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Heat Mass", emu(10.2), emu(1.0), emu(2.75), emu(0.55), theme["warning"], 22000, "parallelogram"))
        shape_id += 1
        for x in [9.4, 9.95, 10.5, 11.05, 11.6, 12.15]:
            shapes.append(line_segment_xml(shape_id, "Rain Dash", emu(x), emu(0.46), emu(x - 0.16), emu(0.95), theme["primary"], 9525, "dash"))
            shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Weather Front", emu(0.7), emu(6.78), emu(12.45), emu(6.78), theme["primary"], 12700, "dash"))
        shape_id += 1
    else:
        shapes.append(translucent_shape_xml(shape_id, "Top Band", emu(0), emu(0), SLIDE_W, emu(0.22), theme["primary"], 17000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Top Accent", emu(8.8), emu(0), emu(4.55), emu(0.62), theme["accent"], 12000))
        shape_id += 1
    return shape_id


def add_header(shapes, slide, theme, shape_id):
    shapes.append(
        text_box_xml(
            shape_id,
            "Kicker",
            emu(0.65),
            emu(0.34),
            emu(5.6),
            emu(0.35),
            [paragraph_xml(clamp_text(slide.get("kicker", ""), 64), 1100, theme["primary"], True)],
        )
    )
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Title",
            emu(0.65),
            emu(0.72),
            emu(11.85),
            emu(0.72),
            [paragraph_xml(clamp_text(slide.get("title", "Untitled"), 56), 3000, theme["text"], True)],
        )
    )
    shape_id += 1
    shapes.append(line_xml(shape_id, emu(0.65), emu(1.55), emu(12.0), theme["line"]))
    return shape_id + 1


def add_note_and_page(shapes, slide, theme, index, shape_id):
    note = slide.get("speaker_note") or slide.get("note") or ""
    shapes.append(
        text_box_xml(
            shape_id,
            "Speaker Note",
            emu(0.8),
            emu(6.45),
            emu(10.8),
            emu(0.35),
            [paragraph_xml(clamp_text(note, 120), 1000, theme["muted"])],
        )
    )
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Page Number",
            emu(12.0),
            emu(6.45),
            emu(0.55),
            emu(0.35),
            [paragraph_xml(str(index), 1000, theme["muted"], True)],
        )
    )
    return shape_id + 1


def render_bullets(
    shapes,
    slide,
    theme,
    shape_id,
    start_y=1.95,
    max_items=None,
    bottom_y=6.15,
    box_h=0.62,
    step=0.78,
    font_size=1800,
    text_limit=80,
):
    y = start_y
    bullet_style = theme.get("bullet_style", "cards")
    fit_count = max(0, int((bottom_y - start_y + 0.001) // step))
    count = min(len(slide.get("bullets", [])), max_items if max_items is not None else 5, fit_count)
    if count == 0 and slide.get("bullets"):
        count = 1
        box_h = min(box_h, max(0.42, bottom_y - start_y))
    for idx, bullet in enumerate(slide.get("bullets", [])[:count]):
        if bullet_style == "outline":
            shapes.append(line_segment_xml(shape_id, "Bullet Guide", emu(0.88), emu(y + 0.32), emu(12.1), emu(y + 0.32), theme["line"], 6350, "dash"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Bullet Dot", emu(0.86), emu(y + 0.21), emu(0.18), emu(0.18), theme["primary"], theme["primary"], "ellipse"))
            shape_id += 1
            text_x, text_w, prefix = 1.18, 10.95, ""
        elif bullet_style in ("numbered", "data_rows"):
            shapes.append(
                text_box_xml(
                    shape_id,
                    "Bullet Background",
                    emu(0.8),
                    emu(y),
                    emu(11.75),
                    emu(box_h),
                    [],
                    fill=theme["surface"],
                    line=theme["line"],
                    radius=False,
                )
            )
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Bullet Number", emu(0.94), emu(y + 0.12), emu(0.34), emu(0.34), theme["primary"], theme["primary"], "rect"))
            shape_id += 1
            text_x, text_w, prefix = 1.42, 10.65, f"{idx + 1}. "
        elif bullet_style in ("alert", "weather"):
            fill = theme["surface_alt"] if idx == 0 else theme["surface"]
            accent = theme["warning"] if idx == 0 or bullet_style == "alert" else theme["primary"]
            shapes.append(
                text_box_xml(
                    shape_id,
                    "Bullet Background",
                    emu(0.8),
                    emu(y),
                    emu(11.75),
                    emu(box_h),
                    [],
                    fill=fill,
                    line=theme["line"],
                    radius=True,
                )
            )
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Bullet Accent", emu(0.8), emu(y), emu(0.08), emu(box_h), accent, accent))
            shape_id += 1
            text_x, text_w, prefix = 1.05, 11.25, "• "
        elif bullet_style == "editorial":
            shapes.append(shape_xml(shape_id, "Bullet Rule", emu(0.8), emu(y + 0.06), emu(0.08), emu(box_h - 0.12), theme["primary"], theme["primary"]))
            shape_id += 1
            text_x, text_w, prefix = 1.05, 11.25, ""
        else:
            shapes.append(
                text_box_xml(
                    shape_id,
                    "Bullet Background",
                    emu(0.8),
                    emu(y),
                    emu(11.75),
                    emu(box_h),
                    [],
                    fill=theme["surface"],
                    line=theme["line"],
                    radius=True,
                )
            )
            shape_id += 1
            text_x, text_w, prefix = 1.05, 11.25, "• "
        shapes.append(
            text_box_xml(
                shape_id,
                "Bullet",
                emu(text_x),
                emu(y + 0.08),
                emu(text_w),
                emu(max(0.32, box_h - 0.18)),
                [paragraph_xml(prefix + clamp_text(bullet, text_limit), font_size, theme["text"])],
            )
        )
        shape_id += 1
        y += step
    return shape_id


def render_metric_dashboard(shapes, slide, theme, shape_id):
    cards = [c for c in slide.get("components", []) if c.get("type") == "metric_card"][:3]
    if not cards:
        return render_bullets(shapes, slide, theme, shape_id)

    variant = layout_variant(theme, "metric_dashboard", "strip_cards")
    if variant in {"scoreboard", "alert_cards"}:
        main = cards[0]
        accent = theme_color(theme, main.get("emphasis", "primary"))
        shapes.append(text_box_xml(shape_id, "Hero Metric", emu(0.85), emu(1.95), emu(5.35), emu(2.35), [], fill=theme["surface_alt"], line=accent, radius=True))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Hero Metric Stripe", emu(0.85), emu(1.95), emu(5.35), emu(0.16), accent, accent))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Label", emu(1.2), emu(2.2), emu(4.6), emu(0.32), [paragraph_xml(clamp_text(main.get("label", ""), 28), 1100, theme["muted"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Value", emu(1.2), emu(2.58), emu(4.7), emu(0.8), [paragraph_xml(clamp_text(main.get("value", ""), 22), 3300, accent, True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Context", emu(1.2), emu(3.48), emu(4.6), emu(0.48), [paragraph_xml(clamp_text(main.get("delta", ""), 34), 1500, theme["text"], True), paragraph_xml(clamp_text(main.get("context", ""), 42), 950, theme["muted"])]))
        shape_id += 1
        y = 1.95
        for card in cards[1:]:
            accent = theme_color(theme, card.get("emphasis", "primary"))
            shapes.append(text_box_xml(shape_id, "Side Metric", emu(6.55), emu(y), emu(5.65), emu(1.08), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Side Metric Dot", emu(6.82), emu(y + 0.28), emu(0.18), emu(0.18), accent, accent, "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Side Metric Text", emu(7.12), emu(y + 0.14), emu(4.8), emu(0.72), [paragraph_xml(clamp_text(card.get("label", ""), 26), 1000, theme["muted"], True), paragraph_xml(clamp_text(card.get("value", ""), 24), 1900, accent, True)]))
            shape_id += 1
            y += 1.28
        return render_bullets(shapes, slide, theme, shape_id, start_y=4.85, max_items=1, font_size=1350, text_limit=70)

    if variant in {"vertical_stack", "compact_table"}:
        y = 1.92
        for idx, card in enumerate(cards):
            accent = theme_color(theme, card.get("emphasis", "primary"))
            shapes.append(text_box_xml(shape_id, "Metric Row", emu(0.95), emu(y), emu(11.1), emu(0.82), [], fill=theme["surface"], line=theme["line"], radius=False))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Metric Row Accent", emu(0.95), emu(y), emu(0.1), emu(0.82), accent, accent))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Row Label", emu(1.25), emu(y + 0.17), emu(3.2), emu(0.36), [paragraph_xml(clamp_text(card.get("label", ""), 28), 1250, theme["text"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Row Value", emu(4.65), emu(y + 0.08), emu(3.0), emu(0.42), [paragraph_xml(clamp_text(card.get("value", ""), 24), 1850, accent, True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Row Context", emu(7.65), emu(y + 0.16), emu(4.0), emu(0.38), [paragraph_xml(clamp_text(card.get("delta") or card.get("context", ""), 46), 1100, theme["muted"])]))
            shape_id += 1
            y += 0.98
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.05, max_items=1, font_size=1300, text_limit=70)

    card_w = 3.72
    x_positions = [0.8, 4.81, 8.82]
    y_offsets = [0, 0.18, 0] if variant == "staggered_cards" else [0, 0, 0]
    for idx, card in enumerate(cards):
        accent = theme_color(theme, card.get("emphasis", "primary"))
        x = x_positions[idx]
        y0 = 2.0 + y_offsets[idx]
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Card",
                emu(x),
                emu(y0),
                emu(card_w),
                emu(2.35),
                [],
                fill=theme["surface"],
                line=theme["line"],
                radius=True,
            )
        )
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Metric Accent", emu(x), emu(y0), emu(0.08), emu(2.35), accent, accent))
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Label",
                emu(x + 0.28),
                emu(y0 + 0.18),
                emu(card_w - 0.45),
                emu(0.35),
                [paragraph_xml(clamp_text(card.get("label", ""), 28), 1150, theme["muted"], True)],
            )
        )
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Value",
                emu(x + 0.28),
                emu(y0 + 0.62),
                emu(card_w - 0.45),
                emu(0.66),
                [paragraph_xml(clamp_text(card.get("value", ""), 22), 2800, accent, True)],
            )
        )
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Delta",
                emu(x + 0.28),
                emu(y0 + 1.3),
                emu(card_w - 0.45),
                emu(0.42),
                [paragraph_xml(clamp_text(card.get("delta", ""), 24), 1600, theme["text"], True)],
            )
        )
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Context",
                emu(x + 0.28),
                emu(y0 + 1.78),
                emu(card_w - 0.45),
                emu(0.34),
                [paragraph_xml(clamp_text(card.get("context", ""), 42), 1050, theme["muted"])],
            )
        )
        shape_id += 1

    return render_bullets(shapes, slide, theme, shape_id, start_y=4.85, max_items=1, font_size=1450, text_limit=70)


def render_bar_comparison(shapes, slide, theme, shape_id):
    charts = [c for c in slide.get("components", []) if c.get("type") == "bar_chart"]
    chart = charts[0] if charts else {"data": []}
    data = chart.get("data", [])[:5]
    values = [float(item.get("value", 0)) for item in data if isinstance(item.get("value", 0), (int, float))]
    max_value = max(values) if values else 1
    variant = layout_variant(theme, "bar_comparison", "horizontal_bars")

    shapes.append(
        text_box_xml(
            shape_id,
            "Chart Title",
            emu(0.85),
            emu(1.9),
            emu(5.5),
            emu(0.38),
            [paragraph_xml(clamp_text(chart.get("title", "비교 지표"), 42), 1400, theme["muted"], True)],
        )
    )
    shape_id += 1

    if variant in {"vertical_columns", "split_panel"}:
        accent = theme_color(theme, chart.get("emphasis", "primary"))
        chart_x, chart_y, chart_w, chart_h = 1.1, 2.55, 10.8, 2.0
        col_w = chart_w / max(len(data), 1)
        for idx, item in enumerate(data):
            label = clamp_text(item.get("label", ""), 12)
            value = float(item.get("value", 0))
            unit = safe_text(chart.get("unit", ""))
            h = max(0.18, chart_h * value / max_value)
            x = chart_x + idx * col_w + 0.25
            y = chart_y + chart_h - h
            fill = accent if idx % 2 == 0 else theme["accent"]
            shapes.append(shape_xml(shape_id, "Column Track", emu(x), emu(chart_y), emu(col_w - 0.5), emu(chart_h), theme["surface"], theme["line"], "rect"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Column Value", emu(x), emu(y), emu(col_w - 0.5), emu(h), fill, fill, "rect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Column Number", emu(x - 0.1), emu(y - 0.32), emu(col_w), emu(0.28), [paragraph_xml(f"{value:g}{unit}", 850, fill, True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Column Label", emu(x - 0.12), emu(chart_y + chart_h + 0.12), emu(col_w), emu(0.32), [paragraph_xml(label, 850, theme["muted"], True)]))
            shape_id += 1
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.15, max_items=1, box_h=0.52, step=0.62, font_size=1250, text_limit=64)

    if variant == "lollipop":
        accent = theme_color(theme, chart.get("emphasis", "primary"))
        y = 2.55
        x0, x1 = 2.25, 10.5
        for item in data:
            label = clamp_text(item.get("label", ""), 18)
            value = float(item.get("value", 0))
            unit = safe_text(chart.get("unit", ""))
            x = x0 + (x1 - x0) * value / max_value
            shapes.append(text_box_xml(shape_id, "Lollipop Label", emu(0.9), emu(y - 0.12), emu(1.5), emu(0.3), [paragraph_xml(label, 1100, theme["text"], True)]))
            shape_id += 1
            shapes.append(line_segment_xml(shape_id, "Lollipop Stem", emu(x0), emu(y + 0.08), emu(x), emu(y + 0.08), theme["line"], 12700, "dash"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Lollipop Dot", emu(x - 0.11), emu(y - 0.03), emu(0.22), emu(0.22), accent, "FFFFFF", "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Lollipop Number", emu(x + 0.18), emu(y - 0.12), emu(1.15), emu(0.3), [paragraph_xml(f"{value:g}{unit}", 1050, accent, True)]))
            shape_id += 1
            y += 0.62
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.05, max_items=1, box_h=0.52, step=0.62, font_size=1250, text_limit=64)

    y = 2.5
    bar_x = 2.65
    bar_max_w = 6.6
    accent = theme_color(theme, chart.get("emphasis", "primary"))
    for item in data:
        label = clamp_text(item.get("label", ""), 18)
        value = float(item.get("value", 0))
        unit = safe_text(chart.get("unit", ""))
        bar_w = max(0.15, bar_max_w * value / max_value)
        shapes.append(text_box_xml(shape_id, "Bar Label", emu(0.9), emu(y), emu(1.55), emu(0.36), [paragraph_xml(label, 1250, theme["text"], True)]))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Bar Track", emu(bar_x), emu(y + 0.04), emu(bar_max_w), emu(0.28), "E5E7EB", None, "roundRect"))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Bar Value", emu(bar_x), emu(y + 0.04), emu(bar_w), emu(0.28), accent, accent, "roundRect"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Bar Number", emu(bar_x + bar_max_w + 0.25), emu(y - 0.03), emu(1.0), emu(0.36), [paragraph_xml(f"{value:g}{unit}", 1250, accent, True)]))
        shape_id += 1
        y += 0.68

    return render_bullets(shapes, slide, theme, shape_id, start_y=4.9, max_items=2, box_h=0.52, step=0.62, font_size=1300, text_limit=64)


def render_line_trend(shapes, slide, theme, shape_id):
    charts = [c for c in slide.get("components", []) if c.get("type") == "line_chart"]
    chart = charts[0] if charts else {"data": []}
    data = chart.get("data", [])[:7]
    if len(data) < 2:
        return render_bullets(shapes, slide, theme, shape_id)

    accent = theme_color(theme, chart.get("emphasis", "primary"))
    panel_x, panel_y, panel_w, panel_h = 0.85, 1.88, 11.65, 2.72
    shapes.append(
        text_box_xml(
            shape_id,
            "Line Chart Panel",
            emu(panel_x),
            emu(panel_y),
            emu(panel_w),
            emu(panel_h),
            [],
            fill=theme["surface"],
            line=theme["line"],
            radius=True,
        )
    )
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Line Chart Title",
            emu(panel_x + 0.28),
            emu(panel_y + 0.18),
            emu(5.5),
            emu(0.35),
            [paragraph_xml(clamp_text(chart.get("title", "추세"), 44), 1250, theme["muted"], True)],
        )
    )
    shape_id += 1

    values = [as_number(item.get("value", 0)) for item in data]
    min_v, max_v = min(values), max(values)
    span = max(max_v - min_v, 1)
    chart_x, chart_y, chart_w, chart_h = panel_x + 0.6, panel_y + 0.78, 10.35, 1.38
    for i in range(4):
        y = chart_y + chart_h * i / 3
        shapes.append(line_segment_xml(shape_id, "Grid Line", emu(chart_x), emu(y), emu(chart_x + chart_w), emu(y), theme["line"], 6350))
        shape_id += 1

    points = []
    for idx, item in enumerate(data):
        x = chart_x + chart_w * idx / (len(data) - 1)
        y = chart_y + chart_h * (1 - (as_number(item.get("value", 0)) - min_v) / span)
        points.append((x, y, item))

    for (x1, y1, _), (x2, y2, __) in zip(points, points[1:]):
        shapes.append(line_segment_xml(shape_id, "Trend Segment", emu(x1), emu(y1), emu(x2), emu(y2), accent, 25400))
        shape_id += 1

    unit = safe_text(chart.get("unit", ""))
    for x, y, item in points:
        shapes.append(shape_xml(shape_id, "Trend Marker", emu(x - 0.06), emu(y - 0.06), emu(0.12), emu(0.12), accent, "FFFFFF", "ellipse"))
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Trend Label",
                emu(x - 0.45),
                emu(chart_y + chart_h + 0.12),
                emu(0.9),
                emu(0.28),
                [paragraph_xml(clamp_text(item.get("label", ""), 10), 850, theme["muted"], True)],
            )
        )
        shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Trend Range",
            emu(panel_x + 0.28),
            emu(panel_y + 2.63),
            emu(6.5),
            emu(0.28),
            [paragraph_xml(f"범위: {min_v:g}{unit} - {max_v:g}{unit}", 950, theme["muted"])],
        )
    )
    shape_id += 1
    return render_bullets(
        shapes,
        slide,
        theme,
        shape_id,
        start_y=4.82,
        max_items=2,
        bottom_y=6.18,
        box_h=0.52,
        step=0.62,
        font_size=1250,
        text_limit=62,
    )


def render_timeline(shapes, slide, theme, shape_id):
    timelines = [c for c in slide.get("components", []) if c.get("type") == "timeline"]
    events = (timelines[0].get("events", []) if timelines else [])[:5]
    if not events:
        events = [{"time": f"{idx + 1}", "label": b, "detail": ""} for idx, b in enumerate(slide.get("bullets", [])[:5])]
    if not events:
        return render_bullets(shapes, slide, theme, shape_id)

    x0, x1, y = 1.05, 12.0, 3.0
    shapes.append(line_segment_xml(shape_id, "Timeline Axis", emu(x0), emu(y), emu(x1), emu(y), theme["line"], 25400))
    shape_id += 1
    accent = theme["primary"]
    step = (x1 - x0) / max(len(events) - 1, 1)
    for idx, event in enumerate(events):
        x = x0 + step * idx
        shapes.append(shape_xml(shape_id, "Timeline Dot", emu(x - 0.14), emu(y - 0.14), emu(0.28), emu(0.28), accent, "FFFFFF", "ellipse"))
        shape_id += 1
        box_y = 2.0 if idx % 2 == 0 else 3.38
        shapes.append(line_segment_xml(shape_id, "Timeline Stem", emu(x), emu(y), emu(x), emu(box_y + (0.82 if idx % 2 == 0 else 0)), theme["line"], 9525))
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Timeline Event",
                emu(x - 1.0),
                emu(box_y),
                emu(2.0),
                emu(0.82),
                [
                    paragraph_xml(clamp_text(event.get("time", ""), 18), 900, accent, True),
                    paragraph_xml(clamp_text(event.get("label", ""), 28), 1050, theme["text"], True),
                    paragraph_xml(clamp_text(event.get("detail", ""), 34), 850, theme["muted"]),
                ],
                fill=theme["surface"],
                line=theme["line"],
                radius=True,
            )
        )
        shape_id += 1
    return render_bullets(shapes, slide, theme, shape_id, start_y=5.15, max_items=1, font_size=1250, text_limit=62)


def render_process_flow(shapes, slide, theme, shape_id):
    flows = [c for c in slide.get("components", []) if c.get("type") == "process_flow"]
    steps = (flows[0].get("steps", []) if flows else [])[:4]
    if not steps:
        steps = [{"label": b, "detail": ""} for b in slide.get("bullets", [])[:4]]
    if not steps:
        return render_bullets(shapes, slide, theme, shape_id)

    x, y, w, h, gap = 0.8, 2.15, 2.65, 1.35, 0.42
    for idx, step in enumerate(steps):
        sx = x + idx * (w + gap)
        accent = theme["primary"] if idx % 2 == 0 else theme["accent"]
        shapes.append(
            text_box_xml(
                shape_id,
                "Process Step",
                emu(sx),
                emu(y),
                emu(w),
                emu(h),
                [
                    paragraph_xml(f"{idx + 1:02d}", 950, accent, True),
                    paragraph_xml(clamp_text(step.get("label", ""), 24), 1300, theme["text"], True),
                    paragraph_xml(clamp_text(step.get("detail", ""), 42), 900, theme["muted"]),
                ],
                fill=theme["surface"],
                line=theme["line"],
                radius=True,
            )
        )
        shape_id += 1
        if idx < len(steps) - 1:
            shapes.append(shape_xml(shape_id, "Flow Arrow", emu(sx + w + 0.1), emu(y + 0.48), emu(0.28), emu(0.28), accent, accent, "triangle"))
            shape_id += 1
    return render_bullets(shapes, slide, theme, shape_id, start_y=4.18, max_items=2, box_h=0.52, step=0.62, font_size=1250, text_limit=62)


def render_comparison(shapes, slide, theme, shape_id):
    comps = [c for c in slide.get("components", []) if c.get("type") == "comparison"]
    left = comps[0].get("left", {}) if comps else {"title": "Before", "items": slide.get("bullets", [])[:3]}
    right = comps[0].get("right", {}) if comps else {"title": "After", "items": slide.get("bullets", [])[3:6]}
    columns = [(left, theme["primary"], 0.85), (right, theme["accent"], 6.75)]
    for col, accent, x in columns:
        shapes.append(
            text_box_xml(
                shape_id,
                "Comparison Column",
                emu(x),
                emu(1.95),
                emu(5.55),
                emu(3.2),
                [paragraph_xml(clamp_text(col.get("title", ""), 30), 1500, accent, True)]
                + [paragraph_xml("• " + clamp_text(item, 46), 1250, theme["text"]) for item in col.get("items", [])[:5]],
                fill=theme["surface"],
                line=accent,
                radius=True,
            )
        )
        shape_id += 1
    shapes.append(shape_xml(shape_id, "Comparison Divider", emu(6.48), emu(2.2), emu(0.08), emu(2.7), theme["line"], None, "rect"))
    shape_id += 1
    return shape_id


def render_risk_matrix(shapes, slide, theme, shape_id):
    matrices = [c for c in slide.get("components", []) if c.get("type") == "risk_matrix"]
    matrix = matrices[0] if matrices else {}
    items = matrix.get("items", [])[:8]
    x, y, w, h = 1.15, 2.0, 9.9, 3.5
    mid_x, mid_y = x + w / 2, y + h / 2
    fills = ["ECFDF5", "FEF3C7", "FEE2E2", "F8FAFC"]
    cells = [(x, y + h / 2, fills[3]), (mid_x, y + h / 2, fills[1]), (x, y, fills[1]), (mid_x, y, fills[2])]
    for cx, cy, fill in cells:
        shapes.append(text_box_xml(shape_id, "Risk Cell", emu(cx), emu(cy), emu(w / 2), emu(h / 2), [], fill=fill, line=theme["line"], radius=True))
        shape_id += 1
    shapes.append(line_segment_xml(shape_id, "Risk Axis X", emu(x), emu(y + h), emu(x + w), emu(y + h), theme["muted"], 12700))
    shape_id += 1
    shapes.append(line_segment_xml(shape_id, "Risk Axis Y", emu(x), emu(y + h), emu(x), emu(y), theme["muted"], 12700))
    shape_id += 1
    shapes.append(text_box_xml(shape_id, "Axis X Label", emu(x + w - 1.55), emu(y + h + 0.12), emu(1.7), emu(0.25), [paragraph_xml("영향도 높음", 850, theme["muted"], True)]))
    shape_id += 1
    shapes.append(text_box_xml(shape_id, "Axis Y Label", emu(x - 0.2), emu(y - 0.32), emu(1.8), emu(0.25), [paragraph_xml("가능성 높음", 850, theme["muted"], True)]))
    shape_id += 1
    for item in items:
        likelihood = max(0, min(1, as_number(item.get("likelihood", 0.5), 0.5)))
        impact = max(0, min(1, as_number(item.get("impact", 0.5), 0.5)))
        px = x + 0.3 + (w - 0.6) * impact
        py = y + 0.3 + (h - 0.6) * (1 - likelihood)
        color = theme_color(theme, item.get("emphasis", "warning"))
        shapes.append(shape_xml(shape_id, "Risk Dot", emu(px - 0.08), emu(py - 0.08), emu(0.16), emu(0.16), color, "FFFFFF", "ellipse"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk Label", emu(px + 0.08), emu(py - 0.12), emu(1.35), emu(0.28), [paragraph_xml(clamp_text(item.get("label", ""), 16), 800, theme["text"], True)]))
        shape_id += 1
    return shape_id


def render_cause_effect(shapes, slide, theme, shape_id):
    comps = [c for c in slide.get("components", []) if c.get("type") == "cause_effect"]
    comp = comps[0] if comps else {}
    groups = [
        ("원인", comp.get("causes", slide.get("bullets", [])[:2]), theme["primary"], 0.85),
        ("사건", comp.get("events", slide.get("bullets", [])[2:4]), theme["warning"], 4.55),
        ("결과", comp.get("effects", slide.get("bullets", [])[4:6]), theme["accent"], 8.25),
    ]
    for idx, (title, items, accent, x) in enumerate(groups):
        shapes.append(
            text_box_xml(
                shape_id,
                "Cause Effect Node",
                emu(x),
                emu(2.15),
                emu(3.05),
                emu(2.15),
                [paragraph_xml(title, 1200, accent, True)] + [paragraph_xml("• " + clamp_text(i, 34), 1050, theme["text"]) for i in items[:4]],
                fill=theme["surface"],
                line=accent,
                radius=True,
            )
        )
        shape_id += 1
        if idx < 2:
            shapes.append(shape_xml(shape_id, "Cause Effect Arrow", emu(x + 3.18), emu(3.0), emu(0.36), emu(0.36), accent, accent, "triangle"))
            shape_id += 1
    return render_bullets(shapes, slide, theme, shape_id, start_y=4.9, max_items=1, font_size=1250, text_limit=62)


def render_architecture_map(shapes, slide, theme, shape_id):
    comps = [c for c in slide.get("components", []) if c.get("type") == "architecture_map"]
    comp = comps[0] if comps else {}
    nodes = comp.get("nodes", [])[:7]
    edges = comp.get("edges", [])
    if not nodes:
        nodes = [
            {"id": "user", "label": "User"},
            {"id": "harness", "label": "Harness"},
            {"id": "model", "label": "Model"},
            {"id": "tools", "label": "Tools"},
        ]
        edges = [{"from": "user", "to": "harness"}, {"from": "harness", "to": "model"}, {"from": "harness", "to": "tools"}]
    positions = {
        "user": (1.05, 3.1),
        "harness": (4.25, 2.85),
        "model": (7.6, 1.95),
        "tools": (7.6, 3.35),
        "memory": (10.25, 1.95),
        "skills": (10.25, 3.35),
        "a2a": (10.25, 4.75),
    }
    node_pos = {}
    fallback = [(1.05, 3.1), (4.25, 2.85), (7.6, 1.95), (7.6, 3.35), (10.25, 1.95), (10.25, 3.35), (10.25, 4.75)]
    for idx, node in enumerate(nodes):
        node_id = safe_text(node.get("id", f"n{idx}"))
        node_pos[node_id] = positions.get(node_id, fallback[idx])
    for edge in edges:
        a, b = node_pos.get(edge.get("from")), node_pos.get(edge.get("to"))
        if a and b:
            shapes.append(line_segment_xml(shape_id, "Architecture Edge", emu(a[0] + 0.7), emu(a[1] + 0.25), emu(b[0]), emu(b[1] + 0.25), theme["line"], 12700))
            shape_id += 1
    for idx, node in enumerate(nodes):
        node_id = safe_text(node.get("id", f"n{idx}"))
        x, y = node_pos[node_id]
        accent = theme_color(theme, node.get("emphasis", "primary" if idx == 1 else "accent"))
        shapes.append(text_box_xml(shape_id, "Architecture Node", emu(x), emu(y), emu(1.55), emu(0.72), [paragraph_xml(clamp_text(node.get("label", node_id), 18), 1100, theme["text"], True)], fill=theme["surface"], line=accent, radius=True))
        shape_id += 1
    return shape_id


def render_callout_focus(shapes, slide, theme, shape_id):
    callouts = [c for c in slide.get("components", []) if c.get("type") == "callout"]
    callout = callouts[0] if callouts else {}
    headline = callout.get("headline") or slide.get("title", "")
    body = callout.get("body") or " ".join(slide.get("bullets", [])[:2])
    shapes.append(translucent_shape_xml(shape_id, "Callout Accent", emu(0.9), emu(1.95), emu(0.22), emu(3.1), theme["primary"], 65000))
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Callout Panel",
            emu(1.25),
            emu(1.95),
            emu(10.9),
            emu(3.1),
            [
                paragraph_xml(clamp_text(headline, 58), 2500, theme["primary"], True),
                paragraph_xml(clamp_text(body, 120), 1450, theme["text"]),
            ],
            fill=theme["surface_alt"],
            line=theme["line"],
            radius=True,
        )
    )
    return shape_id + 1


def render_title_cover(shapes, slide, theme, shape_id):
    kicker = slide.get("kicker") or "Presentation"
    title = slide.get("title") or "Untitled Deck"
    subtitle = slide.get("subtitle") or ""
    variant = layout_variant(theme, "title_cover", "left_panel")
    if variant in {"centered_badge", "weather_alert"}:
        shapes.append(translucent_shape_xml(shape_id, "Cover Badge", emu(2.0), emu(1.15), emu(9.35), emu(4.75), theme["surface_alt"], 64000, "roundRect"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Cover Halo", emu(9.15), emu(0.65), emu(2.1), emu(2.1), theme["accent"], 16000, "ellipse"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Kicker", emu(3.0), emu(1.78), emu(7.2), emu(0.36), [paragraph_xml(clamp_text(kicker, 44), 1200, theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Title", emu(2.65), emu(2.28), emu(8.05), emu(1.55), [paragraph_xml(clamp_text(title, 50), 3400, theme["text"], True)]))
        shape_id += 1
        if subtitle:
            shapes.append(text_box_xml(shape_id, "Cover Subtitle", emu(3.0), emu(4.18), emu(7.3), emu(0.62), [paragraph_xml(clamp_text(subtitle, 76), 1450, theme["muted"])]))
            shape_id += 1
        return shape_id

    if variant in {"diagonal_hero", "executive_memo"}:
        shapes.append(translucent_shape_xml(shape_id, "Cover Diagonal", emu(7.7), emu(0.7), emu(5.1), emu(5.5), theme["primary"], 17000, "parallelogram"))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Cover Top Rule", emu(0.85), emu(1.25), emu(4.1), emu(0.08), theme["accent"], theme["accent"]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Kicker", emu(0.9), emu(1.58), emu(5.8), emu(0.36), [paragraph_xml(clamp_text(kicker, 44), 1200, theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Title", emu(0.85), emu(2.02), emu(7.6), emu(1.85), [paragraph_xml(clamp_text(title, 54), 3600, theme["text"], True)]))
        shape_id += 1
        if subtitle:
            shapes.append(text_box_xml(shape_id, "Cover Subtitle", emu(0.9), emu(4.12), emu(6.8), emu(0.7), [paragraph_xml(clamp_text(subtitle, 76), 1450, theme["muted"])]))
            shape_id += 1
        return shape_id

    shapes.append(translucent_shape_xml(shape_id, "Cover Wash", emu(0.72), emu(1.15), emu(11.85), emu(4.85), theme["surface_alt"], 62000, "roundRect"))
    shape_id += 1
    shapes.append(shape_xml(shape_id, "Cover Accent Bar", emu(0.72), emu(1.15), emu(0.16), emu(4.85), theme["primary"], theme["primary"]))
    shape_id += 1
    shapes.append(translucent_shape_xml(shape_id, "Cover Accent Block", emu(9.65), emu(4.78), emu(2.15), emu(0.28), theme["accent"], 22000, "parallelogram"))
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Cover Kicker",
            emu(1.15),
            emu(1.58),
            emu(5.8),
            emu(0.36),
            [paragraph_xml(clamp_text(kicker, 44), 1200, theme["primary"], True)],
        )
    )
    shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Cover Title",
            emu(1.1),
            emu(2.05),
            emu(10.8),
            emu(1.8),
            [paragraph_xml(clamp_text(title, 58), 3600, theme["text"], True)],
        )
    )
    shape_id += 1
    if subtitle:
        shapes.append(
            text_box_xml(
                shape_id,
                "Cover Subtitle",
                emu(1.15),
                emu(4.05),
                emu(9.8),
                emu(0.72),
                [paragraph_xml(clamp_text(subtitle, 86), 1550, theme["muted"])],
            )
        )
        shape_id += 1
    shapes.append(line_segment_xml(shape_id, "Cover Rule", emu(1.15), emu(5.34), emu(4.5), emu(5.34), theme["primary"], 19050))
    shape_id += 1
    return shape_id


def render_takeaway(shapes, slide, theme, shape_id):
    variant = layout_variant(theme, "takeaway", "summary_panel")
    bullets = slide.get("bullets", [])[:4]
    if variant in {"quote_band", "forecast_brief"}:
        shapes.append(translucent_shape_xml(shape_id, "Takeaway Band", emu(0), emu(2.05), SLIDE_W, emu(2.45), theme["surface_alt"], 60000))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Takeaway Rule", emu(0.85), emu(2.28), emu(0.1), emu(1.95), theme["primary"], theme["primary"]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Head", emu(1.2), emu(2.25), emu(4.0), emu(0.42), [paragraph_xml("핵심 메시지", 1300, theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Quote", emu(1.2), emu(2.85), emu(10.4), emu(1.25), [paragraph_xml(clamp_text(bullets[0] if bullets else slide.get("title", ""), 78), 2300, theme["text"], True)]))
        shape_id += 1
        if len(bullets) > 1:
            return render_bullets(shapes, {"bullets": bullets[1:]}, theme, shape_id, start_y=4.85, max_items=2, box_h=0.48, step=0.56, font_size=1200, text_limit=70)
        return shape_id

    if variant in {"full_bleed_callout", "decision_memo"}:
        shapes.append(text_box_xml(shape_id, "Decision Panel", emu(0.85), emu(1.9), emu(5.5), emu(3.55), [paragraph_xml("결론", 1300, theme["primary"], True), paragraph_xml(clamp_text(slide.get("title", "Takeaway"), 44), 2500, theme["text"], True)], fill=theme["surface_alt"], line=theme["line"], radius=True))
        shape_id += 1
        y = 2.05
        for idx, bullet in enumerate(bullets):
            shapes.append(text_box_xml(shape_id, "Decision Item", emu(6.65), emu(y), emu(5.3), emu(0.62), [paragraph_xml(f"{idx + 1}. {clamp_text(bullet, 58)}", 1350, theme["text"], True)], fill=theme["surface"], line=theme["line"], radius=False))
            shape_id += 1
            y += 0.78
        return shape_id

    shapes.append(
        text_box_xml(
            shape_id,
            "Takeaway Panel",
            emu(0.85),
            emu(1.95),
            emu(11.65),
            emu(3.05),
            [paragraph_xml("핵심 메시지", 1300, theme["primary"], True)]
            + [paragraph_xml("• " + clamp_text(b, 82), 1900, theme["text"]) for b in bullets],
            fill=theme["surface_alt"],
            line=theme["line"],
            radius=True,
        )
    )
    return shape_id + 1


def render_slide_body(shapes, slide, theme, shape_id):
    layout = slide.get("layout", "bullets")
    if layout == "title_cover":
        return render_title_cover(shapes, slide, theme, shape_id)
    if layout == "metric_dashboard":
        return render_metric_dashboard(shapes, slide, theme, shape_id)
    if layout == "bar_comparison":
        return render_bar_comparison(shapes, slide, theme, shape_id)
    if layout == "line_trend":
        return render_line_trend(shapes, slide, theme, shape_id)
    if layout == "timeline":
        return render_timeline(shapes, slide, theme, shape_id)
    if layout == "process_flow":
        return render_process_flow(shapes, slide, theme, shape_id)
    if layout == "comparison":
        return render_comparison(shapes, slide, theme, shape_id)
    if layout == "risk_matrix":
        return render_risk_matrix(shapes, slide, theme, shape_id)
    if layout == "cause_effect":
        return render_cause_effect(shapes, slide, theme, shape_id)
    if layout == "architecture_map":
        return render_architecture_map(shapes, slide, theme, shape_id)
    if layout == "callout_focus":
        return render_callout_focus(shapes, slide, theme, shape_id)
    if layout == "takeaway":
        return render_takeaway(shapes, slide, theme, shape_id)
    return render_bullets(shapes, slide, theme, shape_id)


def make_slide_xml(slide, index, theme):
    shapes = []
    shape_id = add_background_motif(shapes, theme, 2)
    if slide.get("layout") != "title_cover":
        shape_id = add_header(shapes, slide, theme, shape_id)
    shape_id = render_slide_body(shapes, slide, theme, shape_id)
    if slide.get("layout") != "title_cover":
        add_note_and_page(shapes, slide, theme, index, shape_id)

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Slide {index}">
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="{theme["background"]}"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {''.join(shapes)}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def content_types_xml(slide_count):
    overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {overrides}
</Types>"""


def root_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def presentation_xml(slide_count):
    slide_ids = "\n".join(f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, slide_count + 1))
    notes_master_id = slide_count + 2
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                saveSubsetFonts="1" autoCompressPictures="0">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>{slide_ids}</p:sldIdLst>
  <p:notesMasterIdLst><p:notesMasterId r:id="rId{notes_master_id}"/></p:notesMasterIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}"/>
  <p:notesSz cx="6858000" cy="{SLIDE_W}"/>
  <p:defaultTextStyle>
    <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>
    <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>
    <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>
    <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>
    <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>
    <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>
    <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>
    <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>
    <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>
  </p:defaultTextStyle>
</p:presentation>"""


def presentation_rels_xml(slide_count):
    rels = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    ]
    rels.extend(
        f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    )
    base = slide_count + 2
    rels.extend([
        f'<Relationship Id="rId{base}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>',
        f'<Relationship Id="rId{base + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>',
        f'<Relationship Id="rId{base + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>',
        f'<Relationship Id="rId{base + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
        f'<Relationship Id="rId{base + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>',
    ])
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {' '.join(rels)}
</Relationships>"""


def slide_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def slide_layout_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


def slide_layout_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""


def slide_master_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"""


def slide_master_rels_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""


def theme_xml(theme):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="DeckSpec Theme">
  <a:themeElements>
    <a:clrScheme name="DeckSpec">
      <a:dk1><a:srgbClr val="{theme["text"]}"/></a:dk1><a:lt1><a:srgbClr val="{theme["background"]}"/></a:lt1>
      <a:dk2><a:srgbClr val="{theme["muted"]}"/></a:dk2><a:lt2><a:srgbClr val="{theme["surface"]}"/></a:lt2>
      <a:accent1><a:srgbClr val="{theme["primary"]}"/></a:accent1><a:accent2><a:srgbClr val="{theme["accent"]}"/></a:accent2>
      <a:accent3><a:srgbClr val="{theme["warning"]}"/></a:accent3><a:accent4><a:srgbClr val="DC2626"/></a:accent4>
      <a:accent5><a:srgbClr val="7C3AED"/></a:accent5><a:accent6><a:srgbClr val="2563EB"/></a:accent6>
      <a:hlink><a:srgbClr val="{theme["primary"]}"/></a:hlink><a:folHlink><a:srgbClr val="{theme["accent"]}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Aptos">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Aptos"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Malgun Gothic"/><a:cs typeface="Aptos"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Simple">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="90000"/><a:satMod val="105000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="105000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="75000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="12000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="98000"/><a:satMod val="130000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>"""


def core_xml(deckspec):
    title = xml_escape(deckspec.get("deck_title", "Deck generated from DeckSpec OOXML"))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{title}</dc:title>
  <dc:creator>make_ppt.py</dc:creator>
</cp:coreProperties>"""


def app_xml(slide_count):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Python DeckSpec OOXML renderer</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>{slide_count}</Slides>
</Properties>"""


def normalize_deckspec(deckspec):
    if not isinstance(deckspec, dict):
        raise ValueError("DeckSpec must be a JSON object.")
    slides = deckspec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("DeckSpec must include a non-empty slides array.")
    deckspec.setdefault("design", {"theme": "policy_brief"})
    for idx, slide in enumerate(slides, 1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {idx} must be an object.")
        slide.setdefault("layout", "bullets")
        slide.setdefault("kicker", deckspec.get("deck_title", ""))
        slide.setdefault("title", f"Slide {idx}")
        slide.setdefault("subtitle", "")
        slide.setdefault("bullets", [])
        slide.setdefault("components", [])
        slide.setdefault("speaker_note", "")
    return deckspec


def build_ooxml_files(deckspec):
    deckspec = normalize_deckspec(deckspec)
    slides = deckspec["slides"]
    theme = theme_for(deckspec)
    files = {
        "[Content_Types].xml": content_types_xml(len(slides)),
        "_rels/.rels": root_rels_xml(),
        "docProps/core.xml": core_xml(deckspec),
        "docProps/app.xml": app_xml(len(slides)),
        "ppt/presentation.xml": presentation_xml(len(slides)),
        "ppt/_rels/presentation.xml.rels": presentation_rels_xml(len(slides)),
    }
    if BASE_OOXML_DIR.exists():
        for path in BASE_OOXML_DIR.rglob("*"):
            if path.is_file() and not any(part.startswith(".") for part in path.relative_to(BASE_OOXML_DIR).parts):
                files[str(path.relative_to(BASE_OOXML_DIR))] = path.read_text(encoding="utf-8")
    else:
        files.update(
            {
                "ppt/slideMasters/slideMaster1.xml": slide_master_xml(),
                "ppt/slideMasters/_rels/slideMaster1.xml.rels": slide_master_rels_xml(),
                "ppt/slideLayouts/slideLayout1.xml": slide_layout_xml(),
                "ppt/slideLayouts/_rels/slideLayout1.xml.rels": slide_layout_rels_xml(),
                "ppt/theme/theme1.xml": theme_xml(theme),
            }
        )
    for idx, slide in enumerate(slides, 1):
        files[f"ppt/slides/slide{idx}.xml"] = make_slide_xml(slide, idx, theme)
        files[f"ppt/slides/_rels/slide{idx}.xml.rels"] = slide_rels_xml()
    return files


def write_pptx(files, output_path):
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as deck:
        for name, data in files.items():
            deck.writestr(name, data)


def dump_ooxml(files, dump_dir):
    dump_path = Path(dump_dir)
    for name, data in files.items():
        file_path = dump_path / name
        file_path.parent.mkdir(parents=True, exist_ok=True)
        file_path.write_text(data, encoding="utf-8")


def load_deckspec(path=None, json_string=None):
    if json_string:
        return json.loads(json_string)
    if path == "-":
        return json.loads(sys.stdin.read())
    if path:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    return TEST_DECKSPEC


def write_template(path):
    Path(path).write_text(json.dumps(TEST_DECKSPEC, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(
        description="Render a DeckSpec JSON with design/layout/components into OOXML files and package them as PPTX."
    )
    parser.add_argument("deckspec", nargs="?", help="DeckSpec JSON file. If omitted, built-in sample DeckSpec is used.")
    parser.add_argument("--json-string", help="DeckSpec JSON string. Useful when passing JSON directly instead of a file.")
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT, help=f"Output PPTX path. Default: {DEFAULT_OUTPUT}")
    parser.add_argument("--dump-ooxml", help="Optional directory to write generated OOXML files before packaging.")
    parser.add_argument("--write-template", help="Write a DeckSpec JSON template to this path and exit.")
    args = parser.parse_args()

    if args.write_template:
        write_template(args.write_template)
        print(f"Wrote template: {os.path.abspath(args.write_template)}")
        return

    deckspec = load_deckspec(args.deckspec, args.json_string)
    files = build_ooxml_files(deckspec)

    if args.dump_ooxml:
        dump_ooxml(files, args.dump_ooxml)

    output_path = os.path.abspath(args.output)
    write_pptx(files, output_path)

    print(f"Created {output_path}")
    print(f"Slides: {len(deckspec['slides'])}")
    selected_template = deckspec.get("design", {}).get("selected_template")
    if selected_template:
        print(f"Template: {selected_template}")
    if args.dump_ooxml:
        print(f"OOXML dumped to {os.path.abspath(args.dump_ooxml)}")


if __name__ == "__main__":
    main()
