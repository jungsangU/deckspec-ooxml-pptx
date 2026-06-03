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
MIN_READABLE_FONT = 950


def installed_typeface(candidates):
    font_dirs = [
        Path.home() / "Library" / "Fonts",
        Path("/Library/Fonts"),
        Path("/System/Library/Fonts"),
        Path("C:/Windows/Fonts"),
    ]
    installed = []
    for font_dir in font_dirs:
        if not font_dir.exists():
            continue
        try:
            installed.extend(path.name.lower() for path in font_dir.iterdir())
        except OSError:
            continue
    for candidate in candidates:
        key = candidate.lower().replace(" ", "")
        if any(key in name.replace(" ", "") for name in installed):
            return candidate
    return candidates[-1]


TYPEFACE_EAST_ASIAN = installed_typeface(["Pretendard", "SUIT", "Noto Sans CJK KR", "Noto Sans KR", "Apple SD Gothic Neo", "Malgun Gothic"])
TYPEFACE_LATIN = installed_typeface(["Pretendard", "SUIT", "Aptos", "Arial"])
TYPEFACE_COMPLEX = TYPEFACE_LATIN

FONT = {
    "micro": 950,
    "caption": 1050,
    "label": 1200,
    "body_small": 1300,
    "body": 1450,
    "body_large": 1650,
    "section": 1900,
    "display": 2800,
    "hero": 3400,
}

CONTENT = {
    "left": 0.85,
    "top": 1.9,
    "right": 12.45,
    "bottom": 6.15,
    "width": 11.6,
    "gap": 0.28,
    "row_h": 0.76,
    "row_step": 0.94,
}

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
        style_keys = {
            "primary",
            "accent",
            "warning",
            "surface_alt",
            "line",
            "motif",
            "bullet_style",
            "layout_variants",
            "template_style",
            "frame_color",
            "dot_color",
            "title_fill",
            "progress_fill",
        }
        theme.update({k: v for k, v in profile.items() if k in style_keys})
        design["selected_template"] = profile["name"]
        deckspec["design"] = design
    return theme


def theme_color(theme, key):
    return theme.get(key, theme["primary"])


def candidate_variants(entry, default):
    if isinstance(entry, str):
        return entry, [entry], {}
    if isinstance(entry, list):
        candidates = [safe_text(v) for v in entry if safe_text(v)]
        return candidates[0] if candidates else default, candidates or [default], {}
    if isinstance(entry, dict):
        candidates = [safe_text(v) for v in entry.get("candidates", []) if safe_text(v)]
        default_variant = safe_text(entry.get("default", candidates[0] if candidates else default))
        if default_variant and default_variant not in candidates:
            candidates.insert(0, default_variant)
        return default_variant or default, candidates or [default], entry.get("rules", {})
    return default, [default], {}


def slide_variant_hint(slide, candidates):
    requested = safe_text(slide.get("variant") or slide.get("preferred_variant"))
    if requested and requested not in {"auto", "content_aware", "content-aware"}:
        return requested if requested in candidates else None
    return None


def content_variant_scores(layout, slide, candidates, rules):
    scores = {candidate: 0 for candidate in candidates}
    bullets = slide.get("bullets", [])
    components = slide.get("components", [])
    if layout == "metric_dashboard":
        metric_count = len([c for c in components if c.get("type") == "metric_card"])
        if metric_count >= 3:
            for candidate in ["radial_bubbles", "donut_metrics", "staggered_cards"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        if metric_count <= 2:
            for candidate in ["scoreboard", "alert_cards", "strip_cards"]:
                scores[candidate] = scores.get(candidate, 0) + 3
        if len(bullets) >= 3:
            for candidate in ["vertical_stack", "compact_table"]:
                scores[candidate] = scores.get(candidate, 0) + 2
    elif layout == "bar_comparison":
        chart = next((c for c in components if c.get("type") == "bar_chart"), {})
        data_len = len(chart.get("data", []))
        if data_len <= 3:
            for candidate in ["progress_rows", "blue_progress_rows", "lollipop"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        if data_len >= 4:
            for candidate in ["horizontal_bars", "vertical_columns", "split_panel"]:
                scores[candidate] = scores.get(candidate, 0) + 3
    elif layout == "title_cover":
        title_len = len(safe_text(slide.get("title", "")))
        if title_len > 28:
            for candidate in ["left_panel", "diagonal_hero", "executive_memo"]:
                scores[candidate] = scores.get(candidate, 0) + 2
        else:
            for candidate in ["blue_ribbon_title", "center_capsule_title", "centered_badge"]:
                scores[candidate] = scores.get(candidate, 0) + 2
    elif layout == "takeaway":
        if len(bullets) >= 3:
            for candidate in ["ribbon_summary", "decision_memo", "summary_panel"]:
                scores[candidate] = scores.get(candidate, 0) + 3
        else:
            for candidate in ["quote_band", "forecast_brief", "full_bleed_callout"]:
                scores[candidate] = scores.get(candidate, 0) + 2
    elif layout == "line_trend":
        chart_count = len([c for c in components if c.get("type") == "line_chart"])
        point_count = max([len(c.get("data", [])) for c in components if c.get("type") == "line_chart"] or [0])
        if chart_count >= 2:
            for candidate in ["sparkline_stack", "small_multiples"]:
                scores[candidate] = scores.get(candidate, 0) + 5
        elif point_count >= 5:
            for candidate in ["wide_plot_callout", "horizon_plot"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        else:
            for candidate in ["milestone_line", "compact_trend"]:
                scores[candidate] = scores.get(candidate, 0) + 3
    elif layout == "cause_effect":
        comp = next((c for c in components if c.get("type") == "cause_effect"), {})
        groups = comp.get("causes", []) + comp.get("events", []) + comp.get("effects", [])
        dense = len(groups) >= 7 or item_weight(groups) > 125
        if dense:
            for candidate in ["vertical_story", "split_swimlane"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        else:
            for candidate in ["cascade_cards", "center_bridge"]:
                scores[candidate] = scores.get(candidate, 0) + 4
    elif layout == "risk_matrix":
        matrix = next((c for c in components if c.get("type") == "risk_matrix"), {})
        item_count = len(matrix.get("items", []))
        if item_count >= 6:
            for candidate in ["ranked_watchlist", "compact_quadrant"]:
                scores[candidate] = scores.get(candidate, 0) + 5
        else:
            for candidate in ["quadrant_watchlist", "heatmap_focus"]:
                scores[candidate] = scores.get(candidate, 0) + 4
    elif layout == "comparison":
        comp = next((c for c in components if c.get("type") == "comparison"), {})
        left_count = len(comp.get("left", {}).get("items", []))
        right_count = len(comp.get("right", {}).get("items", []))
        if max(left_count, right_count) >= 4:
            for candidate in ["stacked_scorecards", "dense_columns"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        else:
            for candidate in ["split_columns", "before_after_cards"]:
                scores[candidate] = scores.get(candidate, 0) + 4
    elif layout in {"bullets", "title_summary"}:
        density = content_density(bullets)
        if density == "low":
            for candidate in ["tile_grid", "spotlight_list"]:
                scores[candidate] = scores.get(candidate, 0) + 4
        elif density == "high":
            for candidate in ["compact_rows", "sectioned_rows"]:
                scores[candidate] = scores.get(candidate, 0) + 4

    for rule_key, variant in rules.items():
        variant = safe_text(variant)
        if not variant:
            continue
        if rule_key == "many_metrics" and len([c for c in components if c.get("type") == "metric_card"]) >= 3:
            scores[variant] = scores.get(variant, 0) + 6
        if rule_key == "few_big_numbers" and len([c for c in components if c.get("type") == "metric_card"]) <= 2:
            scores[variant] = scores.get(variant, 0) + 6
        if rule_key == "many_bars":
            chart = next((c for c in components if c.get("type") == "bar_chart"), {})
            if len(chart.get("data", [])) >= 4:
                scores[variant] = scores.get(variant, 0) + 6
    return scores


def layout_variant(theme, layout, default="default", slide=None):
    variants = theme.get("layout_variants", {})
    entry = variants.get(layout, default) if isinstance(variants, dict) else default
    default_variant, candidates, rules = candidate_variants(entry, default)
    slide = slide or {}
    hinted = slide_variant_hint(slide, candidates)
    if hinted:
        return hinted
    usage = theme.setdefault("_variant_usage", {})
    scores = content_variant_scores(layout, slide, candidates, rules)
    ranked = sorted(
        candidates,
        key=lambda candidate: (
            -(scores.get(candidate, 0) - usage.get((layout, candidate), 0) * 4),
            usage.get((layout, candidate), 0),
            0 if candidate == default_variant else 1,
            candidates.index(candidate),
        ),
    )
    selected = ranked[0] if ranked else default_variant
    usage[(layout, selected)] = usage.get((layout, selected), 0) + 1
    slide["_selected_variant"] = selected
    return selected


def srgb_fill_xml(color, alpha=None):
    if alpha is None:
        return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    return f'<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr></a:solidFill>'


def paragraph_xml(text, size=2200, color="374151", bold=False, align=None):
    size = max(int(size), MIN_READABLE_FONT)
    bold_attr = ' b="1"' if bold else ""
    align_map = {"center": "ctr", "ctr": "ctr", "right": "r", "r": "r"}
    align_attr = f' algn="{align_map.get(align, align)}"' if align else ""
    return (
        "<a:p>"
        f'<a:pPr marL="0" indent="0"{align_attr}><a:buNone/></a:pPr>'
        "<a:r>"
        f'<a:rPr lang="ko-KR" sz="{size}"{bold_attr}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{TYPEFACE_LATIN}"/><a:ea typeface="{TYPEFACE_EAST_ASIAN}"/><a:cs typeface="{TYPEFACE_COMPLEX}"/>'
        "</a:rPr>"
        f"<a:t>{xml_escape(text)}</a:t>"
        "</a:r>"
        f'<a:endParaRPr lang="ko-KR" sz="{size}"/>'
        "</a:p>"
    )


def text_box_xml(shape_id, name, x, y, w, h, paragraphs, fill=None, line=None, radius=False, anchor="t", margin_x=91440, margin_y=45720):
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
          <a:bodyPr wrap="square" lIns="{margin_x}" tIns="{margin_y}" rIns="{margin_x}" bIns="{margin_y}" anchor="{anchor}"/>
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


def outlined_shape_xml(shape_id, name, x, y, w, h, line, width=25400, shape="roundRect", fill=None):
    fill_xml = "<a:noFill/>" if fill is None else srgb_fill_xml(fill)
    geom = (
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom>'
        if shape == "roundRect"
        else f'<a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>'
    )
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          {geom}
          {fill_xml}
          <a:ln w="{width}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>
        </p:spPr>
      </p:sp>"""


def pattern_shape_xml(shape_id, name, x, y, w, h, bg, fg, pattern="pct5", shape="rect"):
    return f"""
      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{xml_escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>
          <a:prstGeom prst="{shape}"><a:avLst/></a:prstGeom>
          <a:pattFill prst="{pattern}">
            <a:fgClr><a:srgbClr val="{fg}"/></a:fgClr>
            <a:bgClr><a:srgbClr val="{bg}"/></a:bgClr>
          </a:pattFill>
          <a:ln><a:noFill/></a:ln>
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


def text_weight(value):
    text = safe_text(value)
    korean = len(re.findall(r"[가-힣]", text))
    latin = len(re.findall(r"[A-Za-z0-9]", text))
    other = max(0, len(text) - korean - latin)
    return korean * 1.05 + latin * 0.58 + other * 0.5


def item_weight(items):
    return sum(text_weight(item) for item in items if safe_text(item))


def clamp_num(value, low, high):
    return max(low, min(high, value))


def fit_font(text, width_in, base=FONT["body"], minimum=FONT["caption"], maximum=FONT["section"]):
    weight = max(text_weight(text), 1)
    capacity = max(width_in * 11.5, 1)
    if weight <= capacity:
        return int(clamp_num(base, minimum, maximum))
    scale = capacity / weight
    return int(clamp_num(base * max(0.72, scale), minimum, maximum))


def list_font(items, width_in, base=FONT["body_small"], minimum=FONT["caption"]):
    max_weight = max([text_weight(item) for item in items if safe_text(item)] or [1])
    capacity = max(width_in * 11.0, 1)
    if max_weight <= capacity:
        return int(base)
    return int(clamp_num(base * capacity / max_weight, minimum, base))


def compact_text_limit(width_in, font_size):
    return max(14, int(width_in * (13000 / max(font_size, 1))))


def content_density(items):
    count = len([item for item in items if safe_text(item)])
    weight = item_weight(items)
    if count >= 6 or weight > 190:
        return "high"
    if count >= 4 or weight > 105:
        return "medium"
    return "low"


def grid_slots(count, x, y, w, h, gap=0.24, prefer_columns=None):
    count = max(1, count)
    if prefer_columns:
        cols = min(prefer_columns, count)
    elif count <= 2:
        cols = count
    elif count <= 4:
        cols = 2
    else:
        cols = 3
    rows = (count + cols - 1) // cols
    cell_w = (w - gap * (cols - 1)) / cols
    cell_h = (h - gap * (rows - 1)) / rows
    slots = []
    for idx in range(count):
        col = idx % cols
        row = idx // cols
        slots.append((x + col * (cell_w + gap), y + row * (cell_h + gap), cell_w, cell_h))
    return slots


def truncate_items(items, limit):
    return [clamp_text(item, limit) for item in items if safe_text(item)]


def add_role_background(shapes, slide, theme, shape_id):
    layout = slide.get("layout", "bullets")
    accent = theme_color(theme, slide.get("emphasis", "accent"))
    if layout == "metric_dashboard":
        shapes.append(translucent_shape_xml(shape_id, "Metric Canvas Wash", emu(0), emu(0), SLIDE_W, SLIDE_H, theme["surface_alt"], 85000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Metric Giant Disc", emu(7.65), emu(0.75), emu(4.35), emu(4.35), theme["primary"], 12000, "ellipse"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Metric Offset Disc", emu(9.9), emu(3.9), emu(1.7), emu(1.7), theme["accent"], 22000, "ellipse"))
        shape_id += 1
    elif layout == "bar_comparison":
        shapes.append(translucent_shape_xml(shape_id, "Bar Canvas Panel", emu(0.55), emu(1.72), emu(11.95), emu(4.6), theme["surface"], 92000, "roundRect"))
        shape_id += 1
        for x in [2.8, 4.6, 6.4, 8.2, 10.0]:
            shapes.append(line_segment_xml(shape_id, "Bar Grid Guide", emu(x), emu(2.1), emu(x), emu(5.85), theme["line"], 3175, "dash"))
            shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Bar Side Heat", emu(0), emu(0), emu(0.32), SLIDE_H, accent, 18000))
        shape_id += 1
    elif layout == "cause_effect":
        shapes.append(translucent_shape_xml(shape_id, "Cascade Left Field", emu(0), emu(0), emu(3.65), SLIDE_H, theme["surface_alt"], 76000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Cascade Path Glow", emu(3.65), emu(2.8), emu(5.1), emu(0.72), theme["primary"], 15000, "parallelogram"))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Cascade Motion Rail", emu(1.08), emu(5.95), emu(12.0), emu(1.98), theme["line"], 6350, "dash"))
        shape_id += 1
    elif layout == "risk_matrix":
        shapes.append(translucent_shape_xml(shape_id, "Risk Alert Field", emu(8.9), emu(0), emu(4.45), SLIDE_H, theme["surface_alt"], 70000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Risk High Zone", emu(8.7), emu(1.25), emu(2.7), emu(2.4), theme["warning"], 16000, "roundRect"))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Risk Diagonal Watchline", emu(0.75), emu(6.1), emu(12.1), emu(1.25), theme["warning"], 6350, "dash"))
        shape_id += 1
    elif layout == "comparison":
        shapes.append(translucent_shape_xml(shape_id, "Comparison Left Wash", emu(0), emu(0), emu(6.65), SLIDE_H, theme["surface_alt"], 76000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Comparison Right Wash", emu(6.65), emu(0), emu(6.7), SLIDE_H, theme["surface"], 92000))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Comparison Center Rule", emu(6.65), emu(0.65), emu(6.65), emu(6.2), theme["primary"], 12700, "dash"))
        shape_id += 1
    elif layout == "line_trend":
        shapes.append(translucent_shape_xml(shape_id, "Trend Horizon Wash", emu(0), emu(4.35), SLIDE_W, emu(2.45), theme["surface_alt"], 78000))
        shape_id += 1
        for y in [2.15, 3.05, 3.95, 4.85]:
            shapes.append(line_segment_xml(shape_id, "Trend Horizontal Guide", emu(0.9), emu(y), emu(12.05), emu(y), theme["line"], 3175, "dash"))
            shape_id += 1
    elif layout == "takeaway":
        shapes.append(translucent_shape_xml(shape_id, "Takeaway Full Bleed Band", emu(0), emu(4.78), SLIDE_W, emu(2.25), theme["primary"], 18000))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Takeaway Accent Block", emu(9.6), emu(0.6), emu(2.5), emu(2.5), theme["accent"], 16000, "ellipse"))
        shape_id += 1
    elif layout in {"title_summary", "bullets", "callout_focus"}:
        shapes.append(translucent_shape_xml(shape_id, "Summary Quiet Field", emu(0.7), emu(1.75), emu(11.85), emu(4.75), theme["surface"], 94000, "roundRect"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Summary Corner Marker", emu(10.85), emu(0.62), emu(1.4), emu(1.4), theme["surface_alt"], 36000, "ellipse"))
        shape_id += 1
    elif layout in {"timeline", "process_flow", "architecture_map"}:
        shapes.append(translucent_shape_xml(shape_id, "System Blueprint Field", emu(0), emu(0), SLIDE_W, SLIDE_H, theme["surface_alt"], 90000))
        shape_id += 1
        for x in [1.6, 3.7, 5.8, 7.9, 10.0, 12.1]:
            shapes.append(line_segment_xml(shape_id, "System Blueprint Guide", emu(x), emu(1.55), emu(x), emu(6.05), theme["line"], 3175, "dash"))
            shape_id += 1
    return shape_id


def add_background_motif(shapes, theme, shape_id, slide=None):
    if slide and slide.get("layout") != "title_cover":
        return add_role_background(shapes, slide, theme, shape_id)

    if theme.get("template_style") == "blue_ribbon_business":
        navy = theme.get("frame_color", theme["accent"])
        shapes.append(translucent_shape_xml(shape_id, "Blue Ribbon Top Slab", emu(9.35), emu(0.22), emu(3.2), emu(0.82), navy, 17000, "roundRect"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Blue Accent Tab", emu(10.35), emu(1.02), emu(1.35), emu(0.34), theme["primary"], 19000, "roundRect"))
        shape_id += 1
        shapes.append(translucent_shape_xml(shape_id, "Blue Bottom Wash", emu(0), emu(6.75), SLIDE_W, emu(0.28), theme["surface_alt"], 35000))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Blue Ghost Circle", emu(0.55), emu(5.55), emu(1.2), emu(1.2), theme["line"], 9525, "ellipse"))
        shape_id += 1
        return add_role_background(shapes, slide or {}, theme, shape_id)

    if theme.get("template_style") == "mint_dotted_business":
        shapes.append(pattern_shape_xml(shape_id, "Mint Dot Background", emu(0.38), emu(0.25), emu(12.58), emu(6.62), theme["accent"], theme.get("dot_color", "FFFFFF"), "pct10", "roundRect"))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Coral Outer Frame", emu(0.36), emu(0.24), emu(12.62), emu(6.64), theme.get("frame_color", "F27C70"), 57150, "roundRect"))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Dotted Inner Frame", emu(0.52), emu(0.42), emu(12.28), emu(6.28), "FFFFFF", 6350, "roundRect"))
        shape_id += 1
        return add_role_background(shapes, slide or {}, theme, shape_id)

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
    return add_role_background(shapes, slide or {}, theme, shape_id)


def add_role_header(shapes, slide, theme, shape_id):
    layout = slide.get("layout", "bullets")
    kicker = clamp_text(slide.get("kicker", ""), 54)
    title = clamp_text(slide.get("title", "Untitled"), 58)
    if layout == "metric_dashboard":
        shapes.append(text_box_xml(shape_id, "Metric Eyebrow Pill", emu(0.85), emu(0.58), emu(2.55), emu(0.36), [paragraph_xml(kicker, FONT["caption"], "FFFFFF", True, "center")], fill=theme["primary"], line=theme["primary"], radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Metric Statement", emu(0.85), emu(1.0), emu(6.9), emu(0.62), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        return shape_id + 1
    if layout == "bar_comparison":
        shapes.append(shape_xml(shape_id, "Bar Header Rail", emu(0.55), emu(0.46), emu(0.16), emu(1.2), theme["primary"], theme["primary"]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Bar Section Label", emu(0.88), emu(0.46), emu(3.25), emu(0.28), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Bar Claim Title", emu(0.86), emu(0.78), emu(8.65), emu(0.52), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        return shape_id + 1
    if layout == "cause_effect":
        shapes.append(text_box_xml(shape_id, "Cascade Vertical Label", emu(0.62), emu(0.92), emu(2.4), emu(0.34), [paragraph_xml(kicker, FONT["label"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cascade Claim", emu(0.62), emu(1.28), emu(3.05), emu(1.08), [paragraph_xml(title, FONT["body_large"], theme["text"], True)]))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Cascade Claim Pin", emu(0.62), emu(2.58), emu(0.62), emu(0.12), theme["primary"], theme["primary"], "roundRect"))
        return shape_id + 1
    if layout == "risk_matrix":
        shapes.append(text_box_xml(shape_id, "Risk Label Flag", emu(0.82), emu(0.5), emu(2.65), emu(0.36), [paragraph_xml(kicker, FONT["caption"], "FFFFFF", True, "center")], fill=theme["warning"], line=theme["warning"], radius=False, anchor="ctr", margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk Claim Title", emu(0.82), emu(0.98), emu(8.4), emu(0.58), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Risk Header Dash", emu(9.35), emu(1.26), emu(12.1), emu(1.26), theme["warning"], 9525, "dash"))
        return shape_id + 1
    if layout == "comparison":
        shapes.append(text_box_xml(shape_id, "Comparison Header Left", emu(0.85), emu(0.54), emu(2.4), emu(0.3), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Comparison Main Claim", emu(0.85), emu(0.92), emu(11.2), emu(0.58), [paragraph_xml(title, FONT["section"], theme["text"], True, "center")], anchor="ctr", margin_y=0))
        return shape_id + 1
    if layout == "takeaway":
        shapes.append(text_box_xml(shape_id, "Takeaway Small Label", emu(0.92), emu(0.72), emu(3.3), emu(0.32), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Big Claim", emu(0.9), emu(1.12), emu(8.8), emu(0.8), [paragraph_xml(title, FONT["display"], theme["text"], True)]))
        return shape_id + 1
    return None


def add_header(shapes, slide, theme, shape_id):
    role_header = add_role_header(shapes, slide, theme, shape_id)
    if role_header is not None:
        return role_header

    if theme.get("template_style") == "blue_ribbon_business":
        navy = theme.get("frame_color", theme["accent"])
        shapes.append(shape_xml(shape_id, "Header Navy Ribbon", emu(0.62), emu(0.37), emu(11.75), emu(0.5), navy, navy, "roundRect"))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Header Blue Badge", emu(0.45), emu(0.35), emu(0.72), emu(0.55), theme["primary"], theme["primary"], "roundRect"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Header Number", emu(0.58), emu(0.46), emu(0.42), emu(0.25), [paragraph_xml("1", FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Title", emu(1.28), emu(0.43), emu(10.75), emu(0.32), [paragraph_xml(clamp_text(slide.get("title", "PPT PRESENTATION"), 42), FONT["body"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Header Blue Rule", emu(0.68), emu(1.08), emu(12.1), emu(1.08), theme["primary"], 12700))
        return shape_id + 1

    if theme.get("template_style") == "mint_dotted_business":
        shapes.append(outlined_shape_xml(shape_id, "Header Number Badge", emu(0.55), emu(0.42), emu(0.64), emu(0.36), theme["primary"], 12700, "roundRect", theme.get("title_fill", "FFFFFF")))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Header Number", emu(0.68), emu(0.46), emu(0.26), emu(0.22), [paragraph_xml("1", FONT["micro"], theme["primary"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Header Capsule", emu(1.08), emu(0.42), emu(11.15), emu(0.36), theme["primary"], 12700, "roundRect", theme.get("title_fill", "FFFFFF")))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Header Dotted Rule", emu(1.28), emu(0.74), emu(11.95), emu(0.74), theme["primary"], 6350, "dash"))
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Title",
                emu(1.32),
                emu(0.42),
                emu(10.5),
                emu(0.34),
                [paragraph_xml(clamp_text(slide.get("title", "PPT PRESENTATION"), 44), FONT["body"], theme["primary"], True, "center")],
                anchor="ctr",
                margin_x=0,
                margin_y=0,
            )
        )
        return shape_id + 1

    layout = slide.get("layout", "bullets")
    kicker = clamp_text(slide.get("kicker", ""), 54)
    title = clamp_text(slide.get("title", "Untitled"), 58)
    if layout == "metric_dashboard":
        shapes.append(text_box_xml(shape_id, "Metric Eyebrow Pill", emu(0.85), emu(0.58), emu(2.55), emu(0.36), [paragraph_xml(kicker, FONT["caption"], "FFFFFF", True, "center")], fill=theme["primary"], line=theme["primary"], radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Metric Statement", emu(0.85), emu(1.0), emu(6.9), emu(0.62), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        return shape_id + 1
    if layout == "bar_comparison":
        shapes.append(shape_xml(shape_id, "Bar Header Rail", emu(0.55), emu(0.46), emu(0.16), emu(1.2), theme["primary"], theme["primary"]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Bar Section Label", emu(0.88), emu(0.46), emu(3.25), emu(0.28), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Bar Claim Title", emu(0.86), emu(0.78), emu(8.65), emu(0.52), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        return shape_id + 1
    if layout == "cause_effect":
        shapes.append(text_box_xml(shape_id, "Cascade Vertical Label", emu(0.62), emu(0.92), emu(2.4), emu(0.34), [paragraph_xml(kicker, FONT["label"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cascade Claim", emu(0.62), emu(1.28), emu(3.05), emu(1.08), [paragraph_xml(title, FONT["body_large"], theme["text"], True)]))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Cascade Claim Pin", emu(0.62), emu(2.58), emu(0.62), emu(0.12), theme["primary"], theme["primary"], "roundRect"))
        return shape_id + 1
    if layout == "risk_matrix":
        shapes.append(text_box_xml(shape_id, "Risk Label Flag", emu(0.82), emu(0.5), emu(2.65), emu(0.36), [paragraph_xml(kicker, FONT["caption"], "FFFFFF", True, "center")], fill=theme["warning"], line=theme["warning"], radius=False, anchor="ctr", margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk Claim Title", emu(0.82), emu(0.98), emu(8.4), emu(0.58), [paragraph_xml(title, FONT["section"], theme["text"], True)]))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Risk Header Dash", emu(9.35), emu(1.26), emu(12.1), emu(1.26), theme["warning"], 9525, "dash"))
        return shape_id + 1
    if layout == "comparison":
        shapes.append(text_box_xml(shape_id, "Comparison Header Left", emu(0.85), emu(0.54), emu(2.4), emu(0.3), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Comparison Main Claim", emu(0.85), emu(0.92), emu(11.2), emu(0.58), [paragraph_xml(title, FONT["section"], theme["text"], True, "center")], anchor="ctr", margin_y=0))
        return shape_id + 1
    if layout == "takeaway":
        shapes.append(text_box_xml(shape_id, "Takeaway Small Label", emu(0.92), emu(0.72), emu(3.3), emu(0.32), [paragraph_xml(kicker, FONT["caption"], theme["primary"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Big Claim", emu(0.9), emu(1.12), emu(8.8), emu(0.8), [paragraph_xml(title, FONT["display"], theme["text"], True)]))
        return shape_id + 1

    shapes.append(text_box_xml(shape_id, "Kicker", emu(0.65), emu(0.34), emu(5.6), emu(0.35), [paragraph_xml(kicker, FONT["label"], theme["primary"], True)]))
    shape_id += 1
    shapes.append(text_box_xml(shape_id, "Title", emu(0.65), emu(0.72), emu(11.85), emu(0.72), [paragraph_xml(title, FONT["display"], theme["text"], True)]))
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
    bullets = [safe_text(b) for b in slide.get("bullets", []) if safe_text(b)]
    if not bullets:
        return shape_id
    density = content_density(bullets)
    variant = layout_variant(theme, slide.get("layout", "bullets"), "compact_rows", slide)
    if max_items is None and variant in {"tile_grid", "spotlight_list"} and len(bullets) <= 4:
        slots = grid_slots(len(bullets), 0.85, start_y, 11.65, max(1.4, bottom_y - start_y), gap=0.28, prefer_columns=2 if len(bullets) > 1 else 1)
        for idx, (bullet, (x, y, w, h)) in enumerate(zip(bullets, slots)):
            accent = theme["primary"] if idx % 2 == 0 else theme["accent"]
            font = fit_font(bullet, w - 0.65, base=FONT["body_large"], minimum=FONT["body_small"], maximum=FONT["section"])
            tile_h = max(0.82, min(h, 1.38 if variant == "tile_grid" else 1.08))
            shapes.append(text_box_xml(shape_id, "Bullet Tile", emu(x), emu(y), emu(w), emu(tile_h), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            if variant == "spotlight_list":
                shapes.append(shape_xml(shape_id, "Bullet Tile Badge", emu(x + 0.22), emu(y + 0.22), emu(0.38), emu(0.38), accent, accent, "ellipse"))
                shape_id += 1
                text_x, text_w = x + 0.82, w - 1.02
            else:
                shapes.append(shape_xml(shape_id, "Bullet Tile Accent", emu(x), emu(y), emu(0.08), emu(tile_h), accent, accent))
                shape_id += 1
                text_x, text_w = x + 0.3, w - 0.5
            shapes.append(text_box_xml(shape_id, "Bullet Tile Text", emu(text_x), emu(y + 0.15), emu(text_w), emu(max(0.35, tile_h - 0.25)), [paragraph_xml(clamp_text(bullet, compact_text_limit(text_w, font)), font, theme["text"], True)]))
            shape_id += 1
        return shape_id

    y = start_y
    bullet_style = theme.get("bullet_style", "cards")
    fit_count = max(0, int((bottom_y - start_y + 0.001) // step))
    count = min(len(bullets), max_items if max_items is not None else 5, fit_count)
    if count == 0 and bullets:
        count = 1
        box_h = min(box_h, max(0.42, bottom_y - start_y))
    font_size = list_font(bullets[:count], 10.5, base=font_size, minimum=FONT["caption"])
    text_limit = min(text_limit, compact_text_limit(10.5, font_size))
    for idx, bullet in enumerate(bullets[:count]):
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

    treatment = slide.get("visual_treatment", "standard")
    metric_texts = [f"{c.get('label', '')} {c.get('value', '')} {c.get('delta', '')} {c.get('context', '')}" for c in cards]
    density = content_density(metric_texts)
    if len(cards) >= 3 and density == "high":
        slots = grid_slots(len(cards), 0.9, 1.92, 11.6, 3.72, gap=0.22, prefer_columns=3)
        for idx, (card, (x, y, w, h)) in enumerate(zip(cards, slots)):
            accent = theme_color(theme, card.get("emphasis", "primary"))
            value_font = fit_font(card.get("value", ""), w - 0.48, base=FONT["display"], minimum=FONT["section"], maximum=FONT["hero"])
            body_font = list_font([card.get("label", ""), card.get("delta", ""), card.get("context", "")], w - 0.48, base=FONT["caption"], minimum=FONT["micro"])
            shapes.append(text_box_xml(shape_id, "Metric Compact Card", emu(x), emu(y), emu(w), emu(h), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Metric Compact Accent", emu(x), emu(y), emu(w), emu(0.1), accent, accent, "rect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Compact Label", emu(x + 0.22), emu(y + 0.22), emu(w - 0.44), emu(0.32), [paragraph_xml(clamp_text(card.get("label", ""), compact_text_limit(w - 0.44, body_font)), body_font, theme["muted"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Compact Value", emu(x + 0.22), emu(y + 0.68), emu(w - 0.44), emu(0.62), [paragraph_xml(clamp_text(card.get("value", ""), compact_text_limit(w - 0.44, value_font)), value_font, accent, True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Compact Note", emu(x + 0.22), emu(y + 1.42), emu(w - 0.44), emu(0.72), [paragraph_xml(clamp_text(card.get("delta", ""), compact_text_limit(w - 0.44, body_font)), body_font, theme["text"], True), paragraph_xml(clamp_text(card.get("context", ""), compact_text_limit(w - 0.44, FONT["micro"])), FONT["micro"], theme["muted"])]))
            shape_id += 1
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.9, max_items=1, bottom_y=6.25, box_h=0.36, step=0.42, font_size=FONT["caption"], text_limit=56)

    if treatment == "hero_metric_strip" and len(cards) >= 3:
        main = cards[0]
        accent = theme_color(theme, main.get("emphasis", "primary"))
        shapes.append(text_box_xml(shape_id, "Metric Story Hero", emu(0.85), emu(1.92), emu(4.45), emu(3.18), [], fill=theme["surface_alt"], line=accent, radius=True))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Metric Story Accent", emu(0.85), emu(1.92), emu(0.16), emu(3.18), accent, accent))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Metric Story Label", emu(1.22), emu(2.22), emu(3.7), emu(0.38), [paragraph_xml(clamp_text(main.get("label", ""), 26), FONT["label"], theme["muted"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Metric Story Value", emu(1.18), emu(2.72), emu(3.9), emu(0.86), [paragraph_xml(clamp_text(main.get("value", ""), 18), 3800, accent, True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Metric Story Context", emu(1.22), emu(3.82), emu(3.75), emu(0.7), [paragraph_xml(clamp_text(main.get("delta", ""), 34), FONT["body"], theme["text"], True), paragraph_xml(clamp_text(main.get("context", ""), 48), FONT["caption"], theme["muted"])]))
        shape_id += 1
        y = 2.05
        for card in cards[1:]:
            accent = theme_color(theme, card.get("emphasis", "accent"))
            shapes.append(text_box_xml(shape_id, "Metric Story Secondary", emu(5.75), emu(y), emu(5.85), emu(0.98), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Metric Story Dot", emu(6.05), emu(y + 0.34), emu(0.18), emu(0.18), accent, accent, "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Story Secondary Text", emu(6.48), emu(y + 0.16), emu(4.75), emu(0.48), [paragraph_xml(f"{clamp_text(card.get('label', ''), 18)}  {clamp_text(card.get('value', ''), 14)}", FONT["section"], accent, True), paragraph_xml(clamp_text(card.get("context", ""), 46), FONT["caption"], theme["muted"])]))
            shape_id += 1
            y += 1.22
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.35, max_items=1, box_h=0.48, step=0.56, font_size=1200, text_limit=62)

    variant = layout_variant(theme, "metric_dashboard", "strip_cards", slide)
    if variant == "radial_bubbles":
        navy = theme.get("frame_color", theme["accent"])
        center_y = 3.02
        x_positions = [2.0, 5.48, 8.96]
        for idx, card in enumerate(cards):
            accent = theme_color(theme, card.get("emphasis", "primary"))
            x = x_positions[idx]
            shapes.append(shape_xml(shape_id, "Radial Bubble Outer", emu(x), emu(center_y - 0.92), emu(1.85), emu(1.85), theme["surface_alt"], theme["line"], "ellipse"))
            shape_id += 1
            shapes.append(outlined_shape_xml(shape_id, "Radial Bubble Arc", emu(x + 0.08), emu(center_y - 0.84), emu(1.69), emu(1.69), accent, 28575, "arc"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Radial Bubble Dot", emu(x + 1.34), emu(center_y - 0.68), emu(0.28), emu(0.28), theme["warning"], "FFFFFF", "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Radial Value", emu(x + 0.24), emu(center_y - 0.2), emu(1.38), emu(0.38), [paragraph_xml(clamp_text(card.get("value", ""), 10), FONT["body_large"], navy, True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Radial Label", emu(x - 0.38), emu(center_y + 1.08), emu(2.62), emu(0.36), [paragraph_xml(clamp_text(card.get("label", ""), 22), FONT["label"], theme["text"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Radial Context", emu(x - 0.62), emu(center_y + 1.48), emu(3.1), emu(0.46), [paragraph_xml(clamp_text(card.get("delta") or card.get("context", ""), 42), FONT["micro"], theme["muted"], False, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        return shape_id

    if variant == "donut_metrics":
        x_positions = [1.55, 5.05, 8.55]
        for idx, card in enumerate(cards):
            accent = theme_color(theme, card.get("emphasis", "primary"))
            x = x_positions[idx]
            shapes.append(shape_xml(shape_id, "Donut Track", emu(x), emu(2.1), emu(1.75), emu(1.75), "FFFFFF", theme["primary"], "donut"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Donut Accent", emu(x + 1.08), emu(1.9), emu(0.52), emu(0.26), theme.get("progress_fill", theme["warning"]), theme["primary"], "parallelogram"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Donut Value", emu(x + 0.32), emu(2.68), emu(1.12), emu(0.42), [paragraph_xml(clamp_text(card.get("value", ""), 8), FONT["body_large"], theme["text"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Donut Label", emu(x - 0.25), emu(4.22), emu(2.35), emu(0.36), [paragraph_xml(clamp_text(card.get("label", ""), 22), FONT["label"], theme["text"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Donut Context", emu(x - 0.38), emu(4.6), emu(2.6), emu(0.46), [paragraph_xml(clamp_text(card.get("delta") or card.get("context", ""), 42), FONT["micro"], theme["muted"], False, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        return shape_id

    if variant in {"scoreboard", "alert_cards"}:
        main = cards[0]
        accent = theme_color(theme, main.get("emphasis", "primary"))
        shapes.append(text_box_xml(shape_id, "Hero Metric", emu(0.85), emu(1.95), emu(5.35), emu(2.35), [], fill=theme["surface_alt"], line=accent, radius=True))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Hero Metric Stripe", emu(0.85), emu(1.95), emu(5.35), emu(0.16), accent, accent))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Label", emu(1.2), emu(2.2), emu(4.6), emu(0.34), [paragraph_xml(clamp_text(main.get("label", ""), 28), FONT["label"], theme["muted"], True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Value", emu(1.2), emu(2.58), emu(4.7), emu(0.8), [paragraph_xml(clamp_text(main.get("value", ""), 22), FONT["hero"], accent, True)]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Hero Metric Context", emu(1.2), emu(3.48), emu(4.6), emu(0.52), [paragraph_xml(clamp_text(main.get("delta", ""), 34), FONT["body"], theme["text"], True), paragraph_xml(clamp_text(main.get("context", ""), 42), FONT["micro"], theme["muted"])]))
        shape_id += 1
        y = 1.95
        for card in cards[1:]:
            accent = theme_color(theme, card.get("emphasis", "primary"))
            shapes.append(text_box_xml(shape_id, "Side Metric", emu(6.55), emu(y), emu(5.65), emu(1.08), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Side Metric Dot", emu(6.82), emu(y + 0.28), emu(0.18), emu(0.18), accent, accent, "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Side Metric Text", emu(7.12), emu(y + 0.14), emu(4.8), emu(0.72), [paragraph_xml(clamp_text(card.get("label", ""), 26), FONT["caption"], theme["muted"], True), paragraph_xml(clamp_text(card.get("value", ""), 24), FONT["section"], accent, True)]))
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
            shapes.append(text_box_xml(shape_id, "Metric Row Label", emu(1.25), emu(y + 0.16), emu(3.2), emu(0.38), [paragraph_xml(clamp_text(card.get("label", ""), 28), FONT["body_small"], theme["text"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Row Value", emu(4.65), emu(y + 0.08), emu(3.0), emu(0.44), [paragraph_xml(clamp_text(card.get("value", ""), 24), FONT["section"], accent, True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Metric Row Context", emu(7.65), emu(y + 0.16), emu(4.0), emu(0.38), [paragraph_xml(clamp_text(card.get("delta") or card.get("context", ""), 46), FONT["caption"], theme["muted"])]))
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
    treatment = slide.get("visual_treatment", "standard")
    variant = layout_variant(theme, "bar_comparison", "horizontal_bars", slide)

    if len(charts) > 1:
        slots = grid_slots(min(len(charts), 2), 0.85, 1.88, 11.65, 3.9, gap=0.35, prefer_columns=2)
        for chart_idx, (chart, (panel_x, panel_y, panel_w, panel_h)) in enumerate(zip(charts[:2], slots)):
            data = chart.get("data", [])[:4]
            values = [as_number(item.get("value", 0)) for item in data]
            max_value = max(values) if values else 1
            accent = theme_color(theme, chart.get("emphasis", "primary" if chart_idx == 0 else "accent"))
            shapes.append(text_box_xml(shape_id, "Small Multiple Panel", emu(panel_x), emu(panel_y), emu(panel_w), emu(panel_h), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            title_font = fit_font(chart.get("title", "비교 지표"), panel_w - 0.5, base=FONT["label"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Small Multiple Title", emu(panel_x + 0.25), emu(panel_y + 0.22), emu(panel_w - 0.5), emu(0.32), [paragraph_xml(clamp_text(chart.get("title", "비교 지표"), compact_text_limit(panel_w - 0.5, title_font)), title_font, accent, True)]))
            shape_id += 1
            bar_y = panel_y + 0.78
            label_font = list_font([item.get("label", "") for item in data], 1.45, base=FONT["caption"], minimum=FONT["micro"])
            for item in data:
                label = clamp_text(item.get("label", ""), compact_text_limit(1.45, label_font))
                value = as_number(item.get("value", 0))
                unit = safe_text(chart.get("unit", ""))
                bar_w = max(0.12, (panel_w - 2.45) * value / max_value)
                shapes.append(text_box_xml(shape_id, "Small Bar Label", emu(panel_x + 0.25), emu(bar_y - 0.02), emu(1.45), emu(0.28), [paragraph_xml(label, label_font, theme["text"], True)]))
                shape_id += 1
                shapes.append(shape_xml(shape_id, "Small Bar Track", emu(panel_x + 1.85), emu(bar_y + 0.06), emu(panel_w - 2.55), emu(0.18), "E5E7EB", None, "roundRect"))
                shape_id += 1
                shapes.append(shape_xml(shape_id, "Small Bar Value", emu(panel_x + 1.85), emu(bar_y + 0.06), emu(bar_w), emu(0.18), accent, accent, "roundRect"))
                shape_id += 1
                shapes.append(text_box_xml(shape_id, "Small Bar Number", emu(panel_x + panel_w - 0.75), emu(bar_y - 0.03), emu(0.6), emu(0.26), [paragraph_xml(f"{value:g}{unit}", FONT["micro"], accent, True, "right")]))
                shape_id += 1
                bar_y += 0.55
        return render_bullets(shapes, slide, theme, shape_id, start_y=6.02, max_items=1, bottom_y=6.35, box_h=0.32, step=0.36, font_size=FONT["caption"], text_limit=58)

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

    if treatment == "ranked_bars" and data:
        sorted_data = sorted(data, key=lambda item: as_number(item.get("value", 0)), reverse=True)
        accent = theme_color(theme, chart.get("emphasis", "primary"))
        unit = safe_text(chart.get("unit", ""))
        y = 2.34
        bar_x, bar_max_w = 3.05, 6.9
        for idx, item in enumerate(sorted_data[:5]):
            value = as_number(item.get("value", 0))
            bar_w = max(0.16, bar_max_w * value / max_value)
            fill = accent if idx == 0 else (theme["accent"] if idx == 1 else theme["primary"])
            alpha_fill = theme["surface_alt"] if idx == 0 else theme["surface"]
            shapes.append(text_box_xml(shape_id, "Rank Row Background", emu(0.9), emu(y - 0.05), emu(11.35), emu(0.5), [], fill=alpha_fill, line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Rank Number", emu(1.08), emu(y + 0.05), emu(0.42), emu(0.24), [paragraph_xml(str(idx + 1), FONT["micro"], fill, True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Rank Label", emu(1.62), emu(y), emu(1.65), emu(0.32), [paragraph_xml(clamp_text(item.get("label", ""), 16), FONT["body_small"], theme["text"], True)]))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Rank Bar Track", emu(bar_x), emu(y + 0.07), emu(bar_max_w), emu(0.24), "E5E7EB", None, "roundRect"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Rank Bar Value", emu(bar_x), emu(y + 0.07), emu(bar_w), emu(0.24), fill, fill, "roundRect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Rank Value", emu(bar_x + bar_max_w + 0.22), emu(y), emu(1.05), emu(0.32), [paragraph_xml(f"{value:g}{unit}", FONT["body_small"], fill, True)]))
            shape_id += 1
            y += 0.66
        top = sorted_data[0]
        shapes.append(text_box_xml(shape_id, "Rank Insight", emu(8.45), emu(5.42), emu(3.45), emu(0.52), [paragraph_xml(f"최대 항목: {clamp_text(top.get('label', ''), 16)}", FONT["label"], accent, True), paragraph_xml(f"{as_number(top.get('value', 0)):g}{unit}", FONT["section"], theme["text"], True)], fill=theme["surface"], line=accent, radius=True))
        return shape_id + 1

    if variant == "blue_progress_rows":
        y = 2.1
        navy = theme.get("frame_color", theme["accent"])
        accent = theme_color(theme, chart.get("emphasis", "primary"))
        for idx, item in enumerate(data[:3]):
            label = clamp_text(item.get("label", ""), 22)
            value = float(item.get("value", 0))
            pct = int(round((value / max_value) * 100)) if max_value else 0
            shapes.append(shape_xml(shape_id, "Blue Row Badge", emu(1.05), emu(y - 0.04), emu(0.78), emu(0.78), theme["surface_alt"], theme["primary"], "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Blue Row Number", emu(1.22), emu(y + 0.17), emu(0.44), emu(0.24), [paragraph_xml(str(idx + 1), FONT["micro"], navy, True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Blue Row Bar", emu(2.42), emu(y + 0.13), emu(3.75), emu(0.34), [], fill="FFFFFF", line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Blue Row Fill", emu(2.48), emu(y + 0.2), emu(max(0.18, 3.38 * value / max_value)), emu(0.2), theme.get("progress_fill", theme["surface_alt"]), None, "roundRect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Blue Row Percent", emu(4.78), emu(y + 0.12), emu(0.82), emu(0.26), [paragraph_xml(f"{pct}%", FONT["micro"], navy, True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(line_segment_xml(shape_id, "Blue Row Connector", emu(6.28), emu(y + 0.31), emu(7.02), emu(y + 0.31), navy, 6350, "dash"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Blue Content Dot", emu(7.18), emu(y + 0.22), emu(0.18), emu(0.18), theme["warning"], theme["warning"], "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Blue Row Label", emu(7.55), emu(y), emu(3.95), emu(0.32), [paragraph_xml("CONTENTS A", FONT["micro"], theme["text"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Blue Row Detail", emu(7.55), emu(y + 0.34), emu(4.25), emu(0.36), [paragraph_xml(label, FONT["micro"], theme["muted"])]))
            shape_id += 1
            y += 1.12
        return shape_id

    if variant == "progress_rows":
        y = 2.0
        accent = theme_color(theme, chart.get("emphasis", "primary"))
        for idx, item in enumerate(data[:3]):
            label = clamp_text(item.get("label", ""), 18)
            value = float(item.get("value", 0))
            pct = int(round((value / max_value) * 100)) if max_value else 0
            shapes.append(outlined_shape_xml(shape_id, "Icon Circle", emu(1.25), emu(y + 0.05), emu(0.56), emu(0.56), theme["primary"], 12700, "ellipse", theme["surface_alt"]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Icon Glyph", emu(1.39), emu(y + 0.18), emu(0.25), emu(0.2), [paragraph_xml(str(idx + 1), FONT["micro"], theme["primary"], True)]))
            shape_id += 1
            shapes.append(outlined_shape_xml(shape_id, "Progress Capsule", emu(3.0), emu(y + 0.17), emu(3.35), emu(0.28), theme["primary"], 9525, "roundRect", "FFFFFF"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Progress Fill", emu(3.02), emu(y + 0.19), emu(max(0.18, 3.08 * value / max_value)), emu(0.18), theme.get("progress_fill", "F6E87A"), None, "roundRect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Progress Value", emu(4.72), emu(y + 0.14), emu(0.78), emu(0.26), [paragraph_xml(f"{pct}%", FONT["micro"], theme["primary"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(line_segment_xml(shape_id, "Progress Dotted Connector", emu(6.4), emu(y + 0.31), emu(7.2), emu(y + 0.31), theme["primary"], 6350, "dash"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Progress Content", emu(7.45), emu(y + 0.02), emu(3.9), emu(0.3), [paragraph_xml("CONTENTS A", FONT["micro"], theme["text"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Progress Detail", emu(7.45), emu(y + 0.34), emu(4.0), emu(0.36), [paragraph_xml(label, FONT["micro"], theme["muted"])]))
            shape_id += 1
            y += 1.15
        return shape_id

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
    treatment = slide.get("visual_treatment", "standard")
    variant = layout_variant(theme, "line_trend", "wide_plot_callout", slide)

    if variant in {"sparkline_stack", "small_multiples"} and len(charts) > 1:
        slots = grid_slots(min(len(charts), 3), 0.85, 1.86, 11.65, 3.85, gap=0.25, prefer_columns=1)
        for chart_idx, (chart, (panel_x, panel_y, panel_w, panel_h)) in enumerate(zip(charts[:3], slots)):
            data = chart.get("data", [])[:6]
            if len(data) < 2:
                continue
            accent = theme_color(theme, chart.get("emphasis", "primary" if chart_idx == 0 else "accent"))
            values = [as_number(item.get("value", 0)) for item in data]
            min_v, max_v = min(values), max(values)
            span = max(max_v - min_v, 1)
            shapes.append(text_box_xml(shape_id, "Sparkline Row", emu(panel_x), emu(panel_y), emu(panel_w), emu(panel_h), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            title_font = fit_font(chart.get("title", "추세"), 2.8, base=FONT["label"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Sparkline Title", emu(panel_x + 0.25), emu(panel_y + 0.22), emu(2.9), emu(0.32), [paragraph_xml(clamp_text(chart.get("title", "추세"), compact_text_limit(2.8, title_font)), title_font, accent, True)]))
            shape_id += 1
            x0, x1 = panel_x + 3.35, panel_x + panel_w - 1.3
            y0, h0 = panel_y + 0.28, panel_h - 0.52
            points = []
            for idx, item in enumerate(data):
                x = x0 + (x1 - x0) * idx / (len(data) - 1)
                y = y0 + h0 * (1 - (as_number(item.get("value", 0)) - min_v) / span)
                points.append((x, y, item))
            for (x_a, y_a, _), (x_b, y_b, __) in zip(points, points[1:]):
                shapes.append(line_segment_xml(shape_id, "Sparkline Segment", emu(x_a), emu(y_a), emu(x_b), emu(y_b), accent, 19050))
                shape_id += 1
            last = points[-1][2]
            unit = safe_text(chart.get("unit", ""))
            shapes.append(text_box_xml(shape_id, "Sparkline Last Value", emu(panel_x + panel_w - 1.05), emu(panel_y + 0.22), emu(0.82), emu(0.32), [paragraph_xml(f"{as_number(last.get('value', 0)):g}{unit}", FONT["caption"], accent, True, "right")], margin_x=0, margin_y=0))
            shape_id += 1
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.95, max_items=1, bottom_y=6.25, box_h=0.32, step=0.36, font_size=FONT["caption"], text_limit=58)

    panel_x, panel_y, panel_w, panel_h = 0.85, 1.82, 11.65, 3.35
    if variant in {"milestone_line", "compact_trend"}:
        panel_y, panel_h = 2.2, 2.65
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
    chart_x, chart_y, chart_w, chart_h = panel_x + 0.6, panel_y + 0.82, 10.35, 1.92 if variant not in {"milestone_line", "compact_trend"} else 1.35
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
    for point_idx, (x, y, item) in enumerate(points):
        shapes.append(shape_xml(shape_id, "Trend Marker", emu(x - 0.06), emu(y - 0.06), emu(0.12), emu(0.12), accent, "FFFFFF", "ellipse"))
        shape_id += 1
        if variant == "milestone_line" and point_idx in {0, len(points) - 1}:
            shapes.append(text_box_xml(shape_id, "Milestone Value", emu(x - 0.4), emu(y - 0.45), emu(0.8), emu(0.28), [paragraph_xml(f"{as_number(item.get('value', 0)):g}{unit}", FONT["micro"], accent, True, "center")], margin_x=0, margin_y=0))
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
    if treatment in {"forecast_actual_line", "annotated_trend"}:
        last_x, last_y, last_item = points[-1]
        first_value = as_number(points[0][2].get("value", 0))
        last_value = as_number(last_item.get("value", 0))
        delta = last_value - first_value
        delta_prefix = "+" if delta > 0 else ""
        shapes.append(line_segment_xml(shape_id, "Trend Final Guide", emu(last_x), emu(chart_y - 0.08), emu(last_x), emu(chart_y + chart_h + 0.08), accent, 9525, "dash"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Trend Annotation", emu(8.25), emu(panel_y + 0.18), emu(3.3), emu(0.76), [paragraph_xml(clamp_text(last_item.get("label", "최종"), 18), FONT["caption"], theme["muted"], True), paragraph_xml(f"{last_value:g}{unit} ({delta_prefix}{delta:g}{unit})", FONT["body_large"], accent, True)], fill=theme["surface_alt"], line=accent, radius=True))
        shape_id += 1
    shapes.append(
        text_box_xml(
            shape_id,
            "Trend Range",
            emu(panel_x + 0.28),
            emu(panel_y + panel_h - 0.34),
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
        start_y=5.55,
        max_items=1,
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

    if len(events) > 4 or item_weight([e.get("label", "") + e.get("detail", "") for e in events]) > 130:
        slots = grid_slots(len(events), 0.9, 1.88, 11.65, 4.35, gap=0.24, prefer_columns=2)
        for idx, (event, (x, y, w, h)) in enumerate(zip(events, slots)):
            accent = theme["primary"] if idx % 2 == 0 else theme["accent"]
            shapes.append(text_box_xml(shape_id, "Timeline Grid Event", emu(x), emu(y), emu(w), emu(h), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Timeline Grid Accent", emu(x), emu(y), emu(0.08), emu(h), accent, accent))
            shape_id += 1
            title_font = list_font([event.get("label", ""), event.get("detail", "")], w - 0.55, base=FONT["body_small"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Timeline Grid Text", emu(x + 0.25), emu(y + 0.16), emu(w - 0.48), emu(h - 0.25), [paragraph_xml(clamp_text(event.get("time", ""), 16), FONT["micro"], accent, True), paragraph_xml(clamp_text(event.get("label", ""), compact_text_limit(w - 0.48, title_font)), title_font, theme["text"], True), paragraph_xml(clamp_text(event.get("detail", ""), compact_text_limit(w - 0.48, FONT["micro"])), FONT["micro"], theme["muted"])]))
            shape_id += 1
        return shape_id

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

    slots = grid_slots(len(steps), 0.85, 2.02, 11.65, 3.3, gap=0.28, prefer_columns=min(len(steps), 4))
    for idx, step in enumerate(steps):
        sx, y, w, h = slots[idx]
        accent = theme["primary"] if idx % 2 == 0 else theme["accent"]
        text_font = list_font([step.get("label", ""), step.get("detail", "")], w - 0.45, base=FONT["body_small"], minimum=FONT["caption"])
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
                    paragraph_xml(clamp_text(step.get("label", ""), compact_text_limit(w - 0.45, text_font)), text_font, theme["text"], True),
                    paragraph_xml(clamp_text(step.get("detail", ""), compact_text_limit(w - 0.45, FONT["micro"])), FONT["micro"], theme["muted"]),
                ],
                fill=theme["surface"],
                line=theme["line"],
                radius=True,
            )
        )
        shape_id += 1
        if idx < len(steps) - 1:
            next_x, next_y, _, _ = slots[idx + 1]
            if abs(next_y - y) < 0.1:
                shapes.append(shape_xml(shape_id, "Flow Arrow", emu(sx + w + 0.08), emu(y + h / 2 - 0.14), emu(0.28), emu(0.28), accent, accent, "triangle"))
            else:
                shapes.append(line_segment_xml(shape_id, "Flow Down Connector", emu(sx + w / 2), emu(y + h), emu(next_x + w / 2), emu(next_y), accent, 9525, "dash"))
            shape_id += 1
    return render_bullets(shapes, slide, theme, shape_id, start_y=4.18, max_items=2, box_h=0.52, step=0.62, font_size=1250, text_limit=62)


def render_comparison(shapes, slide, theme, shape_id):
    comps = [c for c in slide.get("components", []) if c.get("type") == "comparison"]
    left = comps[0].get("left", {}) if comps else {"title": "Before", "items": slide.get("bullets", [])[:3]}
    right = comps[0].get("right", {}) if comps else {"title": "After", "items": slide.get("bullets", [])[3:6]}
    variant = layout_variant(theme, "comparison", "split_columns", slide)
    max_items = max(len(left.get("items", [])), len(right.get("items", [])))
    dense = max_items > 3 or item_weight(left.get("items", []) + right.get("items", [])) > 140

    if variant in {"stacked_scorecards", "before_after_cards"}:
        rows = [(left, theme["primary"], 2.0), (right, theme["accent"], 4.05)]
        for idx, (col, accent, y) in enumerate(rows):
            items = col.get("items", [])[:4]
            font = list_font(items, 8.6, base=FONT["body_small"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Comparison Stacked Card", emu(0.95), emu(y), emu(11.35), emu(1.56), [], fill=theme["surface"], line=accent, radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Comparison Stacked Band", emu(0.95), emu(y), emu(0.16), emu(1.56), accent, accent))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Comparison Stacked Title", emu(1.35), emu(y + 0.18), emu(2.35), emu(0.38), [paragraph_xml(clamp_text(col.get("title", ""), 24), FONT["body_small"], accent, True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Comparison Stacked Items", emu(3.85), emu(y + 0.16), emu(7.85), emu(1.04), [paragraph_xml(" / ".join(clamp_text(item, compact_text_limit(2.0, font)) for item in items), font, theme["text"], True)]))
            shape_id += 1
            if idx == 0:
                shapes.append(shape_xml(shape_id, "Comparison Down Arrow", emu(6.2), emu(3.58), emu(0.38), emu(0.38), accent, accent, "downArrow"))
                shape_id += 1
        return shape_id

    if slide.get("visual_treatment") == "two_column_scorecard":
        band_h = 0.34 if dense else 0.42
        shapes.append(text_box_xml(shape_id, "Scorecard Band", emu(0.85), emu(1.82), emu(11.55), emu(band_h), [paragraph_xml("비교 관점", FONT["label"], theme["muted"], True, "center")], fill=theme["surface_alt"], line=theme["line"], radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
    columns = [(left, theme["primary"], 0.85), (right, theme["accent"], 6.75)]
    for col, accent, x in columns:
        if slide.get("visual_treatment") == "two_column_scorecard":
            shapes.append(shape_xml(shape_id, "Scorecard Top Rule", emu(x), emu(2.34), emu(5.55), emu(0.08), accent, accent))
            shape_id += 1
        items = col.get("items", [])[:5]
        font = list_font(items, 4.9, base=FONT["body_small"], minimum=FONT["caption"])
        panel_y = 2.3 if slide.get("visual_treatment") == "two_column_scorecard" else 1.95
        panel_h = 3.15 if dense else 2.45
        shapes.append(
            text_box_xml(
                shape_id,
                "Comparison Column",
                emu(x),
                emu(panel_y),
                emu(5.55),
                emu(panel_h),
                [paragraph_xml(clamp_text(col.get("title", ""), compact_text_limit(4.9, FONT["body"])), FONT["body"], accent, True)]
                + [paragraph_xml("• " + clamp_text(item, compact_text_limit(4.9, font)), font, theme["text"]) for item in items],
                fill=theme["surface"],
                line=accent,
                radius=True,
            )
        )
        shape_id += 1
    shapes.append(shape_xml(shape_id, "Comparison Divider", emu(6.48), emu(2.12), emu(0.08), emu(3.2 if dense else 2.55), theme["line"], None, "rect"))
    shape_id += 1
    return shape_id


def render_risk_matrix(shapes, slide, theme, shape_id):
    matrices = [c for c in slide.get("components", []) if c.get("type") == "risk_matrix"]
    matrix = matrices[0] if matrices else {}
    items = matrix.get("items", [])[:8]
    variant = layout_variant(theme, "risk_matrix", "quadrant_watchlist", slide)

    if variant == "ranked_watchlist" or (variant == "compact_quadrant" and len(items) >= 6):
        ranked = sorted(
            items,
            key=lambda item: as_number(item.get("likelihood", 0.5), 0.5) * as_number(item.get("impact", 0.5), 0.5),
            reverse=True,
        )
        slots = grid_slots(min(len(ranked), 6), 0.85, 1.88, 11.65, 4.45, gap=0.22, prefer_columns=2)
        for idx, (item, (x, y, w, h)) in enumerate(zip(ranked[:6], slots), 1):
            score = as_number(item.get("likelihood", 0.5), 0.5) * as_number(item.get("impact", 0.5), 0.5)
            color = theme_color(theme, item.get("emphasis", "warning"))
            fill = "FEE2E2" if score >= 0.55 else theme["surface"]
            label_font = fit_font(item.get("label", ""), w - 1.2, base=FONT["body_small"], minimum=FONT["caption"], maximum=FONT["body"])
            shapes.append(text_box_xml(shape_id, "Risk Ranked Card", emu(x), emu(y), emu(w), emu(h), [], fill=fill, line=color, radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Risk Ranked Badge", emu(x + 0.24), emu(y + 0.24), emu(0.42), emu(0.42), color, color, "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Risk Ranked Number", emu(x + 0.24), emu(y + 0.31), emu(0.42), emu(0.16), [paragraph_xml(str(idx), FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Risk Ranked Label", emu(x + 0.86), emu(y + 0.22), emu(w - 1.1), emu(0.42), [paragraph_xml(clamp_text(item.get("label", ""), compact_text_limit(w - 1.2, label_font)), label_font, theme["text"], True)]))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Risk Ranked Score", emu(x + 0.86), emu(y + 0.72), emu(w - 1.1), emu(0.3), [paragraph_xml(f"가능성 {as_number(item.get('likelihood', 0.5), 0.5):.1f} / 영향 {as_number(item.get('impact', 0.5), 0.5):.1f}", FONT["micro"], theme["muted"])]))
            shape_id += 1
        return shape_id

    if variant == "heatmap_focus":
        x, y, w, h = 0.95, 2.05, 11.35, 3.35
        zones = [
            ("관찰", x, y + h / 2, w / 2, h / 2, "F8FAFC"),
            ("주의", x + w / 2, y + h / 2, w / 2, h / 2, "FEF3C7"),
            ("상승 압력", x, y, w / 2, h / 2, "FEF3C7"),
            ("우선 대응", x + w / 2, y, w / 2, h / 2, "FEE2E2"),
        ]
        for label, cx, cy, cw, ch, fill in zones:
            shapes.append(text_box_xml(shape_id, "Risk Heatmap Zone", emu(cx), emu(cy), emu(cw), emu(ch), [paragraph_xml(label, FONT["label"], theme["muted"], True)], fill=fill, line=theme["line"], radius=True))
            shape_id += 1
        for idx, item in enumerate(items[:6], 1):
            likelihood = max(0, min(1, as_number(item.get("likelihood", 0.5), 0.5)))
            impact = max(0, min(1, as_number(item.get("impact", 0.5), 0.5)))
            px = x + 0.45 + (w - 0.9) * impact
            py = y + 0.45 + (h - 0.9) * (1 - likelihood)
            color = theme_color(theme, item.get("emphasis", "warning"))
            shapes.append(text_box_xml(shape_id, "Risk Heatmap Label", emu(px - 0.78), emu(py - 0.2), emu(1.56), emu(0.4), [paragraph_xml(clamp_text(item.get("label", ""), 14), FONT["micro"], "FFFFFF", True, "center")], fill=color, line="FFFFFF", radius=True, anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        return shape_id

    x, y, w, h = 0.95, 2.05, 7.55, 3.45
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
    shapes.append(text_box_xml(shape_id, "Axis X Label", emu(x + w - 1.35), emu(y + h + 0.12), emu(1.45), emu(0.25), [paragraph_xml("영향도 높음", FONT["micro"], theme["muted"], True)]))
    shape_id += 1
    shapes.append(text_box_xml(shape_id, "Axis Y Label", emu(x - 0.04), emu(y - 0.34), emu(1.55), emu(0.25), [paragraph_xml("가능성 높음", FONT["micro"], theme["muted"], True)]))
    shape_id += 1

    side_x = 8.85
    side_y = 2.05
    shapes.append(text_box_xml(shape_id, "Risk Watchlist Panel", emu(side_x), emu(side_y), emu(3.55), emu(3.48), [], fill=theme["surface"], line=theme["line"], radius=True))
    shape_id += 1
    shapes.append(text_box_xml(shape_id, "Risk Watchlist Title", emu(side_x + 0.28), emu(side_y + 0.22), emu(2.8), emu(0.32), [paragraph_xml("우선 점검 항목", FONT["label"], theme["muted"], True)]))
    shape_id += 1

    ranked = sorted(
        enumerate(items, 1),
        key=lambda pair: as_number(pair[1].get("likelihood", 0.5), 0.5) * as_number(pair[1].get("impact", 0.5), 0.5),
        reverse=True,
    )
    rank_by_original = {original_idx: rank for rank, (original_idx, _) in enumerate(ranked, 1)}
    for original_idx, item in enumerate(items, 1):
        likelihood = max(0, min(1, as_number(item.get("likelihood", 0.5), 0.5)))
        impact = max(0, min(1, as_number(item.get("impact", 0.5), 0.5)))
        px = x + 0.3 + (w - 0.6) * impact
        py = y + 0.3 + (h - 0.6) * (1 - likelihood)
        color = theme_color(theme, item.get("emphasis", "warning"))
        shapes.append(shape_xml(shape_id, "Risk Dot", emu(px - 0.12), emu(py - 0.12), emu(0.24), emu(0.24), color, "FFFFFF", "ellipse"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk Dot Number", emu(px - 0.12), emu(py - 0.09), emu(0.24), emu(0.16), [paragraph_xml(str(rank_by_original[original_idx]), FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1

    for rank, (_, item) in enumerate(ranked[:5], 1):
        row_y = side_y + 0.68 + (rank - 1) * 0.52
        color = theme_color(theme, item.get("emphasis", "warning"))
        shapes.append(shape_xml(shape_id, "Risk List Badge", emu(side_x + 0.28), emu(row_y + 0.04), emu(0.28), emu(0.28), color, color, "ellipse"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk List Number", emu(side_x + 0.28), emu(row_y + 0.075), emu(0.28), emu(0.16), [paragraph_xml(str(rank), FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Risk List Label", emu(side_x + 0.68), emu(row_y), emu(2.6), emu(0.38), [paragraph_xml(clamp_text(item.get("label", ""), 24), FONT["caption"], theme["text"], True)]))
        shape_id += 1
    return shape_id


def render_cause_effect(shapes, slide, theme, shape_id):
    comps = [c for c in slide.get("components", []) if c.get("type") == "cause_effect"]
    comp = comps[0] if comps else {}
    variant = layout_variant(theme, "cause_effect", "cascade_cards", slide)
    groups = [
        ("원인", comp.get("causes", slide.get("bullets", [])[:2]), theme["primary"], 3.35),
        ("전이", comp.get("events", slide.get("bullets", [])[2:4]), theme["warning"], 6.25),
        ("결과", comp.get("effects", slide.get("bullets", [])[4:6]), theme["accent"], 9.15),
    ]

    if variant == "vertical_story":
        y = 1.92
        for idx, (title, items, accent, _) in enumerate(groups):
            row_h = 1.08
            shapes.append(shape_xml(shape_id, "Story Rail Dot", emu(1.02), emu(y + 0.28), emu(0.28), emu(0.28), accent, "FFFFFF", "ellipse"))
            shape_id += 1
            if idx < len(groups) - 1:
                shapes.append(line_segment_xml(shape_id, "Story Rail", emu(1.16), emu(y + 0.56), emu(1.16), emu(y + 1.42), theme["line"], 12700, "dash"))
                shape_id += 1
            shapes.append(text_box_xml(shape_id, "Story Stage Label", emu(1.55), emu(y + 0.12), emu(1.05), emu(0.3), [paragraph_xml(title, FONT["label"], accent, True)]))
            shape_id += 1
            item_font = list_font(items[:3], 8.4, base=FONT["body_small"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Story Stage Card", emu(2.75), emu(y), emu(9.25), emu(row_h), [paragraph_xml(" · ".join(clamp_text(i, 24) for i in items[:3]), item_font, theme["text"], True)], fill=theme["surface"], line=accent, radius=True, anchor="ctr", margin_y=0))
            shape_id += 1
            y += 1.36
        return render_bullets(shapes, slide, theme, shape_id, start_y=5.9, max_items=1, bottom_y=6.25, box_h=0.32, step=0.36, font_size=FONT["caption"], text_limit=58)

    if variant == "split_swimlane":
        left_title, left_items, left_accent, _ = groups[0]
        mid_title, mid_items, mid_accent, _ = groups[1]
        right_title, right_items, right_accent, _ = groups[2]
        lanes = [
            (left_title, left_items, left_accent, 0.95),
            (right_title, right_items, right_accent, 8.15),
        ]
        for title, items, accent, x in lanes:
            font = list_font(items[:4], 3.55, base=FONT["body_small"], minimum=FONT["caption"])
            shapes.append(text_box_xml(shape_id, "Swimlane Panel", emu(x), emu(2.08), emu(3.85), emu(3.35), [paragraph_xml(title, FONT["label"], accent, True)] + [paragraph_xml("• " + clamp_text(i, compact_text_limit(3.35, font)), font, theme["text"]) for i in items[:4]], fill=theme["surface"], line=accent, radius=True))
            shape_id += 1
        shapes.append(text_box_xml(shape_id, "Swimlane Bridge", emu(4.55), emu(2.45), emu(3.25), emu(2.55), [paragraph_xml(mid_title, FONT["label"], mid_accent, True, "center")] + [paragraph_xml(clamp_text(i, 24), FONT["caption"], theme["text"], False, "center") for i in mid_items[:3]], fill=theme["surface_alt"], line=mid_accent, radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Swimlane Left Connector", emu(3.95), emu(3.72), emu(4.55), emu(3.72), left_accent, 12700, "dash"))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Swimlane Right Connector", emu(7.8), emu(3.72), emu(8.15), emu(3.72), right_accent, 12700, "dash"))
        shape_id += 1
        return shape_id

    for idx, (title, items, accent, x) in enumerate(groups):
        if slide.get("visual_treatment") == "cascade_flow":
            shapes.append(shape_xml(shape_id, "Cascade Stage Chip", emu(x), emu(2.0), emu(0.46), emu(0.28), accent, accent, "roundRect"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Cascade Stage Number", emu(x), emu(2.04), emu(0.46), emu(0.16), [paragraph_xml(f"{idx + 1}", FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Cause Effect Node",
                emu(x),
                emu(2.38),
                emu(2.45),
                emu(1.58),
                [paragraph_xml(title, FONT["label"], accent, True)] + [paragraph_xml("• " + clamp_text(i, 25), FONT["caption"], theme["text"]) for i in items[:3]],
                fill=theme["surface"],
                line=accent,
                radius=True,
            )
        )
        shape_id += 1
        if idx < 2:
            shapes.append(line_segment_xml(shape_id, "Cause Effect Connector", emu(x + 2.55), emu(3.14), emu(x + 2.86), emu(3.14), accent, 12700, "dash"))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Cause Effect Arrow", emu(x + 2.82), emu(2.98), emu(0.28), emu(0.28), accent, accent, "triangle"))
            shape_id += 1
    all_items = [safe_text(i) for _, items, _, _ in groups for i in items[:1] if safe_text(i)]
    if all_items:
        shapes.append(text_box_xml(shape_id, "Cascade Summary Strip", emu(3.35), emu(4.88), emu(8.6), emu(0.66), [paragraph_xml(" → ".join(clamp_text(i, 16) for i in all_items[:3]), FONT["body_small"], theme["text"], True, "center")], fill=theme["surface_alt"], line=theme["line"], radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
    return render_bullets(shapes, slide, theme, shape_id, start_y=5.78, max_items=1, bottom_y=6.22, box_h=0.4, step=0.46, font_size=FONT["caption"], text_limit=62)


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
    slots = grid_slots(len(nodes), 0.95, 2.0, 11.45, 3.75, gap=0.42, prefer_columns=4 if len(nodes) > 4 else len(nodes))
    positions = {
        "user": (0.95, 3.05),
        "harness": (3.75, 3.05),
        "model": (6.55, 2.2),
        "tools": (6.55, 3.8),
        "memory": (9.35, 2.2),
        "skills": (9.35, 3.8),
        "a2a": (9.35, 5.05),
    }
    node_pos = {}
    for idx, node in enumerate(nodes):
        node_id = safe_text(node.get("id", f"n{idx}"))
        node_pos[node_id] = positions.get(node_id, (slots[idx][0], slots[idx][1]))
    for edge in edges:
        a, b = node_pos.get(edge.get("from")), node_pos.get(edge.get("to"))
        if a and b:
            shapes.append(line_segment_xml(shape_id, "Architecture Edge", emu(a[0] + 0.7), emu(a[1] + 0.25), emu(b[0]), emu(b[1] + 0.25), theme["line"], 12700))
            shape_id += 1
    for idx, node in enumerate(nodes):
        node_id = safe_text(node.get("id", f"n{idx}"))
        x, y = node_pos[node_id]
        accent = theme_color(theme, node.get("emphasis", "primary" if idx == 1 else "accent"))
        label = safe_text(node.get("label", node_id))
        width = 1.75 if len(nodes) <= 5 else 1.55
        font = fit_font(label, width - 0.25, base=FONT["caption"], minimum=FONT["micro"], maximum=FONT["body_small"])
        shapes.append(text_box_xml(shape_id, "Architecture Node", emu(x), emu(y), emu(width), emu(0.72), [paragraph_xml(clamp_text(label, compact_text_limit(width - 0.25, font)), font, theme["text"], True, "center")], fill=theme["surface"], line=accent, radius=True, anchor="ctr", margin_y=0))
        shape_id += 1
    return shape_id


def render_callout_focus(shapes, slide, theme, shape_id):
    callouts = [c for c in slide.get("components", []) if c.get("type") == "callout"]
    callout = callouts[0] if callouts else {}
    headline = callout.get("headline") or slide.get("title", "")
    body = callout.get("body") or " ".join(slide.get("bullets", [])[:2])
    density = content_density([headline, body])
    panel_h = 2.25 if density == "low" else 3.1
    panel_y = 2.35 if density == "low" else 1.95
    shapes.append(translucent_shape_xml(shape_id, "Callout Accent", emu(0.9), emu(panel_y), emu(0.22), emu(panel_h), theme["primary"], 65000))
    shape_id += 1
    headline_font = fit_font(headline, 10.2, base=2500, minimum=FONT["section"], maximum=FONT["display"])
    body_font = fit_font(body, 10.2, base=FONT["body"], minimum=FONT["body_small"], maximum=FONT["body_large"])
    shapes.append(
        text_box_xml(
            shape_id,
            "Callout Panel",
            emu(1.25),
            emu(panel_y),
            emu(10.9),
            emu(panel_h),
            [
                paragraph_xml(clamp_text(headline, compact_text_limit(10.2, headline_font)), headline_font, theme["primary"], True),
                paragraph_xml(clamp_text(body, compact_text_limit(10.2, body_font) * 2), body_font, theme["text"]),
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
    variant = layout_variant(theme, "title_cover", "left_panel", slide)
    if variant == "blue_ribbon_title":
        navy = theme.get("frame_color", theme["accent"])
        shapes.append(shape_xml(shape_id, "Cover Navy Center Slab", emu(2.1), emu(2.35), emu(9.15), emu(1.38), navy, navy, "roundRect"))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Cover Blue Label Tab", emu(3.0), emu(2.05), emu(1.85), emu(0.38), theme["primary"], theme["primary"], "roundRect"))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Cover White Title Plate", emu(2.55), emu(2.64), emu(8.25), emu(0.72), "FFFFFF", 12700, "roundRect", "FFFFFF"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Title", emu(2.82), emu(2.72), emu(7.72), emu(0.48), [paragraph_xml(clamp_text(title, 38), 2200, navy, True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Cover Orange Dot", emu(10.05), emu(2.88), emu(0.22), emu(0.22), theme["warning"], "FFFFFF", "ellipse"))
        shape_id += 1
        if subtitle:
            shapes.append(text_box_xml(shape_id, "Cover Subtitle", emu(3.0), emu(3.9), emu(7.35), emu(0.42), [paragraph_xml(clamp_text(subtitle, 78), FONT["caption"], theme["muted"], False, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        if kicker:
            shapes.append(text_box_xml(shape_id, "Cover Kicker", emu(3.18), emu(2.09), emu(1.5), emu(0.26), [paragraph_xml(clamp_text(kicker, 18), FONT["micro"], "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        return shape_id

    if variant == "center_capsule_title":
        shapes.append(outlined_shape_xml(shape_id, "Cover Center Capsule", emu(3.15), emu(2.62), emu(7.05), emu(0.78), theme["primary"], 19050, "roundRect", theme.get("title_fill", "FFFFFF")))
        shape_id += 1
        shapes.append(line_segment_xml(shape_id, "Cover Capsule Dots", emu(3.32), emu(3.29), emu(9.98), emu(3.29), theme["primary"], 6350, "dash"))
        shape_id += 1
        shapes.append(outlined_shape_xml(shape_id, "Cover Handle", emu(6.35), emu(2.18), emu(0.78), emu(0.1), theme["primary"], 9525, "roundRect", "FFFFFF"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Cover Title", emu(3.42), emu(2.74), emu(6.45), emu(0.5), [paragraph_xml(clamp_text(title, 34), 2200, theme["primary"], True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        if subtitle:
            shapes.append(text_box_xml(shape_id, "Cover Subtitle", emu(4.05), emu(3.58), emu(5.2), emu(0.34), [paragraph_xml(clamp_text(subtitle, 70), FONT["micro"], "FFFFFF", False, "center")], anchor="ctr", margin_x=0, margin_y=0))
            shape_id += 1
        return shape_id

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
    variant = layout_variant(theme, "takeaway", "summary_panel", slide)
    bullets = slide.get("bullets", [])[:4]
    density = content_density([slide.get("title", "")] + bullets)
    if variant == "ribbon_summary":
        navy = theme.get("frame_color", theme["accent"])
        lead_h = 3.2 if density == "low" else 3.9
        shapes.append(shape_xml(shape_id, "Takeaway Navy Block", emu(0.95), emu(1.78), emu(3.15), emu(lead_h), navy, navy, "roundRect"))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Takeaway Blue Tab", emu(3.72), emu(2.2), emu(0.62), emu(0.58), theme["primary"], theme["primary"], "roundRect"))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Label", emu(1.35), emu(2.18), emu(2.35), emu(0.38), [paragraph_xml("CONTENTS", 1150, "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        title_font = fit_font(slide.get("title", "핵심 메시지"), 2.35, base=1800, minimum=FONT["body"], maximum=FONT["section"])
        shapes.append(text_box_xml(shape_id, "Takeaway Lead", emu(1.25), emu(2.78), emu(2.55), emu(1.35), [paragraph_xml(clamp_text(slide.get("title", "핵심 메시지"), compact_text_limit(2.35, title_font)), title_font, "FFFFFF", True, "center")], anchor="ctr", margin_x=0, margin_y=0))
        shape_id += 1
        y = 1.95
        for idx, bullet in enumerate(bullets[:3]):
            item_font = fit_font(bullet, 5.65, base=FONT["body_small"], minimum=FONT["caption"], maximum=FONT["body"])
            shapes.append(text_box_xml(shape_id, "Takeaway Ribbon Item", emu(4.75), emu(y), emu(6.95), emu(0.72), [], fill=theme["surface"], line=theme["line"], radius=True))
            shape_id += 1
            shapes.append(shape_xml(shape_id, "Takeaway Item Dot", emu(5.05), emu(y + 0.25), emu(0.18), emu(0.18), theme["warning"] if idx == 0 else theme["primary"], None, "ellipse"))
            shape_id += 1
            shapes.append(text_box_xml(shape_id, "Takeaway Item Text", emu(5.42), emu(y + 0.15), emu(5.85), emu(0.34), [paragraph_xml(clamp_text(bullet, compact_text_limit(5.65, item_font)), item_font, theme["text"], True)], anchor="ctr", margin_y=0))
            shape_id += 1
            y += 0.95
        return shape_id

    if variant in {"quote_band", "forecast_brief"}:
        band_h = 2.05 if density == "low" else 2.65
        shapes.append(translucent_shape_xml(shape_id, "Takeaway Band", emu(0), emu(2.05), SLIDE_W, emu(band_h), theme["surface_alt"], 60000))
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Takeaway Rule", emu(0.85), emu(2.28), emu(0.1), emu(1.95), theme["primary"], theme["primary"]))
        shape_id += 1
        shapes.append(text_box_xml(shape_id, "Takeaway Head", emu(1.2), emu(2.25), emu(4.0), emu(0.42), [paragraph_xml("핵심 메시지", 1300, theme["primary"], True)]))
        shape_id += 1
        quote = bullets[0] if bullets else slide.get("title", "")
        quote_font = fit_font(quote, 10.2, base=2300, minimum=FONT["section"], maximum=FONT["display"])
        shapes.append(text_box_xml(shape_id, "Takeaway Quote", emu(1.2), emu(2.85), emu(10.4), emu(band_h - 0.8), [paragraph_xml(clamp_text(quote, compact_text_limit(10.2, quote_font) * 2), quote_font, theme["text"], True)]))
        shape_id += 1
        if len(bullets) > 1:
            return render_bullets(shapes, {"bullets": bullets[1:]}, theme, shape_id, start_y=4.85, max_items=2, box_h=0.48, step=0.56, font_size=1200, text_limit=70)
        return shape_id

    if variant in {"full_bleed_callout", "decision_memo"}:
        title_font = fit_font(slide.get("title", "Takeaway"), 5.0, base=2500, minimum=FONT["section"], maximum=FONT["display"])
        shapes.append(text_box_xml(shape_id, "Decision Panel", emu(0.85), emu(1.9), emu(5.5), emu(3.55), [paragraph_xml("결론", FONT["body_small"], theme["primary"], True), paragraph_xml(clamp_text(slide.get("title", "Takeaway"), compact_text_limit(5.0, title_font) * 2), title_font, theme["text"], True)], fill=theme["surface_alt"], line=theme["line"], radius=True))
        shape_id += 1
        y = 2.05
        for idx, bullet in enumerate(bullets):
            item_font = fit_font(bullet, 4.85, base=FONT["body_small"], minimum=FONT["caption"], maximum=FONT["body"])
            shapes.append(text_box_xml(shape_id, "Decision Item", emu(6.65), emu(y), emu(5.3), emu(0.62), [paragraph_xml(f"{idx + 1}. {clamp_text(bullet, compact_text_limit(4.85, item_font))}", item_font, theme["text"], True)], fill=theme["surface"], line=theme["line"], radius=False))
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
            [paragraph_xml("핵심 메시지", FONT["body_small"], theme["primary"], True)]
            + [paragraph_xml("• " + clamp_text(b, compact_text_limit(10.5, list_font(bullets, 10.5, base=FONT["body_large"], minimum=FONT["body_small"]))), list_font(bullets, 10.5, base=FONT["body_large"], minimum=FONT["body_small"]), theme["text"]) for b in bullets],
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
    shape_id = add_background_motif(shapes, theme, 2, slide)
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
    <a:fontScheme name="{TYPEFACE_LATIN}">
      <a:majorFont><a:latin typeface="{TYPEFACE_LATIN}"/><a:ea typeface="{TYPEFACE_EAST_ASIAN}"/><a:cs typeface="{TYPEFACE_COMPLEX}"/></a:majorFont>
      <a:minorFont><a:latin typeface="{TYPEFACE_LATIN}"/><a:ea typeface="{TYPEFACE_EAST_ASIAN}"/><a:cs typeface="{TYPEFACE_COMPLEX}"/></a:minorFont>
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


def component_types(slide):
    return {safe_text(c.get("type")) for c in slide.get("components", []) if isinstance(c, dict)}


def infer_layout_from_evidence(slide):
    evidence = safe_text(slide.get("evidence_type", "")).lower()
    types = component_types(slide)
    if "comparison" in types:
        return "comparison"
    if "line_chart" in types or "time_series" in evidence or "forecast" in evidence:
        return "line_trend"
    if "bar_chart" in types or "category_comparison" in evidence:
        return "bar_comparison"
    if "metric_card" in types or "headline_metrics" in evidence:
        return "metric_dashboard"
    if "risk_matrix" in types or "risk" in evidence or "uncertainty" in evidence:
        return "risk_matrix"
    if "cause_effect" in types or "cause_effect" in evidence:
        return "cause_effect"
    if "timeline" in types or "dated_sequence" in evidence:
        return "timeline"
    if "process_flow" in types or "workflow" in evidence:
        return "process_flow"
    if "architecture_map" in types or "architecture" in evidence:
        return "architecture_map"
    if "callout" in types and slide.get("layout") not in {"takeaway", "title_cover"}:
        return "callout_focus"
    return slide.get("layout", "bullets")


def choose_visual_treatment(slide):
    layout = slide.get("layout", "bullets")
    story = safe_text(slide.get("_story_archetype", "")).lower()
    types = component_types(slide)
    components = slide.get("components", [])
    title_text = safe_text(slide.get("title", "")).lower()
    if layout == "metric_dashboard":
        metrics = [c for c in components if c.get("type") == "metric_card"]
        if story in {"decision_brief", "risk_monitoring"} and len(metrics) >= 2:
            return "hero_metric_panel"
        if len(metrics) >= 3:
            return "hero_metric_strip"
        return "hero_metric_panel"
    if layout == "bar_comparison":
        chart = next((c for c in components if c.get("type") == "bar_chart"), {})
        data_len = len(chart.get("data", []))
        if story in {"risk_monitoring", "decision_brief"} and data_len >= 4:
            return "ranked_bars"
        if data_len >= 5:
            return "ranked_bars"
        if any(word in title_text for word in ["예상", "전망", "실제", "상회", "before", "after"]):
            return "benchmark_bars"
        return "highlight_top_bar"
    if layout == "line_trend":
        if any(word in title_text for word in ["예상", "전망", "실제", "상회", "forecast"]):
            return "forecast_actual_line"
        return "annotated_trend"
    if layout == "cause_effect":
        return "cascade_flow"
    if layout == "risk_matrix":
        return "quadrant_focus"
    if layout == "comparison":
        return "two_column_scorecard"
    if layout == "takeaway" and "callout" in types:
        return "callout_takeaway"
    return "standard"


def preflight_slide(slide):
    inferred_layout = infer_layout_from_evidence(slide)
    if slide.get("layout") in {"bullets", "title_summary"} and inferred_layout not in {"bullets", "title_summary"}:
        slide["layout"] = inferred_layout
    elif component_types(slide) and inferred_layout != slide.get("layout") and slide.get("layout") not in {"title_cover", "takeaway"}:
        slide["layout"] = inferred_layout
    slide.setdefault("variant", "auto")
    slide.setdefault("visual_treatment", choose_visual_treatment(slide))
    return slide


def infer_story_archetype(deckspec):
    requested = safe_text(deckspec.get("story_archetype", "")).lower()
    if requested and requested != "auto":
        return requested
    slides = deckspec.get("slides", [])
    layouts = [safe_text(slide.get("layout", "")).lower() for slide in slides]
    text = " ".join(
        safe_text(value)
        for slide in slides
        for value in [
            slide.get("kicker", ""),
            slide.get("title", ""),
            slide.get("subtitle", ""),
            slide.get("evidence_type", ""),
            slide.get("layout_reason", ""),
        ]
    ).lower()
    if "architecture_map" in layouts or any(word in text for word in ["architecture", "아키텍처", "하네스", "a2a", "agent", "에이전트", "모델 라우팅"]):
        return "architecture_brief"
    if "timeline" in layouts or any(word in text for word in ["연표", "타임라인", "일정", "chronology"]):
        return "timeline_brief"
    if "cause_effect" in layouts or any(word in text for word in ["파급", "확산", "원인", "영향", "전이", "trigger"]):
        return "cause_to_effect"
    if "risk_matrix" in layouts and any(word in text for word in ["리스크", "위험", "불확실", "전망", "monitoring"]):
        return "risk_monitoring"
    if "comparison" in layouts and any(word in text for word in ["결정", "판단", "선택", "trade", "의사결정"]):
        return "decision_brief"
    if any(layout in layouts for layout in ["metric_dashboard", "bar_comparison", "line_trend"]):
        return "data_brief"
    return "explainer"


def narrative_rank(story_archetype, slide, original_index):
    layout = slide.get("layout", "bullets")
    if layout == "title_cover":
        return (-100, original_index)
    if layout == "takeaway":
        return (1000, original_index)
    role_orders = {
        "data_brief": {
            "metric_dashboard": 10,
            "line_trend": 20,
            "bar_comparison": 30,
            "comparison": 40,
            "cause_effect": 50,
            "risk_matrix": 60,
            "title_summary": 70,
            "callout_focus": 80,
            "process_flow": 90,
            "timeline": 95,
            "architecture_map": 100,
            "bullets": 110,
        },
        "cause_to_effect": {
            "cause_effect": 10,
            "metric_dashboard": 20,
            "bar_comparison": 30,
            "line_trend": 35,
            "comparison": 40,
            "risk_matrix": 50,
            "timeline": 60,
            "title_summary": 70,
            "callout_focus": 80,
            "process_flow": 90,
            "architecture_map": 100,
            "bullets": 110,
        },
        "risk_monitoring": {
            "risk_matrix": 10,
            "metric_dashboard": 20,
            "line_trend": 30,
            "bar_comparison": 40,
            "cause_effect": 50,
            "comparison": 60,
            "timeline": 70,
            "title_summary": 80,
            "callout_focus": 90,
            "process_flow": 100,
            "architecture_map": 110,
            "bullets": 120,
        },
        "decision_brief": {
            "comparison": 10,
            "metric_dashboard": 20,
            "bar_comparison": 30,
            "line_trend": 35,
            "risk_matrix": 40,
            "process_flow": 50,
            "cause_effect": 60,
            "title_summary": 70,
            "callout_focus": 80,
            "timeline": 90,
            "architecture_map": 100,
            "bullets": 110,
        },
        "explainer": {
            "title_summary": 10,
            "cause_effect": 20,
            "process_flow": 30,
            "architecture_map": 35,
            "metric_dashboard": 40,
            "bar_comparison": 50,
            "line_trend": 55,
            "comparison": 60,
            "risk_matrix": 70,
            "timeline": 80,
            "callout_focus": 90,
            "bullets": 100,
        },
        "timeline_brief": {
            "timeline": 10,
            "line_trend": 20,
            "metric_dashboard": 30,
            "cause_effect": 40,
            "bar_comparison": 50,
            "comparison": 60,
            "risk_matrix": 70,
            "title_summary": 80,
            "callout_focus": 90,
            "process_flow": 100,
            "architecture_map": 110,
            "bullets": 120,
        },
        "architecture_brief": {
            "architecture_map": 10,
            "process_flow": 20,
            "comparison": 30,
            "metric_dashboard": 40,
            "cause_effect": 50,
            "risk_matrix": 60,
            "bar_comparison": 70,
            "line_trend": 75,
            "title_summary": 80,
            "callout_focus": 90,
            "timeline": 100,
            "bullets": 110,
        },
    }
    order = role_orders.get(story_archetype, role_orders["data_brief"])
    return (order.get(layout, 500), original_index)


def apply_narrative_order(deckspec, story_archetype):
    design = deckspec.get("design", {})
    mode = safe_text(design.get("narrative_order", "auto")).lower()
    if mode in {"preserve", "manual", "off", "none"}:
        return deckspec["slides"]
    indexed = list(enumerate(deckspec["slides"], 1))
    ordered = sorted(indexed, key=lambda item: narrative_rank(story_archetype, item[1], item[0]))
    deckspec["slides"] = [slide for _, slide in ordered]
    return deckspec["slides"]


def normalize_deckspec(deckspec):
    if not isinstance(deckspec, dict):
        raise ValueError("DeckSpec must be a JSON object.")
    slides = deckspec.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("DeckSpec must include a non-empty slides array.")
    deckspec.setdefault("design", {"theme": "policy_brief"})
    story_archetype = infer_story_archetype(deckspec)
    deckspec["story_archetype"] = story_archetype
    for idx, slide in enumerate(slides, 1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {idx} must be an object.")
        slide["_story_archetype"] = story_archetype
        slide.setdefault("layout", "bullets")
        slide.setdefault("variant", "auto")
        slide.setdefault("kicker", deckspec.get("deck_title", ""))
        slide.setdefault("title", f"Slide {idx}")
        slide.setdefault("subtitle", "")
        slide.setdefault("bullets", [])
        slide.setdefault("components", [])
        slide.setdefault("speaker_note", "")
        preflight_slide(slide)
    apply_narrative_order(deckspec, story_archetype)
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
