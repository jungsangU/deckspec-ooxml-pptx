import argparse
import html
import json
import os
import re
import sys
import zipfile
from pathlib import Path


DEFAULT_OUTPUT = "deck_from_deckspec_ooxml.pptx"
BASE_OOXML_DIR = Path(__file__).with_name("base_ooxml")

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


def theme_for(deckspec):
    design = deckspec.get("design", {}) if isinstance(deckspec, dict) else {}
    theme_name = design.get("theme", "policy_brief")
    return THEMES.get(theme_name, THEMES["policy_brief"])


def theme_color(theme, key):
    return theme.get(key, theme["primary"])


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


def render_bullets(shapes, slide, theme, shape_id, start_y=1.95):
    y = start_y
    for bullet in slide.get("bullets", [])[:5]:
        shapes.append(
            text_box_xml(
                shape_id,
                "Bullet Background",
                emu(0.8),
                emu(y),
                emu(11.75),
                emu(0.62),
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
                "Bullet",
                emu(1.05),
                emu(y + 0.08),
                emu(11.25),
                emu(0.44),
                [paragraph_xml("• " + clamp_text(bullet, 80), 1800, theme["text"])],
            )
        )
        shape_id += 1
        y += 0.78
    return shape_id


def render_metric_dashboard(shapes, slide, theme, shape_id):
    cards = [c for c in slide.get("components", []) if c.get("type") == "metric_card"][:3]
    if not cards:
        return render_bullets(shapes, slide, theme, shape_id)

    card_w = 3.72
    x_positions = [0.8, 4.81, 8.82]
    for idx, card in enumerate(cards):
        accent = theme_color(theme, card.get("emphasis", "primary"))
        x = x_positions[idx]
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Card",
                emu(x),
                emu(2.0),
                emu(card_w),
                emu(2.35),
                [],
                fill=theme["surface"],
                line=theme["line"],
                radius=True,
            )
        )
        shape_id += 1
        shapes.append(shape_xml(shape_id, "Metric Accent", emu(x), emu(2.0), emu(0.08), emu(2.35), accent, accent))
        shape_id += 1
        shapes.append(
            text_box_xml(
                shape_id,
                "Metric Label",
                emu(x + 0.28),
                emu(2.18),
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
                emu(2.62),
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
                emu(3.3),
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
                emu(3.78),
                emu(card_w - 0.45),
                emu(0.34),
                [paragraph_xml(clamp_text(card.get("context", ""), 42), 1050, theme["muted"])],
            )
        )
        shape_id += 1

    return render_bullets(shapes, slide, theme, shape_id, start_y=4.85)


def render_bar_comparison(shapes, slide, theme, shape_id):
    charts = [c for c in slide.get("components", []) if c.get("type") == "bar_chart"]
    chart = charts[0] if charts else {"data": []}
    data = chart.get("data", [])[:5]
    values = [float(item.get("value", 0)) for item in data if isinstance(item.get("value", 0), (int, float))]
    max_value = max(values) if values else 1

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

    return render_bullets(shapes, slide, theme, shape_id, start_y=5.0)


def render_takeaway(shapes, slide, theme, shape_id):
    shapes.append(
        text_box_xml(
            shape_id,
            "Takeaway Panel",
            emu(0.85),
            emu(1.95),
            emu(11.65),
            emu(3.05),
            [paragraph_xml("핵심 메시지", 1300, theme["primary"], True)]
            + [paragraph_xml("• " + clamp_text(b, 82), 1900, theme["text"]) for b in slide.get("bullets", [])[:4]],
            fill=theme["surface_alt"],
            line=theme["line"],
            radius=True,
        )
    )
    return shape_id + 1


def render_slide_body(shapes, slide, theme, shape_id):
    layout = slide.get("layout", "bullets")
    if layout == "metric_dashboard":
        return render_metric_dashboard(shapes, slide, theme, shape_id)
    if layout == "bar_comparison":
        return render_bar_comparison(shapes, slide, theme, shape_id)
    if layout == "takeaway":
        return render_takeaway(shapes, slide, theme, shape_id)
    return render_bullets(shapes, slide, theme, shape_id)


def make_slide_xml(slide, index, theme):
    shapes = []
    shape_id = add_background_motif(shapes, theme, 2)
    shape_id = add_header(shapes, slide, theme, shape_id)
    shape_id = render_slide_body(shapes, slide, theme, shape_id)
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
    for idx, slide in enumerate(slides, 1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {idx} must be an object.")
        slide.setdefault("layout", "bullets")
        slide.setdefault("kicker", deckspec.get("deck_title", ""))
        slide.setdefault("title", f"Slide {idx}")
        slide.setdefault("bullets", [])
        slide.setdefault("components", [])
        slide.setdefault("speaker_note", "")
    deckspec.setdefault("design", {"theme": "policy_brief"})
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
            if path.is_file():
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
    if args.dump_ooxml:
        print(f"OOXML dumped to {os.path.abspath(args.dump_ooxml)}")


if __name__ == "__main__":
    main()
