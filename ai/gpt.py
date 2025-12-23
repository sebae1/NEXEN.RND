import os
import json
import openai
from io import BytesIO
from datetime import datetime
from pydantic import BaseModel

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from util import ExceptionWithMessage

class BudgetInsightsModel(BaseModel):
    title: str = ""
    executive_summary: list[str]
    insights: list[str] = []
    risks: list[str] = []
    recommendations: list[str] = []

class GPT:
    MODEL = "o4-mini"
    API_KEY = None
    _client = None
    _font_name = "Helvetica"  # 한글 폰트가 등록되면 동적으로 교체

    # ---------- 공개 메서드 ----------

    @classmethod
    def set_api_key(cls, key: str):
        """
        OpenAI API 키를 설정하고, 간단한 호출로 유효성을 검증한다.
        유효하지 않으면 ValueError 예외를 발생시킨다.
        """
        if not key or not isinstance(key, str):
            raise ValueError("유효한 문자열 형태의 API 키를 제공하세요.")

        # 최신 클라이언트 생성
        cls._client = openai.OpenAI(api_key=key)

        # 키 유효성 확인: lightweight 엔드포인트 호출
        try:
            _ = cls._client.models.list()  # 인증 실패시 예외 발생
        except Exception as e:
            # 가능하면 상태코드/메시지를 살펴서 명확한 에러를 던진다.
            msg = getattr(e, "message", None) or str(e)
            status = getattr(e, "status_code", None) or getattr(e, "status", None)
            if status == 401 or "Unauthorized" in msg or "authentication" in msg.lower():
                raise ExceptionWithMessage("API 키가 유효하지 않거나 권한이 없습니다.")

        cls.API_KEY = key
        # 한글 폰트 자동 등록 시도(없으면 기본 Helvetica 사용)
        cls._try_register_korean_font()

    @classmethod
    def make_report(cls, instruction: str, budget_of_categories: dict[str, any], budget_of_teams: dict[str, any]) -> bytes:
        """
        트리 dict 두 개를 입력받아 PDF 보고서(바이트)를 반환.
        - budget_of_categories: 카테고리 기반 트리
        - budget_of_teams: 조직/팀 기반 트리
        leaf는 (planned, actual) 튜플/리스트여야 합니다.
        """
        # 1) 입력 정규화 (리프만 추출)
        cat_items = cls._flatten_budget_tree(budget_of_categories, joiner=" > ")
        team_items = cls._flatten_budget_tree(budget_of_teams, joiner=" > ")

        # 2) 지표 계산
        cat_items = cls._postcalc_items(cat_items)
        team_items = cls._postcalc_items(team_items)
        cat_totals = cls._aggregate(cat_items)
        team_totals = cls._aggregate(team_items)

        # 3) GPT 인사이트
        ai = cls._generate_insights_v2(instruction, cat_items, cat_totals, team_items, team_totals)

        # 4) PDF 생성
        buffer = BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=A4,
            leftMargin=20 * mm, rightMargin=20 * mm,
            topMargin=18 * mm, bottomMargin=18 * mm
        )
        styles = _build_styles(cls._font_name)

        story = []

        # 제목/메타
        title = ai.get("title") or "예산 집행 보고서"
        story.append(Paragraph(title, styles["Title"]))
        story.append(Paragraph(datetime.now().strftime("%Y-%m-%d"), styles["Meta"]))
        story.append(Spacer(1, 6 * mm))

        # 핵심 KPI (카테고리/조직)
        story.append(Paragraph("핵심 지표 (카테고리 기준)", styles["Heading"]))
        story.append(_kpi_table(cat_totals, cls._font_name))
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph("핵심 지표 (조직/팀 기준)", styles["Heading"]))
        story.append(_kpi_table(team_totals, cls._font_name))
        story.append(Spacer(1, 6 * mm))

        # 모델 요약
        if ai.get("executive_summary"):
            story.append(Paragraph("요약", styles["Heading"]))
            for p in ai["executive_summary"]:
                story.append(Paragraph(f"• {p}", styles["Bullet"]))
            story.append(Spacer(1, 4 * mm))

        # 항목 테이블 (카테고리)
        story.append(Paragraph("항목별 현황 (카테고리 트리)", styles["Heading"]))
        story.append(_items_table(cat_items, cls._font_name))
        story.append(Spacer(1, 6 * mm))

        # 항목 테이블 (팀)
        story.append(Paragraph("항목별 현황 (조직/팀 트리)", styles["Heading"]))
        story.append(_items_table(team_items, cls._font_name))
        story.append(Spacer(1, 6 * mm))

        # 인사이트 / 리스크 / 권고
        def add_section(title_k: str, key: str):
            vals = ai.get(key) or []
            if not vals:
                return
            story.append(Paragraph(title_k, styles["Heading"]))
            for v in vals:
                story.append(Paragraph(f"• {v}", styles["Bullet"]))
            story.append(Spacer(1, 4 * mm))

        add_section("인사이트", "insights")
        add_section("리스크", "risks")
        add_section("권고사항", "recommendations")

        doc.build(story)
        return buffer.getvalue()

    # ---------- 내부 유틸 ----------

    @classmethod
    def _generate_insights_v2(
        cls,
        instruction: str,
        cat_items: list[dict], cat_totals: dict,
        team_items: list[dict], team_totals: dict
    ) -> dict:
        prompt = {
            "instruction": instruction,
            "requirements": {
                "style": "숏폼 문장, 과장 금지, 숫자는 단위/백분율 명확히",
                "sections": [
                    "title(짧은 제목)",
                    "executive_summary(2~5개 bullet)",
                    "insights(3~6개 bullet)",
                    "risks(0~4개 bullet)",
                    "recommendations(2~5개 bullet)"
                ],
                "notes": "카테고리/팀 합계 불일치 시 가능한 원인 가설 포함"
            },
            "category_totals": cat_totals,
            "category_items": cat_items,
            "team_totals": team_totals,
            "team_items": team_items,
        }

        try:
            # 1) 권장 경로: Responses.parse (스키마 강제 + 자동 파싱)
            if hasattr(cls._client, "responses") and hasattr(cls._client.responses, "parse"):
                resp = cls._client.responses.parse(
                    model=cls.MODEL,
                    reasoning={"effort": "medium"},
                    input=[
                        {"role": "system",
                         "content": "You are a senior financial analyst. 결과는 단일 JSON으로."},
                        {"role": "user",
                         "content": json.dumps(prompt, ensure_ascii=False)},
                    ],
                    text_format=BudgetInsightsModel,   # ← 파싱 스키마
                    max_output_tokens=4000
                )
                parsed = resp.output_parsed          # Pydantic 객체
                data = parsed.model_dump()           # dict로 변환
            else:
                # 2) 구버전 우회: beta.chat.completions.parse 사용 (Chat 전용 모델 필요)
                chat_model = "gpt-5-chat-latest"
                resp = cls._client.beta.chat.completions.parse(
                    model=chat_model,
                    temperature=0.2,
                    messages=[
                        {"role": "system",
                         "content": "You are a senior financial analyst. 결과는 단일 JSON으로."},
                        {"role": "user",
                         "content": json.dumps(prompt, ensure_ascii=False)},
                    ],
                    response_format=BudgetInsightsModel,  # Pydantic 스키마
                    max_tokens=1100,
                )
                parsed = resp.choices[0].message.parsed   # Pydantic 객체
                data = parsed.model_dump()

            # 최소 키 보정
            data.setdefault("title", "")
            for k in ["executive_summary", "insights", "risks", "recommendations"]:
                data.setdefault(k, [])
            return data

        except Exception:
            # 안전 Fallback
            return {
                "title": "예산 집행 보고서",
                "executive_summary": [
                    f'[카테고리] 계획 {int(round(cat_totals.get("planned", 0))):,} 원, '
                    f'집행 {int(round(cat_totals.get("actual", 0))):,} 원 '
                    f'(집행률 {cat_totals.get("execution_rate", 0):.1f}%, '
                    f'잔액 {int(round(cat_totals.get("remaining", 0))):,} 원).',
                    f'[조직/팀] 계획 {int(round(team_totals.get("planned", 0))):,} 원, '
                    f'집행 {int(round(team_totals.get("actual", 0))):,} 원 '
                    f'(집행률 {team_totals.get("execution_rate", 0):.1f}%, '
                    f'잔액 {int(round(team_totals.get("remaining", 0))):,} 원).',
                ],
                "insights": [],
                "risks": [],
                "recommendations": [],
            }

    @staticmethod
    def _flatten_budget_tree(tree: any, joiner: str = " / ") -> list[dict]:
        """
        임의 깊이의 트리(dict)를 순회하여 leaf (planned, actual)만 추출.
        return: [{ "path": "A > B > C", "planned": float, "actual": float, "depth": int }, ...]
        leaf 판단: (list/tuple) & 길이>=2 & 숫자 2개
        """
        items: list[dict] = []

        def is_pair(val: any) -> bool:
            if isinstance(val, (list, tuple)) and len(val) >= 2:
                try:
                    float(val[0]); float(val[1])
                    return True
                except Exception:
                    return False
            return False

        def walk(node: any, path: list[str]):
            if is_pair(node):
                planned, actual = float(node[0] or 0), float(node[1] or 0)
                items.append({
                    "path": joiner.join(path),
                    "depth": len(path),
                    "planned": planned,
                    "actual": actual,
                })
                return
            if isinstance(node, dict):
                for k, v in node.items():
                    walk(v, path + [str(k)])
            else:
                # 스칼라/알 수 없는 형식은 무시
                return

        if isinstance(tree, dict):
            for k, v in tree.items():
                walk(v, [str(k)])
        else:
            raise ValueError("트리 입력은 dict 형태여야 합니다.")

        # 경로 기준 정렬
        items.sort(key=lambda x: x["path"])
        return items

    @staticmethod
    def _postcalc_items(items: list[dict]) -> list[dict]:
        out: list[dict] = []
        for it in items:
            planned = float(it["planned"])
            actual = float(it["actual"])
            remaining = planned - actual
            rate = (actual / planned * 100) if planned > 0 else 0.0
            out.append({
                **it,
                "remaining": remaining,
                "execution_rate": rate,
            })
        return out

    @staticmethod
    def _aggregate(items: list[dict]) -> dict:
        planned = sum(x["planned"] for x in items)
        actual = sum(x["actual"] for x in items)
        remaining = planned - actual
        rate = (actual / planned * 100) if planned > 0 else 0.0
        return {
            "planned": planned,
            "actual": actual,
            "remaining": remaining,
            "execution_rate": rate,
        }

    @classmethod
    def _try_register_korean_font(cls):
        """
        시스템에 존재할 법한 한글 폰트를 자동 등록.
        실패 시 기본 Helvetica 사용.
        """
        candidates = [
            r"C:\Windows\Fonts\malgun.ttf",
            r"C:\Windows\Fonts\malgunbd.ttf",
        ]
        for path in candidates:
            try:
                if os.path.exists(path):
                    font_name = "KRBody"
                    pdfmetrics.registerFont(TTFont(font_name, path))
                    cls._font_name = font_name
                    return
            except Exception:
                continue
        # 실패 시 기본 Helvetica 유지


# ---------- PDF 빌더/헬퍼 ----------

def _kpi_table(totals: dict, font_name: str) -> Table:
    rows = [
        ["총 계획", _fmt_money(totals["planned"])],
        ["총 집행", _fmt_money(totals["actual"])],
        ["집행률", f'{totals["execution_rate"]:.1f}%'],
        ["잔액", _fmt_money(totals["remaining"])],
    ]
    t = Table(rows, colWidths=[35 * mm, 120 * mm])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _items_table(items: list[dict], font_name: str) -> Table:
    data = [["경로", "계획", "집행", "집행률", "잔액"]]
    for it in items:
        path = it["path"]
        # 깊이에 따라 살짝 들여쓰기 표기
        indent = "· " * max(0, it.get("depth", 1) - 1)
        data.append([
            indent + path.split(" > ")[-1] if indent else path,
            _fmt_money(it["planned"]),
            _fmt_money(it["actual"]),
            f'{it["execution_rate"]:.1f}%',
            _fmt_money(it["remaining"]),
        ])
    tbl = Table(
        data,
        colWidths=[70 * mm, 28 * mm, 28 * mm, 22 * mm, 27 * mm]
    )
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return tbl


def _build_styles(font_name: str):
    styles = getSampleStyleSheet()

    # 기본 폰트/본체
    styles["Normal"].fontName = font_name
    styles["Normal"].fontSize = 10
    styles["Normal"].leading = 14

    # 기존 Title, Bullet이 이미 있음: 추가하지 말고 수정만
    if "Title" in styles.byName:
        s = styles["Title"]
        s.fontName = font_name
        s.fontSize = 18
        s.leading = 22
        s.spaceAfter = 6
        s.alignment = 0  # LEFT
    if "Bullet" in styles.byName:
        s = styles["Bullet"]
        s.fontName = font_name
        s.leftIndent = 8
        s.spaceBefore = 2
        s.spaceAfter = 0
        # s.fontSize 그대로 두거나 필요 시 조정

    # Meta: 없으면 생성, 있으면 수정
    if "Meta" in styles.byName:
        s = styles["Meta"]
        s.fontName = font_name
        s.fontSize = 9
        s.textColor = colors.grey
        s.spaceAfter = 6
    else:
        styles.add(ParagraphStyle(
            name="Meta", parent=styles["Normal"],
            fontSize=9, textColor=colors.grey, spaceAfter=6
        ))

    # Heading: 없으면 생성, 있으면 수정
    if "Heading" in styles.byName:
        s = styles["Heading"]
        s.fontName = font_name
        s.fontSize = 12
        s.leading = 16
        s.spaceBefore = 6
        s.spaceAfter = 4
        s.textColor = colors.black
    else:
        styles.add(ParagraphStyle(
            name="Heading", parent=styles["Normal"],
            fontSize=12, leading=16, spaceBefore=6, spaceAfter=4,
            textColor=colors.black
        ))

    return styles


def _fmt_money(x: float) -> str:
    try:
        return f"{int(round(x)):,} 원"
    except Exception:
        return str(x)