# -*- coding: utf-8 -*-
"""큐비즘 전시 견학 보고서 .docx 생성 스크립트 (사진 삽입 포함)"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

FONT = "맑은 고딕"
IMG = "report_images"

doc = Document()

# 기본 폰트 설정 (한글 포함)
style = doc.styles["Normal"]
style.font.name = FONT
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", FONT
)


def set_korean_font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", FONT
    )
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def heading(text, size=14, color=RGBColor(0x1F, 0x37, 0x64), space_before=12, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    set_korean_font(r, size=size, bold=True, color=color)
    return p


_fignum = [0]


def figure(filename, caption, width_cm=12):
    """이미지 + 캡션(번호 자동)을 가운데 정렬로 삽입."""
    _fignum[0] += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(os.path.join(IMG, filename), width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(8)
    r = c.add_run(f"[사진 {_fignum[0]}] {caption}")
    set_korean_font(r, size=9, color=RGBColor(0x66, 0x66, 0x66))
    r.font.italic = True


# 세로(인물) 8cm, 가로 12cm 기준
def figure_p(filename, caption):
    figure(filename, caption, width_cm=8)


def body(text, size=10.5, bullet=False, indent=False):
    p = doc.add_paragraph(style="List Bullet" if bullet else None)
    if indent and not bullet:
        p.paragraph_format.left_indent = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_korean_font(r, size=size)
    return p


# ── 제목 ──────────────────────────────────────────────
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("큐비즘 전시 견학 보고서")
set_korean_font(r, size=20, bold=True, color=RGBColor(0x1F, 0x37, 0x64))

s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = s.add_run("― 퐁피두센터 한화 개관전 〈큐비스트: 시각의 혁신가들〉 관람 ―")
set_korean_font(r, size=12, bold=True, color=RGBColor(0x55, 0x55, 0x55))

m = doc.add_paragraph()
m.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = m.add_run("작성자: 김민지    |    작성일: 2026. 6. 18.")
set_korean_font(r, size=10.5)

doc.add_paragraph()

# ── 1. 견학 개요 ──────────────────────────────────────
heading("1. 견학 개요")
for txt in [
    "전시명: 퐁피두센터 한화 개관전 〈큐비스트: 시각의 혁신가들〉",
    "장소: 퐁피두센터 한화 (서울 영등포구 여의도 63빌딩 별관)",
    "관람일: 2026년 6월 18일",
    "전시 기간: 2026. 6. 4 ~ 10. 4",
    "견학 목적: 글로벌 미술관 브랜드의 국내 진출 사례를 직접 살펴보고, 관람객을 끌어들이는 "
    "전시 기획·공간 연출과 시각 표현 방식을 우리 업무에 참고할 만한 시사점 중심으로 정리하기 위함.",
]:
    body(txt, bullet=True)
body("한 줄 요약: 프랑스 파리 퐁피두센터가 여의도 63빌딩에 문을 연 첫 전시로, 1907~1927년 큐비즘 "
     "20년의 흐름을 9개 섹션·작가 54인·작품 112점으로 조망하며, 한국 근대미술과의 연결까지 엮어낸 "
     "대형 기획전이다. 글로벌 브랜드의 국내 안착과 관람 경험 설계라는 측면에서 참고할 지점이 많았다.")

# ── 2. 전시 핵심 내용 요약 ────────────────────────────
heading("2. 전시 핵심 내용 요약")
heading("(1) 전시를 관통하는 메시지", size=11.5, color=RGBColor(0x33, 0x33, 0x33), space_before=6)
body("전시 입구의 기욤 아폴리네르 인용문이 전체 주제를 압축한다.")
q = doc.add_paragraph()
q.paragraph_format.left_indent = Pt(18)
q.paragraph_format.space_after = Pt(4)
rq = q.add_run("“큐비즘이 이전의 회화와 구별되는 지점은, 그것이 모방의 예술이 아니라 창조에까지 "
               "이르고자 하는 개념(conceptual)의 예술이라는 점이다.” (《입체파 화가들》, 1913)")
set_korean_font(rq, size=10.5)
rq.font.italic = True
figure_p("fig01_apollinaire.jpg", "전시 도입부 텍스트 월 ― 아폴리네르의 큐비즘 정의")
body("즉 큐비즘은 '보이는 대로 재현'하는 미술에서 '대상을 분해해 하나의 메시지로 재구성'하는 미술로의 "
     "전환을 선언한 사조다. 이 '재현에서 개념으로'라는 관점이 보고서 전체의 키워드다.")

heading("(2) 전시 구성과 대표 작품 (관람 동선 기준)", size=11.5, color=RGBColor(0x33, 0x33, 0x33), space_before=6)

body("큐비즘의 출발 ― 피카소 〈여인의 두상〉(1907) 등: 인체를 단순화·해체하기 시작한 초기 실험.",
     bullet=True)
figure_p("fig02_picasso_bust.jpg", "파블로 피카소 〈여인의 두상(Bust of a Woman)〉, 1907")

body("분석적 큐비즘 ― 회색·갈색 톤의 다면 분할 회화: 대상을 여러 시점에서 잘게 쪼개 화면에 재배열. "
     "형태가 거의 추상에 가깝게 분해됨.", bullet=True)
figure_p("fig03_head_of_woman.jpg", "피카소 〈여인 두상(Head of a Woman)〉 ― 분석적 큐비즘")
figure_p("fig04_reclining_nude.jpg", "누워 있는 인물 ― 형태의 단순화·해체")
figure("fig05_analytic.jpg", "다면 분할로 거의 추상에 이른 분석적 큐비즘 회화", width_cm=8)
figure_p("fig06_large_cubist.jpg", "여러 인물·형태가 중첩된 대형 큐비즘 회화")
figure("fig07_green_work.jpg", "점묘·패턴이 결합된 종합적 큐비즘 회화", width_cm=12)

body("색채로의 확장(오르피즘) ― 소니아 들로네의 원형 색채 추상: 큐비즘의 형태 실험에 강렬한 색채와 "
     "리듬을 결합. 전시에서 가장 시각적으로 강한 구간이자, 관람객의 발길이 가장 오래 머무는 '포토 스폿'.",
     bullet=True)
figure("fig08_delaunay_single.jpg", "소니아 들로네 ― 원형 색채 추상(오르피즘)", width_cm=11)
figure("fig09_delaunay_wall.jpg", "들로네 작품의 전시 벽면 전경 ― 색채 구간 연출", width_cm=13)

body("큐비즘의 응용·확장 ― 조르주 야쿨로프 〈Bull〉, 대형 무대막 등: 회화를 넘어 무대·디자인으로 번진 "
     "큐비즘의 영향력.", bullet=True)
figure("fig10_yakulov_bull.jpg", "조르주 야쿨로프 〈Bull〉", width_cm=13)

body("코리아 포커스: 모던 아방가르드를 향한 꿈의 지도 ― 김환기·유영국·박래현 등 한국 작가 11인, "
    "작품 21점. 서구 아방가르드가 한국적 현실에서 어떻게 수용·변주됐는지 조명한 특별 섹션.", bullet=True)

heading("(3) 공간·건축", size=11.5, color=RGBColor(0x33, 0x33, 0x33), space_before=6)
body("건축가 장 미셸 빌모트가 63빌딩 별관을 리모델링했다. 황금빛 본관과 대비되는 흰색 '빛의 상자' "
     "콘셉트로, 대형 작품을 아트리움 중앙에 배치해 압도적 첫인상을 연출한다.")
figure("fig11_atrium.jpg", "아트리움 중앙의 대형 무대막 작품 ― '한 방' 공간 연출", width_cm=13)

# ── 3. 시사점 ─────────────────────────────────────────
heading("3. 관람을 통해 얻은 시사점 (핵심)")
heading("(A) 기획 · 공간 연출 측면", size=11.5, color=RGBColor(0x33, 0x33, 0x33), space_before=6)
for txt in [
    "내러티브형 구성: '개념 선언 → 형태 해체(분석) → 색채 확장 → 응용 → 현지(한국) 연결'로 이어지는 "
    "9개 섹션 구조가 관람객을 자연스럽게 몰입시킨다. 콘텐츠를 '메시지 먼저, 사례는 단계적 심화'로 "
    "풀어내는 구성은 우리 기획·발표물에도 그대로 옮겨올 수 있다.",
    "랜드마크 재생: 노후한 63빌딩 별관을 글로벌 미술관으로 탈바꿈시킨 사례. 기존 자산을 새로운 "
    "콘셉트로 재해석해 화제성과 방문 동기를 만들어낸 점이 인상적이다.",
    "첫인상 동선 설계: 대형 작품을 중앙·고천장 공간에 배치한 '한 방' 연출, 흰 벽·여백·조명으로 "
    "작품을 부각하는 방식은 관람객의 시선과 체류 시간을 의도적으로 설계한 결과다. 전시·쇼룸·행사 부스 "
    "등 오프라인 체험 공간 연출에 직접 참고할 만하다.",
    "정보 전달 방식: 작품 캡션과 QR 코드를 통해 다국어·심화 정보를 제공 ― 오프라인 현장에 디지털 "
    "콘텐츠를 자연스럽게 얹어 관람 경험을 확장하는 방식이 돋보였다.",
]:
    body(txt, bullet=True)

heading("(B) 브랜드 · 비주얼 측면", size=11.5, color=RGBColor(0x33, 0x33, 0x33), space_before=6)
for txt in [
    "글로벌 브랜드의 현지화: '퐁피두'라는 강력한 글로벌 브랜드를 들여오되, '코리아 포커스'로 한국 "
    "미술사를 결합한 글로컬(글로벌+로컬) 접근. 외부 브랜드·콘텐츠와 협업할 때 현지 정서와 어떻게 "
    "연결할지에 대한 좋은 본보기다.",
    "시각 언어로서의 큐비즘: 형태 해체·재구성, 다시점 동시 표현은 '하나의 화면에 여러 메시지를 압축'"
    "하는 구성 원리로, 비주얼·레이아웃 작업에 응용 가능하다.",
    "색채·기하 모티프: 소니아 들로네식의 대담한 색 대비와 원형·기하 리듬은 강한 인상을 남기는 "
    "비주얼 아이덴티티와 그래픽 모티프 개발에 영감을 준다. 실제로 이 구간이 관람객 반응(촬영·체류)이 "
    "가장 컸다.",
    "핵심 원칙 ― “재현이 아닌 개념 전달”: 보이는 그대로가 아니라 전하려는 메시지를 시각적으로 "
    "재구성한다는 큐비즘의 태도는, '무엇을 말할지'를 먼저 정하고 표현은 그 다음이라는 우리 작업 원칙과 "
    "정확히 맞닿는다.",
]:
    body(txt, bullet=True)

# ── 4. 결론 및 적용 제안 ──────────────────────────────
heading("4. 결론 및 적용 제안")
body("큐비즘은 100년 전 미술이지만, 이번 전시는 '글로벌 콘텐츠의 국내 안착 + 공간 리브랜딩 + 현지화'"
     "라는, 오늘날 우리 업무에 그대로 시사점을 주는 사례였다. 참고할 만한 적용 지점은 다음과 같다.")
for txt in [
    "구성 틀 차용: 차기 기획·발표·캠페인에서 “개념 선언 → 단계적 심화 → 로컬 연결”의 "
    "내러티브 틀을 템플릿으로 시험 적용.",
    "비주얼 레퍼런스 보강: 큐비즘/오르피즘의 색채 대비·기하 모티프를 참고 무드보드에 추가하고, "
    "'메시지 우선' 원칙을 작업 체크리스트에 반영.",
    "단체 견학 제안: 전시 기간이 10월 4일까지이므로 부서 단체 관람을 검토 ― 공간 연출·관람 동선을 "
    "직접 경험하는 것이 자료보다 효과적이다.",
]:
    body(txt, bullet=True)

# ── 5. 부록 ───────────────────────────────────────────
heading("5. 부록 ― 현장 사진 목록")
body("현장 사진(총 11장)은 본문 각 항목에 삽입했으며, 도판 목록은 다음과 같다.")
for txt in [
    "사진 1: 전시 도입부 아폴리네르 인용문",
    "사진 2: 피카소 〈여인의 두상〉(1907)",
    "사진 3: 피카소 〈여인 두상(Head of a Woman)〉",
    "사진 4: 누워 있는 인물 (분석적 큐비즘)",
    "사진 5: 다면 분할 분석적 큐비즘 회화",
    "사진 6: 인물·형태가 중첩된 대형 큐비즘 회화",
    "사진 7: 점묘·패턴이 결합된 종합적 큐비즘 회화",
    "사진 8: 소니아 들로네 원형 색채 추상 (단독)",
    "사진 9: 들로네 작품 전시 벽면 전경",
    "사진 10: 조르주 야쿨로프 〈Bull〉",
    "사진 11: 아트리움 대형 무대막 (공간 연출 전경)",
]:
    body(txt, indent=True)

# ── 참고 출처 ─────────────────────────────────────────
heading("참고 출처", size=11.5, color=RGBColor(0x33, 0x33, 0x33))
for txt in [
    "여의도 63빌딩에 퐁피두센터 문 연다 — 경향신문 (khan.co.kr/article/202605191731001)",
    "63빌딩에 온 퐁피두, '큐비즘'으로 눈도장 — 경향신문 (khan.co.kr/article/202605192038005)",
    "퐁피두센터 한화 — 한화문화재단 공식 (hanwhafoundation.org/business/pompidou)",
    "퐁피두센터 한화, 4일 여의도 63빌딩에 문 연다 — 코리아넷 (korean-culture.org)",
]:
    p = body(txt, size=9, bullet=True)

out = "큐비즘_전시_견학_보고서.docx"
doc.save(out)
print("saved:", out)
