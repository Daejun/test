# -*- coding: utf-8 -*-
"""큐비즘 전시 견학 보고서 .docx 생성 스크립트 (상세판 · 사진 삽입 포함)"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

FONT = "맑은 고딕"
IMG = "report_images"
NAVY = RGBColor(0x1F, 0x37, 0x64)
GRAY = RGBColor(0x33, 0x33, 0x33)
LGRAY = RGBColor(0x66, 0x66, 0x66)

doc = Document()

style = doc.styles["Normal"]
style.font.name = FONT
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(
    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", FONT
)
EA = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia"


def set_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(EA, FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if italic is not None:
        run.font.italic = italic


def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    set_font(p.add_run(text), size=14, bold=True, color=NAVY)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    set_font(p.add_run(text), size=11.5, bold=True, color=GRAY)
    return p


def body(text, size=10.5, bullet=False, indent=False, color=None):
    p = doc.add_paragraph(style="List Bullet" if bullet else None)
    if indent and not bullet:
        p.paragraph_format.left_indent = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    set_font(p.add_run(text), size=size, color=color)
    return p


def bullet_lead(lead, rest, size=10.5):
    """• 굵은 머리어구: 본문 형태의 불릿."""
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    set_font(p.add_run(lead), size=size, bold=True)
    set_font(p.add_run(rest), size=size)
    return p


_fignum = [0]


def figure(filename, caption, width_cm=12):
    _fignum[0] += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(os.path.join(IMG, filename), width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(8)
    set_font(c.add_run(f"[사진 {_fignum[0]}] {caption}"), size=9, color=LGRAY, italic=True)


def figp(filename, caption):
    figure(filename, caption, width_cm=8)


def add_table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, htext in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.paragraphs[0].clear()
        set_font(cell.paragraphs[0].add_run(htext), size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].paragraphs[0].clear()
            set_font(cells[i].paragraphs[0].add_run(val), size=9.5)
    if widths:
        for i, w in enumerate(widths):
            for r in t.rows:
                r.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def box(title, lines):
    """핵심 요약용 음영 박스(테이블 1칸)."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Light Shading Accent 1"
    cell = t.rows[0].cells[0]
    cell.paragraphs[0].clear()
    set_font(cell.paragraphs[0].add_run(title), size=11, bold=True, color=NAVY)
    for ln in lines:
        p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        set_font(p.add_run("· " + ln), size=10)
    doc.add_paragraph()


# ══════════════════════════════════════════════════════
# 표지/제목
# ══════════════════════════════════════════════════════
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(t.add_run("큐비즘 전시 견학 보고서"), size=22, bold=True, color=NAVY)

s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(s.add_run("― 퐁피두센터 한화 개관전 〈큐비스트: 시각의 혁신가들〉 관람 분석 ―"),
         size=12, bold=True, color=GRAY)

m = doc.add_paragraph()
m.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(m.add_run("작성자: 김민지    |    작성일: 2026. 6. 18."), size=10.5)
doc.add_paragraph()

# ══════════════════════════════════════════════════════
# 0. 핵심 요약
# ══════════════════════════════════════════════════════
h1("핵심 요약 (Executive Summary)")
body("2026년 6월 4일 여의도 63빌딩에 문을 연 '퐁피두센터 한화'는 프랑스 파리 퐁피두센터가 한화문화재단과 "
     "맺은 4년 파트너십을 기반으로 운영되는 국내 첫 상설 거점으로, 향후 4년간 매년 2회 기획전을 선보인다. "
     "그 개관전이 바로 이번에 관람한 〈큐비스트: 시각의 혁신가들〉이다. 본 전시는 단순한 명화전이 아니라, "
     "글로벌 미술관 브랜드의 국내 진출·랜드마크 재생·현지화 콘텐츠가 결합된 사례로, 전시 기획과 브랜드 "
     "운영 양 측면에서 참고할 지점이 많았다. 견학에서 도출한 핵심 시사점은 다음과 같다.")
box("관람을 통해 도출한 핵심 시사점", [
    "내러티브 큐레이션: 8개 주제 섹션이 '개념 제시 → 형태 해체 → 색채 확장 → 전후 변주'로 이어지며 "
    "관람객을 단계적으로 몰입시킨다. 메시지를 먼저 세우고 사례를 점층적으로 배치하는 구성은 우리 기획·발표물에 직접 응용 가능.",
    "고객 경험(여정) 설계: 대형 작품을 활용한 강한 첫인상, 색채 구간의 포토 스폿, QR 기반 심화 정보 등 "
    "'도입–몰입–절정–공유'의 흐름이 의도적으로 설계돼 있어 오프라인 체험 공간 기획에 시사점이 크다.",
    "글로컬(글로벌+로컬) 전략: '퐁피두'라는 글로벌 브랜드에 '코리아 포커스' 특별 섹션을 결합해 외부 IP를 "
    "현지 맥락과 연결한 점은 외부 콘텐츠·브랜드 협업 시 참고할 모범 사례.",
    "랜드마크 자산의 재해석: 노후 아쿠아리움 공간을 '빛의 상자' 미술관으로 전환해 63빌딩에 새로운 방문 "
    "동기와 화제성을 부여한 공간 리브랜딩 사례.",
])

# ══════════════════════════════════════════════════════
# 1. 견학 개요
# ══════════════════════════════════════════════════════
h1("1. 견학 개요")
add_table(
    ["구분", "내용"],
    [
        ["전시명", "퐁피두센터 한화 개관전 〈큐비스트: 시각의 혁신가들〉 (The Cubists: Inventing Modern Vision)"],
        ["장소", "퐁피두센터 한화 (서울 영등포구 여의도 63빌딩 별관, 4개 층·연면적 약 10,000㎡)"],
        ["전시 기간", "2026. 6. 4 ~ 10. 4 (개관전)"],
        ["관람일", "2026. 6. 18."],
        ["관람료", "성인 28,000원 / 청소년 19,600원 / 어린이 16,800원 / 65세 이상 19,600원"],
        ["견학 목적", "글로벌 미술관 브랜드의 국내 진출 사례를 직접 확인하고, 전시 기획·공간 연출·"
                   "관람 경험 설계와 브랜드 운영 방식을 업무 관점의 시사점 중심으로 정리"],
    ],
    widths=[2.5, 13.5],
)
body("참고로 이번 견학은 '명화 감상'이 아니라 전시를 하나의 콘텐츠·공간 운영 사례로 본다는 관점에서 진행했으며, "
     "본 보고서도 작품 해설보다 기획·연출·브랜드 측면의 관찰과 적용 가능성에 무게를 두었다.")

# ══════════════════════════════════════════════════════
# 2. 전시 분석
# ══════════════════════════════════════════════════════
h1("2. 전시 분석")

h2("2.1 운영·사업 구조 ― 글로벌 미술관 IP의 국내 거점화")
body("퐁피두센터 한화는 2023년 3월 한화문화재단과 파리 퐁피두센터가 체결한 4년 파트너십을 토대로 한다. "
     "한화가 공간의 리모델링과 운영을 맡고, 퐁피두는 큐레이션 방향과 소장품 대여를 담당하는 역할 분담 구조다. "
     "향후 4년간 매년 2회, 1,500㎡ 규모의 전시실 2개를 활용해 기획전을 이어간다는 점에서, 일회성 블록버스터가 "
     "아니라 '브랜드 거점의 지속 운영' 모델이라는 점이 핵심이다.")
bullet_lead("앵커 IP 전략 — ",
            "글로벌 인지도가 확립된 브랜드(퐁피두)를 끌어와 공간 전체의 격(格)과 방문 동기를 단숨에 끌어올리는 "
            "구조. 강력한 외부 IP가 자산의 가치와 집객을 견인하는 전형적 사례다.")
bullet_lead("지속 운영형 콘텐츠 — ",
            "연 2회 전시 교체로 재방문 동기를 설계. 단발성 이벤트가 아니라 '시즌제 콘텐츠'로 고객을 묶는 방식.")
bullet_lead("프리미엄 가격 정책 — ",
            "성인 28,000원의 객단가는 일반 기획전 대비 높은 편으로, 브랜드 가치와 작품 희소성을 가격에 반영한 "
            "포지셔닝. 가격 자체가 '특별한 경험'이라는 메시지를 전달한다.")

h2("2.2 기획 콘셉트와 내러티브 ― 8개 섹션의 점층 구성")
body("전시 입구의 기욤 아폴리네르 인용문이 전체 주제를 압축한다.")
q = doc.add_paragraph()
q.paragraph_format.left_indent = Pt(18)
q.paragraph_format.space_after = Pt(4)
set_font(q.add_run("“큐비즘이 이전의 회화와 구별되는 지점은, 그것이 모방의 예술이 아니라 창조에까지 "
                   "이르고자 하는 개념(conceptual)의 예술이라는 점이다.” (《입체파 화가들》, 1913)"),
         size=10.5, italic=True)
figp("fig01_apollinaire.jpg", "전시 도입부 텍스트 월 ― 아폴리네르의 큐비즘 정의")
body("전시는 파리 시기(1907~1927) 큐비즘을 8개 주제 섹션으로 나누어 '세잔의 영향과 초기 실험 → 분석적 "
     "큐비즘과 형태 해체 → 살롱으로 진출한 큐비즘 → 오르피즘과 빛의 리듬 → 전후의 양식적 변주'로 전개한다. "
     "관람객이 큐비즘이라는 다소 난해한 사조를 '한 개의 개념에서 출발해 점차 확장되는 이야기'로 자연스럽게 "
     "따라가도록 설계한 점이 돋보인다. 이는 복잡한 주제를 청중에게 전달할 때 메시지를 먼저 제시하고 사례를 "
     "단계적으로 심화시키는 구성 원리와 정확히 일치한다.")

h2("2.3 대표 작품과 관람 포인트")
body("관람 동선에서 인상적이었던 작품과 구간을 사진과 함께 정리한다.")

bullet_lead("큐비즘의 출발 — ",
            "피카소 〈여인의 두상〉(1907) 등 초기 실험. 인체를 단순화·해체하며 '재현에서 개념으로'의 전환을 보여준다.")
figp("fig02_picasso_bust.jpg", "파블로 피카소 〈여인의 두상(Bust of a Woman)〉, 1907")

bullet_lead("분석적 큐비즘 — ",
            "회색·갈색의 절제된 색조 속에서 대상을 여러 시점으로 쪼개 재배열. 형태가 거의 추상에 가깝게 분해된다.")
figp("fig03_head_of_woman.jpg", "피카소 〈여인 두상(Head of a Woman)〉 ― 분석적 큐비즘")
figp("fig04_reclining_nude.jpg", "누워 있는 인물 ― 양감의 단순화와 해체")
figp("fig05_analytic.jpg", "다면 분할로 추상에 근접한 분석적 큐비즘 회화")
figp("fig06_large_cubist.jpg", "다수 인물·형태가 중첩된 대형 큐비즘 회화")
figure("fig07_green_work.jpg", "점묘·패턴이 결합된 종합적 큐비즘 회화", width_cm=12)

bullet_lead("색채로의 확장(오르피즘) — ",
            "소니아 들로네의 원형 색채 추상. 절제된 분석적 큐비즘과 대비되는 강렬한 색채와 리듬으로, 전시에서 "
            "시각적 임팩트가 가장 큰 구간이자 관람객의 체류·촬영이 집중되는 사실상의 '포토 스폿'이다.")
figure("fig08_delaunay_single.jpg", "소니아 들로네 ― 원형 색채 추상(오르피즘)", width_cm=11)
figure("fig09_delaunay_wall.jpg", "들로네 작품의 전시 벽면 전경 ― 색채 구간 연출", width_cm=13)

bullet_lead("응용·확장 — ",
            "조르주 야쿨로프 〈Bull〉, 대형 무대막 등. 큐비즘이 회화를 넘어 무대·디자인으로 번져 나간 영향력을 보여준다.")
figure("fig10_yakulov_bull.jpg", "조르주 야쿨로프 〈Bull〉", width_cm=13)

h2("2.4 공간·건축 연출 ― '빛의 상자'")
body("건축가 장 미셸 빌모트가 과거 아쿠아리움이 있던 63빌딩 별관을 리모델링했다. 반투명 이중 유리 외피로 "
     "낮에는 자연광을 끌어들이고 밤에는 도시를 향해 빛나는 '빛의 상자'를 구현했으며, 외피의 곡선은 한국 전통 "
     "기와의 선을 차용했다. 황금빛 본관과 대비되는 흰색 매스(mass)가 그 자체로 강한 시각적 상징이 된다. "
     "내부에서는 대형 작품을 아트리움 중앙·고천장에 배치해 압도적 첫인상을 만들고, 흰 벽·여백·조명으로 작품에 "
     "시선을 집중시켜 관람객의 동선과 체류 시간을 의도적으로 설계한 점이 인상적이었다.")
figure("fig11_atrium.jpg", "아트리움 중앙의 대형 무대막 작품 ― '한 방' 공간 연출", width_cm=13)

h2("2.5 현지화 전략 ― 코리아 포커스")
body("갤러리2 메자닌에 마련된 특별 섹션 '코리아 포커스: 모던 아방가르드를 향한 꿈의 지도'는 김환기·유영국·"
     "박래현 등 한국 근대 작가 11인의 21점을 통해, 서구 아방가르드가 한국적 현실에서 어떻게 수용·변주됐는지 "
     "조명한다. 이로써 본 전시는 파리 시기 큐비즘 40여 작가·90여 점에 한국 섹션을 더해 총 작가 54인·작품 "
     "112점 규모가 됐다. 해외 IP를 그대로 들여오는 데 그치지 않고 한국 미술사라는 현지 맥락과 연결함으로써, "
     "국내 관객에게 '우리 이야기'라는 접점을 만들어 준 점이 전략적으로 가장 영리한 대목이다.")

# ══════════════════════════════════════════════════════
# 3. 업무 관점 심층 시사점
# ══════════════════════════════════════════════════════
h1("3. 업무 관점 심층 시사점")

h2("(A) 전시·공간 기획 및 고객 경험 설계")
bullet_lead("내러티브 큐레이션 — ",
            "'개념 선언 → 단계적 심화 → 확장'으로 이어지는 8개 섹션 구조는 어려운 주제도 관객이 따라오게 만든다. "
            "기획안·제안서·발표 자료를 '핵심 메시지 → 근거의 점층적 전개'로 구성하는 틀로 차용할 수 있다.")
bullet_lead("고객 경험 여정 설계 — ",
            "도입(아폴리네르 텍스트 월) → 몰입(분석적 큐비즘) → 절정(들로네 색채/포토 스폿) → 공유(촬영·SNS) → "
            "전환(굿즈·재방문)으로 이어지는 흐름이 의도적으로 설계돼 있다. 오프라인 체험 공간을 기획할 때 "
            "'어디서 멈추고, 어디서 찍고, 어디서 사게 할 것인가'를 동선 위에 설계하는 관점이 그대로 적용된다.")
bullet_lead("포토 스폿과 자발적 확산 — ",
            "색채 구간처럼 시각적으로 강한 한 장면을 의도적으로 배치하면 관람객의 촬영·공유(UGC)가 자연 발생한다. "
            "콘텐츠·행사 기획 시 '공유를 부르는 한 장면'을 명시적으로 설계할 필요가 있다.")
bullet_lead("정보 전달의 디지털 보완 — ",
            "캡션 옆 QR로 다국어·심화 정보를 제공해 오프라인 경험에 디지털 레이어를 얹는다. 현장 콘텐츠에 "
            "확장 정보(영상·해설·구매 링크)를 연결하는 방식으로 응용 가능.")

h2("(B) 브랜드·콘텐츠·커뮤니케이션")
bullet_lead("앵커 IP를 통한 가치 제고 — ",
            "강력한 외부 브랜드와의 제휴가 공간·콘텐츠 전체의 신뢰도와 집객을 끌어올린다. 협업 파트너 선정이 "
            "곧 포지셔닝 전략임을 보여준다.")
bullet_lead("글로컬 현지화 — ",
            "글로벌 IP에 '코리아 포커스'라는 로컬 레이어를 더해 관객과의 정서적 접점을 만든 구성. 외부 콘텐츠 "
            "도입 시 '현지 관객에게 어떤 우리 이야기로 연결할 것인가'를 반드시 함께 설계해야 함을 시사한다.")
bullet_lead("시각 언어로서의 큐비즘 — ",
            "형태 해체·재구성과 다시점 동시 표현은 '하나의 화면에 여러 메시지를 압축'하는 구성 원리로, "
            "비주얼·레이아웃·키 비주얼 개발에 응용 가능하다.")
bullet_lead("색채·기하 모티프 — ",
            "들로네식의 대담한 색 대비와 원형·기하 리듬은 강한 인상을 남기는 비주얼 아이덴티티 개발에 영감을 준다.")
bullet_lead("핵심 원칙, '재현이 아닌 개념 전달' — ",
            "보이는 그대로가 아니라 전하려는 메시지를 시각적으로 재구성한다는 큐비즘의 태도는, '무엇을 말할지'를 "
            "먼저 정하고 표현은 그 다음이라는 작업 원칙과 정확히 맞닿는다.")

h2("시사점 요약")
add_table(
    ["관찰 포인트", "전시에서의 구현", "업무 적용 방향"],
    [
        ["내러티브 구성", "8개 섹션의 점층적 전개", "기획·발표물을 '메시지→점층 근거' 구조로 설계"],
        ["고객 경험 여정", "도입–몰입–절정–공유–전환 동선", "체험 공간을 정지·촬영·구매 지점 기준으로 설계"],
        ["포토 스폿/UGC", "들로네 색채 구간의 촬영 집중", "공유를 부르는 '한 장면'을 의도적으로 배치"],
        ["앵커 IP·제휴", "퐁피두 브랜드의 집객·격 제고", "파트너 선정을 포지셔닝 전략으로 접근"],
        ["글로컬 현지화", "코리아 포커스 특별 섹션", "외부 IP에 '우리 이야기' 로컬 레이어 결합"],
        ["공간 리브랜딩", "아쿠아리움→'빛의 상자' 전환", "기존 자산의 재해석으로 방문 동기 창출"],
    ],
    widths=[3.5, 6.0, 6.5],
)

# ══════════════════════════════════════════════════════
# 4. 적용 제안
# ══════════════════════════════════════════════════════
h1("4. 적용 제안")
add_table(
    ["제안 과제", "실행 아이디어", "기대 효과 / 검토 지표"],
    [
        ["기획 구성 템플릿화",
         "'개념 선언 → 단계적 심화 → 로컬 연결'의 내러티브 틀을 표준 제안서 템플릿으로 정립",
         "메시지 전달력 향상 / 제안 채택률·이해도"],
        ["체험 동선 설계 가이드",
         "오프라인 행사·부스에 '정지–촬영–전환' 포인트를 사전 설계하는 체크리스트 도입",
         "체류시간·UGC 발생량·전환율"],
        ["비주얼 레퍼런스 보강",
         "큐비즘/오르피즘의 색채·기하 모티프를 무드보드에 추가, '메시지 우선' 원칙 명문화",
         "비주얼 일관성·차별성"],
        ["단체 견학 추진",
         "전시 종료(10/4) 전 부서 단체 관람 ― 공간 연출·동선을 직접 체험",
         "벤치마킹 내재화 / 후속 아이디어 도출"],
    ],
    widths=[3.5, 7.0, 5.5],
)

# ══════════════════════════════════════════════════════
# 5. 관찰된 한계 및 유의점
# ══════════════════════════════════════════════════════
h1("5. 관찰된 한계 및 유의점")
body("균형 잡힌 참고를 위해 한계점도 함께 기록한다.", color=LGRAY)
bullet_lead("높은 객단가 — ",
            "성인 28,000원의 가격은 진입 장벽으로 작용할 수 있어, 가격 대비 경험 가치(전시 밀도·체류시간·부대 "
            "콘텐츠)에 대한 기대 관리가 필요하다.")
bullet_lead("주제의 난도 — ",
            "큐비즘은 대중적으로 다소 어려운 사조로, 내러티브 설계에도 불구하고 사전 지식이 없으면 몰입이 "
            "떨어질 수 있다. 도슨트·디지털 해설 등 보조 장치의 중요성을 시사한다.")
bullet_lead("IP 의존도 — ",
            "브랜드 파워에 기댄 집객은 강력하지만, 콘텐츠 자체의 차별성과 현지화 깊이가 뒷받침되지 않으면 "
            "지속성에 한계가 있을 수 있다.")

# ══════════════════════════════════════════════════════
# 6. 결론
# ══════════════════════════════════════════════════════
h1("6. 결론")
body("〈큐비스트: 시각의 혁신가들〉은 100년 전 미술 사조를 다루지만, 그 운영과 연출 방식은 '글로벌 콘텐츠의 "
     "국내 안착 + 공간 리브랜딩 + 현지화'라는, 오늘날 우리 업무에 그대로 시사점을 주는 사례였다. 특히 어려운 "
     "주제를 단계적 내러티브로 풀어내고, 관람객의 경험 여정을 동선 위에 설계하며, 글로벌 브랜드에 로컬 맥락을 "
     "결합한 세 가지 지점은 향후 기획·콘텐츠 업무에 구체적으로 차용할 가치가 있다. 전시는 10월 4일까지 이어지는 "
     "만큼, 부서 차원의 단체 견학을 통해 현장 경험을 공유하고 후속 아이디어로 발전시킬 것을 제안한다.")

# ══════════════════════════════════════════════════════
# 부록
# ══════════════════════════════════════════════════════
h1("부록 1. 현장 사진 목록")
body("현장 사진(총 11장)은 본문 각 항목에 삽입했으며, 도판 목록은 다음과 같다.")
for txt in [
    "사진 1: 전시 도입부 아폴리네르 인용문 (텍스트 월)",
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
    body(txt, indent=True, size=10)

h1("부록 2. 참고 출처")
for txt in [
    "퐁피두센터 한화 — 한화문화재단 공식 (hanwhafoundation.org/business/pompidou)",
    "Centre Pompidou Hanwha — Centre Pompidou 공식 (centrepompidou.fr)",
    "여의도 63빌딩에 퐁피두센터 문 연다 — 경향신문 (khan.co.kr/article/202605191731001)",
    "퐁피두센터 한화 개관, 첫 전시 《큐비스트: 시각의 혁신가들》 — 데일리아트 (d-art.co.kr)",
    "Centre Pompidou Hanwha unveils Cubist masterpieces in Seoul — The Korea Herald",
    "Inside Centre Pompidou Hanwha — Hanwha Newsroom / Artsy / Dezeen / designboom",
]:
    body(txt, indent=True, size=9)

out = "큐비즘_전시_견학_보고서.docx"
doc.save(out)
print("saved:", out)
