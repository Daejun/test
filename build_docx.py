# -*- coding: utf-8 -*-
"""큐비즘 전시 견학 보고서 .docx 생성 (고찰 중심 개정판 · 사진 삽입)"""
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
EA = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia"

doc = Document()
style = doc.styles["Normal"]
style.font.name = FONT
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(EA, FONT)


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


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    set_font(p.add_run(text), size=11.5, bold=True, color=GRAY)


def para(text, size=10.5, color=None, after=6):
    """본문 한 문단."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.25
    set_font(p.add_run(text), size=size, color=color)
    return p


def lead(lead_txt, rest, size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.25
    set_font(p.add_run(lead_txt), size=size, bold=True)
    set_font(p.add_run(rest), size=size)
    return p


def bullet(text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.2
    set_font(p.add_run(text), size=size)
    return p


_fig = [0]


def figure(filename, caption, width_cm=12):
    _fig[0] += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(os.path.join(IMG, filename), width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraph_format.space_after = Pt(8)
    set_font(c.add_run(f"[사진 {_fig[0]}] {caption}"), size=9, color=LGRAY, italic=True)


def figp(filename, caption):
    figure(filename, caption, width_cm=8)


def quote(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(20)
    p.paragraph_format.right_indent = Pt(20)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    set_font(p.add_run(text), size=10.5, italic=True, color=GRAY)


# ── 제목 ──────────────────────────────────────────────
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(t.add_run("큐비즘 전시 견학 보고서"), size=22, bold=True, color=NAVY)
s = doc.add_paragraph()
s.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(s.add_run("― 〈큐비스트: 시각의 혁신가들〉을 '보는 방식의 전환'으로 읽다 ―"),
         size=12, bold=True, color=GRAY)
m = doc.add_paragraph()
m.alignment = WD_ALIGN_PARAGRAPH.CENTER
set_font(m.add_run("퐁피두센터 한화 · 2026. 6. 18. 관람    |    작성자: 김민지"), size=10.5)
doc.add_paragraph()

# ══════════════════════════════════════════════════════
# 들어가며
# ══════════════════════════════════════════════════════
h1("들어가며 ― 이 보고서가 말하려는 것")
para("이 전시를 '유명 화가들의 찌그러진 그림 모음'으로 보면 견학에서 가져갈 것이 거의 없다. 큐비즘의 "
     "핵심은 그림의 양식이 아니라 '세상을 보는 방식'을 바꾼 사건에 있기 때문이다. 르네상스 이후 약 500년간 "
     "서양 회화를 지배한 단일 시점(원근법)을 큐비즘이 깨뜨린 그 전환은, 공교롭게도 지금 우리가 하는 "
     "커뮤니케이션 업무가 통과하고 있는 변화와 정확히 닮아 있었다.")
para("그래서 이 보고서는 작품 해설이 아니라, 전시가 던진 한 가지 질문 ― ‘하나의 정리된 이미지를 일방적으로 "
     "보여주던 시대가 끝났을 때, 우리는 무엇을 어떻게 전달해야 하는가’ ― 를 중심에 두고 정리했다. 결론부터 "
     "말하면, 큐비즘은 100년 앞서 그려 보인 ‘파편화된 인식’의 시각 이론이며, 그 안에 우리 업무에 대한 세 가지 "
     "구체적인 교훈이 들어 있었다.")

# ══════════════════════════════════════════════════════
# 1. 견학 개요
# ══════════════════════════════════════════════════════
h1("1. 견학 개요")
tb = doc.add_table(rows=0, cols=2)
tb.style = "Light Grid Accent 1"
tb.alignment = WD_TABLE_ALIGNMENT.CENTER
for k, v in [
    ("전시", "퐁피두센터 한화 개관전 〈큐비스트: 시각의 혁신가들〉 (1907~1927 파리 큐비즘)"),
    ("장소·기간", "여의도 63빌딩 별관 / 2026. 6. 4 ~ 10. 4 (관람일 6. 18.)"),
    ("규모", "파리 시기 40여 작가·90여 점 + 한국 작가 11인·21점(코리아 포커스), 총 54인·112점 / 8개 주제 섹션"),
    ("운영 구조", "한화문화재단·파리 퐁피두 4년 파트너십(2023~). 한화=운영, 퐁피두=큐레이션·소장품. 연 2회 전시"),
    ("관람료", "성인 28,000원 (청소년 19,600 / 어린이 16,800 / 65세+ 19,600)"),
]:
    cells = tb.add_row().cells
    cells[0].paragraphs[0].clear(); set_font(cells[0].paragraphs[0].add_run(k), size=9.5, bold=True)
    cells[1].paragraphs[0].clear(); set_font(cells[1].paragraphs[0].add_run(v), size=9.5)
    cells[0].width = Cm(3); cells[1].width = Cm(13)
doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ══════════════════════════════════════════════════════
# 2. 무엇을 보았나
# ══════════════════════════════════════════════════════
h1("2. 무엇을 보았나 ― 전시의 흐름")
para("전시는 입구의 아폴리네르 문장에서 시작한다.")
quote("“큐비즘이 이전의 회화와 구별되는 지점은, 그것이 모방의 예술이 아니라 창조에까지 이르고자 하는 "
      "개념(conceptual)의 예술이라는 점이다.” ― 기욤 아폴리네르, 《입체파 화가들》, 1913")
figp("fig01_apollinaire.jpg", "전시 도입부 텍스트 월 ― 큐비즘은 ‘모방’이 아니라 ‘개념’의 예술이다")
para("이 한 문장이 전시 전체의 주제다. 8개 섹션은 큐비즘이 '세잔의 영향 → 형태의 해체(분석적 큐비즘) → "
     "색채로의 확장(오르피즘) → 전후의 변주'로 나아가는 과정을, 그리고 마지막에 한국 근대미술과의 만남"
     "(코리아 포커스)을 보여준다. 관람하며 특히 눈에 남은 장면은 다음과 같다.")

lead("형태의 환원 ― ", "피카소의 1907년 〈여인의 두상〉은 얼굴을 사실적으로 그리지 않고 몇 개의 면과 선으로 "
     "줄여 버린다. 그런데 그 ‘덜어냄’이 오히려 강한 인상을 남긴다.")
figp("fig02_picasso_bust.jpg", "파블로 피카소 〈여인의 두상〉, 1907 ― 묘사가 아니라 환원")

lead("해체의 극단 ― ", "분석적 큐비즘 구간의 회색 작품들은 대상을 여러 시점으로 잘게 쪼개, 무엇을 그렸는지 "
     "거의 알아보기 어려운 지점까지 밀고 간다. 아름답지만 동시에 ‘의미가 사라지기 직전’의 긴장이 느껴진다.")
figp("fig03_head_of_woman.jpg", "피카소 〈여인 두상〉 ― 다시점으로 분해된 얼굴")
figp("fig04_reclining_nude.jpg", "누워 있는 인물 ― 양감의 단순화")
figp("fig05_analytic.jpg", "거의 추상에 다다른 분석적 큐비즘 ― 식별의 한계 지점")
figp("fig06_large_cubist.jpg", "여러 인물·형태가 중첩된 대형 화면")
figure("fig07_green_work.jpg", "점묘·패턴이 결합된 종합적 큐비즘 ― 다시 ‘읽히는’ 요소의 회복", width_cm=12)

lead("색채의 폭발 ― ", "소니아 들로네의 원형 추상은 절제된 회색조와 정반대로, 강렬한 색과 리듬으로 시선을 "
     "붙든다. 관람객의 발길과 카메라가 가장 오래 머무는 구간이었다.")
figure("fig08_delaunay_single.jpg", "소니아 들로네 ― 원형 색채 추상(오르피즘)", width_cm=11)
figure("fig09_delaunay_wall.jpg", "색채 구간의 벽면 전경 ― 사실상의 ‘포토 스폿’", width_cm=13)

lead("회화를 넘어서 ― ", "야쿨로프의 〈Bull〉과 대형 무대막은 큐비즘이 캔버스를 벗어나 무대·디자인으로 "
     "번져 나갔음을 보여준다.")
figure("fig10_yakulov_bull.jpg", "조르주 야쿨로프 〈Bull〉", width_cm=13)
para("공간 자체도 하나의 작품처럼 연출됐다. 건축가 장 미셸 빌모트는 과거 아쿠아리움 자리를 반투명 유리의 "
     "‘빛의 상자’로 바꾸고, 대형 무대막을 아트리움 중앙·고천장에 걸어 압도적인 첫 장면을 만들었다.")
figure("fig11_atrium.jpg", "아트리움에 걸린 대형 무대막 ― 공간으로 만든 ‘첫인상’", width_cm=13)

# ══════════════════════════════════════════════════════
# 3. 고찰 (핵심)
# ══════════════════════════════════════════════════════
h1("3. 고찰 ― 큐비즘이 우리 업무에 던지는 질문")

h2("(1) 단일 시점의 붕괴: 우리는 더 이상 ‘하나의 공식 이미지’를 보여줄 수 없다")
para("르네상스 원근법은 ‘하나의 고정된 눈’을 전제한다. 화가는 한 자리에 서서, 관객도 한 자리에 세워두고, "
     "세계를 단일한 시점으로 정리해 보여준다. 권위적이고 일방적인 시선이다. 큐비즘은 바로 그 전제를 깼다. "
     "하나의 대상을 정면·측면·위에서 본 모습을 한 화면에 동시에 펼쳐 놓고, 그 파편을 조합해 의미를 만드는 "
     "일은 관객에게 넘긴다.")
para("이 구도는 지금의 커뮤니케이션과 정확히 겹친다. 과거의 브랜드 커뮤니케이션은 르네상스 회화처럼 작동했다 "
     "― 잘 정돈된 단일 이미지를, 정해진 자리에 앉은 청중에게 일방적으로 송출하는 방식. 그러나 지금 소비자는 "
     "광고·매장·리뷰·지인의 한마디·짧은 영상 같은 서로 다른 ‘시점의 파편’을 제각기 마주치고, 그것을 머릿속에서 "
     "스스로 조합해 하나의 브랜드 상(像)을 만든다. 큐비즘 그림 앞에서 흩어진 면들을 모아 ‘아, 얼굴이구나’ 하고 "
     "재구성하는 관객처럼. 결국 우리가 통제할 수 있는 것은 ‘완성된 한 장의 그림’이 아니라, 소비자가 조합하게 "
     "될 ‘파편들의 질과 그것들 사이의 정합성’뿐이다. 산출물을 채널별로 따로 평가하던 습관을, ‘소비자 머릿속에서 "
     "합쳐졌을 때 하나로 읽히는가’라는 기준으로 바꿔야 하는 이유가 여기에 있다.")

h2("(2) 해체에는 반드시 ‘축’이 필요하다: 전시가 정직하게 보여준 분석적 큐비즘의 한계")
para("주목할 점은, 전시가 이 해체를 끝까지 밀어붙인 ‘실패의 문턱’까지 정직하게 보여준다는 것이다. 분석적 "
     "큐비즘 구간의 회색 작품들(사진 3·5)은 형태가 거의 식별 불가능한 지점까지 부서져 있다. 미술사에서도 "
     "큐비즘은 이 시기 ‘추상의 문턱에서 의미를 잃을 위기’에 부딪혔고, 그래서 피카소와 브라크는 화면에 신문 "
     "글자나 실제 사물 조각을 붙이는 콜라주(종합적 큐비즘)로 방향을 틀어, 알아볼 수 있는 기호를 다시 끌어들였다"
     "(사진 7의 ‘읽히는 요소’가 그 회복이다).")
para("이 전개의 교훈은 흔한 조언과 정반대다. ‘채널을 쪼개고 메시지를 잘게 나누라’는 말은 절반만 맞다. 전시는 "
     "해체가 지나치면 의미 자체가 증발한다고 경고한다. 파편화가 성립하려면, 흩어진 조각을 다시 하나로 읽히게 "
     "묶어주는 ‘개념의 축’이 먼저 있어야 한다. 즉 채널을 늘릴수록 중요한 것은 채널의 수가 아니라, 모든 파편을 "
     "관통하는 단 하나의 개념이다. 캠페인을 시작할 때 채널 플랜보다 ‘모든 조각을 한 문장으로 묶는 개념’을 "
     "먼저 확정해야 하는 이유다.")

h2("(3) 더하는 기술이 아니라 ‘덜어내는’ 기술")
para("큐비즘 작품은 디테일을 더해서가 아니라, 본질만 남기고 덜어내서 강해진다. 피카소의 1907년 두상(사진 2)은 "
     "얼굴의 사실적 묘사를 포기한 대신 강렬해졌다. ‘개념의 예술’이라는 아폴리네르의 말은, 결국 ‘무엇을 그릴까’가 "
     "아니라 ‘무엇을 버릴까’의 예술이라는 뜻이다. 정보가 과잉인 시대에 메시지의 힘은 ‘얼마나 많이 담았는가’가 "
     "아니라 ‘무엇을 덜어낼 수 있었는가’에서 나온다. 산출물 리뷰 단계에 ‘무엇을 더 더할까’ 대신 ‘무엇을 더 뺄 수 "
     "있을까’라는 질문을 정식으로 넣는 것만으로도 결과물의 밀도는 달라질 것이다.")

h2("(4) 비판적으로 ― ‘코리아 포커스’는 또 하나의 시점인가, 잘 만든 각주인가")
para("한 가지는 비판적으로 보았다. 전시는 ‘코리아 포커스’로 한국 근대미술을 끌어들여 현지화를 시도했지만, 이 "
     "섹션은 메인 전시동이 아니라 갤러리2의 메자닌(중간층)에 놓여 있다. 의도는 분명히 좋다. 다만 물리적 배치는 "
     "‘본편(프랑스 큐비즘) + 부록(한국)’의 구도로 읽힐 여지를 남긴다. 큐비즘의 논리를 빌리자면, 진정한 현지화란 "
     "로컬을 또 하나의 ‘대등한 시점’으로 화면 중앙에 끌어들이는 일이지, 잘 만든 각주로 덧붙이는 일이 아니다. "
     "외부 IP·콘텐츠를 들여올 때 우리가 경계할 지점도 똑같다 ― 로컬 맥락이 장식이나 명분에 그치는 순간, 그것은 "
     "현지화가 아니라 면피가 된다. 글로벌 콘텐츠를 가져오는 일이 잦아질수록, ‘우리 이야기’를 서사의 중심에 둘 "
     "수 있는가가 진짜 실력이 된다.")

# ══════════════════════════════════════════════════════
# 4. 업무에의 적용
# ══════════════════════════════════════════════════════
h1("4. 업무에의 적용")
para("위 고찰에서 곧바로 도출되는, 우리가 실제로 바꿔볼 만한 것들이다.")
bullet("‘완성된 한 장’이 아니라 ‘조합될 파편들’을 만든다 ― 산출물을 채널별로 따로 보지 말고, 소비자 머릿속에서 "
       "합쳐졌을 때의 정합성으로 평가한다. 메시지 가이드의 기준을 ‘톤이 일치하는가’에서 ‘하나의 개념으로 "
       "환원되는가’로 옮긴다.")
bullet("채널보다 개념을 먼저 ― 캠페인 착수 시 채널 플랜에 앞서, 모든 조각을 한 문장으로 묶는 개념 축을 먼저 "
       "확정한다(파편화의 전제 조건).")
bullet("‘덜어내기’를 정식 절차로 ― 산출물 리뷰에 ‘무엇을 더 뺄 수 있는가’를 고정 질문으로 넣는다.")
bullet("현지화는 부록이 아니라 중심에 ― 외부 콘텐츠·IP를 도입할 때 로컬 요소를 서사의 중앙에 배치할 방법을 "
       "기획 단계에서부터 설계한다.")
bullet("부서 단체 견학(10/4 종료 전) ― 공간 연출과 동선은 글로 옮기기 어렵다. 직접 보는 것이 가장 빠른 학습이다.")

# ══════════════════════════════════════════════════════
# 5. 맺으며
# ══════════════════════════════════════════════════════
h1("5. 맺으며")
para("이번 견학에서 가장 값진 것은 유명 작품을 봤다는 사실이 아니라, ‘무엇을 보여주느냐’보다 ‘사람들이 어떻게 "
     "보게 만드느냐’가 더 근본적인 문제라는 점을 다시 확인한 것이다. 큐비즘은 보는 방식을 바꿈으로써 미술의 "
     "다음 100년을 열었다. 우리 일도 결국 같은 질문 위에 서 있다. 다만 한 가지는 분명하다 ― 시점을 흩뜨릴수록, "
     "그것을 하나로 묶어줄 개념은 더 단단해야 한다.")

# ── 부록 ──────────────────────────────────────────────
h1("부록. 도판 목록 및 출처")
para("현장 사진 11장은 본문에 삽입했다. (사진 1 도입 인용문 / 2 피카소 두상 / 3~6 분석적 큐비즘 / "
     "7 종합적 큐비즘 / 8~9 들로네 색채 / 10 야쿨로프 〈Bull〉 / 11 아트리움 무대막)", size=9, color=LGRAY)
para("참고 출처: 퐁피두센터 한화·파리 퐁피두 공식 / 경향신문·데일리아트 / The Korea Herald·Artsy·Dezeen 보도. "
     "※ 작품명·섹션 구성·수치는 전시 자료와 보도를 종합한 것으로, 일부는 출처 간 차이가 있어 보정했다.",
     size=9, color=LGRAY)

out = "큐비즘_전시_견학_보고서.docx"
doc.save(out)
print("saved:", out)
