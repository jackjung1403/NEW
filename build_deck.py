#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the 《一遍一遍》 ADD MV pitch deck (.pptx)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- palette ----
BG     = RGBColor(0x0D, 0x0E, 0x12)
PANEL  = RGBColor(0x17, 0x1A, 0x22)
PANEL2 = RGBColor(0x20, 0x24, 0x2F)
TX     = RGBColor(0xEC, 0xEE, 0xF4)
MUT    = RGBColor(0x9A, 0xA0, 0xB0)
BLUE   = RGBColor(0x5B, 0x8C, 0xFF)
RED    = RGBColor(0xFF, 0x4D, 0x5E)
WARM   = RGBColor(0xFF, 0xB2, 0x4D)
GREEN  = RGBColor(0x3D, 0xDC, 0x84)
GREY   = RGBColor(0x8A, 0x90, 0xA0)
LINE   = RGBColor(0x2A, 0x2E, 0x3A)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x10, 0x12, 0x18)

FONT = "Arial"
EMU_W, EMU_H = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


def slide(bg=BG):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, EMU_W, EMU_H)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    return s


def rect(s, x, y, w, h, fill=None, line=None, line_w=1.0, round_=False):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp


def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=4, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (text,size,color,bold) tuples."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        p.space_before = Pt(0)
        p.line_spacing = line_spacing
        for (t, sz, col, bold) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col
            r.font.bold = bold; r.font.name = FONT
    return tb


def kicker(s, text, color=BLUE):
    txt(s, Inches(0.7), Inches(0.5), Inches(11), Inches(0.4),
        [[(text.upper(), 12, color, True)]])


def title(s, text, y=0.85, size=34, color=TX):
    txt(s, Inches(0.7), Inches(y), Inches(12), Inches(1.0),
        [[(text, size, color, True)]])


def accent(s, color=BLUE, y=0.95, h=0.55):
    rect(s, Inches(0.5), Inches(y), Inches(0.06), Inches(h), fill=color)


def footer(s, n):
    txt(s, Inches(0.7), Inches(7.05), Inches(10), Inches(0.3),
        [[("一遍一遍  ·  ADD  ·  MV CONCEPT DECK", 9, MUT, False)]])
    txt(s, Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.3),
        [[(str(n), 9, MUT, False)]], align=PP_ALIGN.RIGHT)


# ============================================================ 1 · TITLE
s = slide()
# tone bar across
bx = Inches(0.0); bw = EMU_W / 3
for i, c in enumerate([RGBColor(0x2F,0x4F,0x8F), RED, WARM]):
    rect(s, Emu(int(bw*i)), Inches(0.0), Emu(int(bw)+9525), Inches(0.18), fill=c)
txt(s, Inches(0.7), Inches(2.5), Inches(12), Inches(0.5),
    [[("MUSIC VIDEO · CONCEPT & TREATMENT", 14, MUT, True)]])
txt(s, Inches(0.7), Inches(3.0), Inches(12), Inches(1.6),
    [[("一遍 一遍", 72, TX, True)]])
txt(s, Inches(0.72), Inches(4.4), Inches(12), Inches(0.6),
    [[("한 번 또 한 번   ·   ADD (와지지와)", 22, MUT, False)]])
txt(s, Inches(0.7), Inches(6.4), Inches(12), Inches(0.5),
    [[("Beijing · Ultra–low–budget · One-location performance film", 13, GREY, False)]])

# ============================================================ 2 · ARTIST
s = slide(); accent(s, BLUE); kicker(s, "The Artist"); title(s, "ADD — 6인조 보이그룹")
card = rect(s, Inches(0.7), Inches(1.9), Inches(5.4), Inches(4.6), fill=PANEL, line=LINE, round_=True)
txt(s, Inches(1.0), Inches(2.15), Inches(4.9), Inches(4.2),
    [[("소속", 12, MUT, True)],
     [("와지지와 엔터테인먼트 (Wajijiwa)", 16, TX, True)],
     [("", 6, MUT, False)],
     [("데뷔", 12, MUT, True)],
     [("2025. 09. 12  ·  〈X-Fire Camp〉 출신", 16, TX, True)],
     [("", 6, MUT, False)],
     [("전신", 12, MUT, True)],
     [("Xiaowei Music Club (小哇音乐社)", 16, TX, True)],
     [("", 6, MUT, False)],
     [("구성", 12, MUT, True)],
     [("6 MEMBERS", 22, GREEN, True)]], line_spacing=1.05)
card2 = rect(s, Inches(6.4), Inches(1.9), Inches(6.2), Inches(4.6), fill=PANEL, line=LINE, round_=True)
txt(s, Inches(6.7), Inches(2.15), Inches(5.6), Inches(0.4),
    [[("MEMBERS", 12, MUT, True)]])
members = ["张瑞希  Zhang Ruixi", "朱梓豪  Zhu Zihao", "陈俊旭  Chen Junxu",
           "王世成  Wang Shicheng", "黄一珈  Huang Yijia", "沈逸峰  Shen Yifeng"]
for i, m in enumerate(members):
    yy = 2.6 + i*0.62
    rect(s, Inches(6.7), Inches(yy), Inches(0.34), Inches(0.34), fill=PANEL2, line=BLUE, round_=True)
    txt(s, Inches(6.74), Inches(yy-0.02), Inches(0.3), Inches(0.34),
        [[(str(i+1), 13, BLUE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(7.2), Inches(yy-0.01), Inches(5.0), Inches(0.4),
        [[(m, 15, TX, False)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 2)

# ============================================================ 3 · SONG / THEME
s = slide(); accent(s, RED); kicker(s, "Song & Theme", RED); title(s, "곡이 말하는 것 — 새벽까지 깨어 글 쓰는 사람")
txt(s, Inches(0.7), Inches(1.95), Inches(11.8), Inches(0.9),
    [[("주인공은 밤에 키보드로 자기만의 ‘규칙=세계’를 쓰는 사람. ",16,TX,False),
      ("동트기 전 그가 본 텅 빈 풍경",16,WARM,True),
      ("이 곡의 정서다.",16,TX,False)]], line_spacing=1.15)
# lyric quote box
q = rect(s, Inches(0.7), Inches(3.0), Inches(11.9), Inches(1.5), fill=PANEL, line=LINE, round_=True)
txt(s, Inches(1.1), Inches(3.25), Inches(11.2), Inches(1.1),
    [[("“键盘声在回响 / 写下我的规则随便一躺”", 20, TX, True)],
     [("키보드 소리가 울려 퍼지고 / 내 규칙을 적어 내려가며 아무렇게나 드러눕지  — Verse 1", 13, MUT, False)]],
    line_spacing=1.1)
# three theme chips
chips = [("天亮之前我主宰", "동트기 전 내가 지배한다 — 새벽의 장악", BLUE),
         ("明天醒来世界是空旷一片", "내일 깨면 세상은 텅 빈 한 조각 — 공허", GREY),
         ("ADD将干扰撕碎", "그룹명 ADD를 가사에 박은 셀프 레퍼런스", RED)]
for i,(cn,ko,col) in enumerate(chips):
    xx = 0.7 + i*4.05
    rect(s, Inches(xx), Inches(4.95), Inches(3.85), Inches(1.5), fill=PANEL2, line=col, round_=True)
    txt(s, Inches(xx+0.25), Inches(5.15), Inches(3.4), Inches(1.2),
        [[(cn, 15, col, True)],[("",4,MUT,False)],[(ko, 12, TX, False)]], line_spacing=1.05)
footer(s, 3)

# ============================================================ 4 · CONCEPT
s = slide(); accent(s, WARM); kicker(s, "Concept", WARM)
title(s, "끝없이 반복되는 콘크리트 = 一遍 一遍")
txt(s, Inches(0.7), Inches(1.95), Inches(11.9), Inches(1.0),
    [[("모두가 사라진 텅 빈 도시. 똑같이 ",17,TX,False),
      ("무한 반복되는 콘크리트 빌딩",17,WARM,True),
      ("이 곧 후렴 ‘一遍 一遍’의 시각적 번역이 된다. ",17,TX,False),
      ("이미 지어진 도시라 세트비 0원.",17,GREEN,True)]], line_spacing=1.2)
pts = [("無限 반복", "똑같은 빌딩의 복붙 = 곡의 반복 구조를 그대로 영상화"),
       ("孤独 → 支配", "압도적 스케일 속 인물 = 고독한 지배자 정서"),
       ("空旷 한 조각", "텅 빈 도시 = ‘明天醒来…空旷一片’"),
       ("0원 미술", "공간이 곧 미술. 초저예산 최적")]
for i,(h,b) in enumerate(pts):
    xx = 0.7 + (i%2)*6.05; yy = 3.2 + (i//2)*1.6
    rect(s, Inches(xx), Inches(yy), Inches(5.85), Inches(1.4), fill=PANEL, line=LINE, round_=True)
    txt(s, Inches(xx+0.3), Inches(yy+0.2), Inches(5.3), Inches(1.0),
        [[(h, 17, WARM, True)],[(b, 13, TX, False)]], line_spacing=1.1)
footer(s, 4)

# ============================================================ 5 · VISUAL TONE ARC
s = slide(); accent(s, BLUE); kicker(s, "Visual Tone"); title(s, "톤 아크 — 조명만 바꿔 한 공간에서")
tones = [("① BLUE", "고독 · 글 쓰는 밤", RGBColor(0x1E,0x2E,0x52), BLUE, "Verse"),
         ("② RED", "지배 · 세상 반전", RGBColor(0x4A,0x16,0x1C), RED, "Chorus"),
         ("BLACK", "무음 · 실루엣", RGBColor(0x14,0x16,0x1C), GREY, "Leavin’"),
         ("③ WARM", "해방 · 선셋", RGBColor(0x4A,0x33,0x12), WARM, "Final")]
cw = 2.95
for i,(h,b,fill,col,tag) in enumerate(tones):
    xx = 0.7 + i*3.05
    rect(s, Inches(xx), Inches(2.1), Inches(cw), Inches(2.7), fill=fill, line=col, round_=True)
    txt(s, Inches(xx+0.25), Inches(2.35), Inches(cw-0.4), Inches(0.5),
        [[(h, 20, col, True)]])
    txt(s, Inches(xx+0.25), Inches(4.0), Inches(cw-0.4), Inches(0.7),
        [[(b, 14, TX, True)]])
    rect(s, Inches(xx+0.25), Inches(4.45), Inches(1.3), Inches(0.32), fill=None, line=col, round_=True)
    txt(s, Inches(xx+0.3), Inches(4.45), Inches(1.2), Inches(0.32),
        [[(tag, 11, col, True)]], anchor=MSO_ANCHOR.MIDDLE)
txt(s, Inches(0.7), Inches(5.2), Inches(11.9), Inches(1.2),
    [[("키프레임(AI 레퍼런스) 원본 — 실제 사진은 옌자오/톈퉁위안 로케로 대체:",13,MUT,True)],
     [("blue · red · silhouette · sunrise · concrete  (cloudfront 링크는 룩북 HTML 참조)",12,GREY,False)]],
    line_spacing=1.1)
footer(s, 5)

# ============================================================ 6 · KEY VISUAL
s = slide(); accent(s, GREY); kicker(s, "Key Visual", GREY); title(s, "핵심 비주얼 — 콘크리트 매트릭스")
big = rect(s, Inches(0.7), Inches(1.95), Inches(7.4), Inches(4.5),
           fill=RGBColor(0x3A,0x3D,0x44), line=LINE, round_=True)
# fake skyline of repeating blocks
import random
random.seed(7)
base_y = 6.45
for col_i in range(11):
    bx2 = 0.95 + col_i*0.64
    bh = 2.4 + (col_i%3)*0.5
    rect(s, Inches(bx2), Inches(base_y-bh), Inches(0.5), Inches(bh),
         fill=RGBColor(0x4A,0x4E,0x57), line=RGBColor(0x5A,0x5E,0x67), line_w=0.5)
    for wy in range(int(bh/0.28)):
        for wx in range(3):
            rect(s, Inches(bx2+0.06+wx*0.15), Inches(base_y-bh+0.15+wy*0.28),
                 Inches(0.08), Inches(0.12), fill=RGBColor(0x6A,0x6E,0x78))
# lone figure
rect(s, Inches(4.3), Inches(6.05), Inches(0.12), Inches(0.32), fill=DARK)
txt(s, Inches(0.9), Inches(6.0), Inches(7.0), Inches(0.4),
    [[("▲ 텅 빈 광장 정중앙의 인물 1명 + 무한 복붙 콘크리트 (AI 키프레임)", 11, RGBColor(0xE0,0xE0,0xE0), True)]])
# right column text
txt(s, Inches(8.4), Inches(2.1), Inches(4.3), Inches(4.3),
    [[("WHY IT WORKS", 12, GREY, True)],
     [("",4,MUT,False)],
     [("• 반복 빌딩 = ‘一遍 一遍’ 직역", 15, TX, False)],
     [("• 회색 무채색 = 차가운 무드 베이스", 15, TX, False)],
     [("• 후렴서 레드 / 선셋 웜 얹어 대비", 15, TX, False)],
     [("• 옌자오·톈퉁위안 실사 = 로케비 0원", 15, TX, False)],
     [("",8,MUT,False)],
     [("REFERENCE", 12, GREY, True)],
     [("张瑞希 외 5인 — 6인 대칭 포메이션을", 13, MUT, False)],
     [("이 반복 구도 한가운데 배치", 13, MUT, False)]], line_spacing=1.15)
footer(s, 6)

# ============================================================ 7 · LOCATION
s = slide(); accent(s, BLUE); kicker(s, "Location"); title(s, "로케이션 — 베이징, 전부 무료 공공 공간")
locs = [("燕郊 옌자오", "★ 1순위", "반복 빌딩 + 낮엔 텅 빈 거리 + 가까움. 베드타운.", GREEN),
        ("天通苑 톈퉁위안", "반복 최강", "아시아 최대 단지(70만). 지하철 접근. 콘크리트 협곡.", BLUE),
        ("回龙观 후이룽관", "대안", "옆 초대형 단지. 보조 로케.", MUT),
        ("옥상 / 고층 빈집", "원테이크 코어", "작가의 방 + 창밖 반복 도시 + 선셋, 한 공간에.", WARM)]
for i,(name,tag,desc,col) in enumerate(locs):
    yy = 1.95 + i*1.14
    rect(s, Inches(0.7), Inches(yy), Inches(11.9), Inches(1.0), fill=PANEL, line=LINE, round_=True)
    txt(s, Inches(1.0), Inches(yy+0.13), Inches(3.4), Inches(0.8),
        [[(name, 18, TX, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(4.4), Inches(yy+0.28), Inches(1.7), Inches(0.44), fill=None, line=col, round_=True)
    txt(s, Inches(4.42), Inches(yy+0.28), Inches(1.66), Inches(0.44),
        [[(tag, 12, col, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(6.3), Inches(yy+0.13), Inches(6.1), Inches(0.8),
        [[(desc, 14, MUT, False)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 7)

# ============================================================ 8 · PRODUCTION APPROACH
s = slide(); accent(s, WARM); kicker(s, "Production", WARM); title(s, "연출 — 단일 로케 · 원테이크 · 선셋")
appr = [("ONE LOCATION", "옥상/고층 빈집 1곳. 실내(작가의 방) ↔ 창밖·옥상(반복 도시) 연결.", WARM),
        ("ONE-TAKE FILM", "컷을 쪼개지 않고 짐벌로 끊김 없이. ‘반복’을 카메라 흐름으로.", BLUE),
        ("REAL SUNSET", "골든아워를 클라이맥스로. 블루(밤) → 웜(해방) 자연 전환.", RED),
        ("MINIMAL GEAR", "자연광 + 짐벌 1대 + 헤이즈. 게릴라 핸드헬드.", GREEN)]
for i,(h,b,col) in enumerate(appr):
    xx = 0.7 + (i%2)*6.05; yy = 2.0 + (i//2)*1.55
    rect(s, Inches(xx), Inches(yy), Inches(5.85), Inches(1.35), fill=PANEL, line=col, round_=True)
    txt(s, Inches(xx+0.3), Inches(yy+0.18), Inches(5.3), Inches(1.0),
        [[(h, 16, col, True)],[(b, 13, TX, False)]], line_spacing=1.1)
note = rect(s, Inches(0.7), Inches(5.35), Inches(11.9), Inches(1.05), fill=PANEL2, line=LINE, round_=True)
txt(s, Inches(1.0), Inches(5.55), Inches(11.3), Inches(0.8),
    [[("TIP  ",12,WARM,True),
      ("가사 정서는 ‘새벽’이지만 촬영은 선셋 골든아워가 압도적으로 유리 → 선셋을 찍고 새벽처럼 연출(블루→웜). 관객은 방향을 모른다.",13,TX,False)]],
    anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
footer(s, 8)

# ============================================================ 9 · STORYBOARD
s = slide(); accent(s, BLUE); kicker(s, "Storyboard"); title(s, "구간별 샷 리스트 (단일 공간 원테이크)")
rows = [("Verse 1","어두운 실내, 책상·키보드 앞 인물. 카메라 서서히 다가감","BLUE 블루·헤이즈",BLUE),
        ("Pre · Chorus","창가로 이동 → 창밖 반복 도시 드러남. 퍼포먼스 시작","RED 레드 반전",RED),
        ("‘一遍 一遍’","같은 안무 반복 / 6인 대칭 포메이션 = 무한루프 훅","RED 점멸",RED),
        ("Leavin’ (무음)","조명 다 끄고 단일 백라이트 실루엣 · 슬로모","BLACK 정적",GREY),
        ("Verse 2 · Bridge","옥상으로(문 통과=공간 연결). 도시 배경 안무","→ 골든아워 전환",WARM),
        ("Chorus 3","골든아워 정점, 빛 쏟아짐, 가장 넓은 앵글 마무리","WARM 일출",WARM)]
y0 = 2.0; rh = 0.78
# header
rect(s, Inches(0.7), Inches(y0), Inches(11.9), Inches(0.5), fill=PANEL2, line=LINE)
for cx,cw2,t in [(0.9,2.3,"SECTION"),(3.3,6.5,"ACTION"),(9.9,2.5,"TONE")]:
    txt(s, Inches(cx), Inches(y0), Inches(cw2), Inches(0.5),
        [[(t,12,MUT,True)]], anchor=MSO_ANCHOR.MIDDLE)
for i,(sec,act,tone,col) in enumerate(rows):
    yy = y0+0.5+i*rh
    rect(s, Inches(0.7), Inches(yy), Inches(11.9), Inches(rh),
         fill=PANEL if i%2==0 else BG, line=LINE, line_w=0.5)
    txt(s, Inches(0.9), Inches(yy), Inches(2.3), Inches(rh),
        [[(sec,13,TX,True)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(3.3), Inches(yy), Inches(6.4), Inches(rh),
        [[(act,12.5,RGBColor(0xCF,0xD3,0xDD),False)]], anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(9.9), Inches(yy), Inches(2.5), Inches(rh),
        [[(tone,12,col,True)]], anchor=MSO_ANCHOR.MIDDLE)
footer(s, 9)

# ============================================================ 10 · MEMBER BLOCKING
s = slide(); accent(s, GREEN); kicker(s, "6-Member Blocking", GREEN); title(s, "6인 운용 — 하이브리드 추천")
opts = [("A · 1인 리드 서사", "‘글 쓰는 나’를 센터가 맡고\n나머지 5인은 그가 본 도시 속 존재", BLUE),
        ("B · 6인 군무 중심", "6명 모두 불면자.\n반복 옥상에서 6인 대칭 포메이션", RED)]
for i,(h,b,col) in enumerate(opts):
    xx = 0.7 + i*4.05
    rect(s, Inches(xx), Inches(2.0), Inches(3.85), Inches(2.4), fill=PANEL, line=col, round_=True)
    txt(s, Inches(xx+0.3), Inches(2.25), Inches(3.3), Inches(2.0),
        [[(h, 17, col, True)],[("",4,MUT,False)]]+[[(ln,13,TX,False)] for ln in b.split("\n")],
        line_spacing=1.15)
# hybrid recommended big card
rect(s, Inches(8.9), Inches(2.0), Inches(3.7), Inches(2.4), fill=RGBColor(0x10,0x21,0x0F), line=GREEN, round_=True)
txt(s, Inches(9.2), Inches(2.2), Inches(3.2), Inches(2.1),
    [[("★ HYBRID (추천)", 16, GREEN, True)],[("",4,MUT,False)],
     [("Verse = A (개별 서사)", 13, TX, False)],
     [("Chorus = B (6인 군무)", 13, TX, False)],
     [("",4,MUT,False)],
     [("곡 구조(고독한 벌스 →", 12, MUT, False)],
     [("폭발 후렴)에 정확히 일치", 12, MUT, False)]], line_spacing=1.15)
txt(s, Inches(0.7), Inches(4.9), Inches(11.9), Inches(1.4),
    [[("포메이션 메모",13,GREEN,True)],
     [("• 반복 빌딩의 1점 투시 소실점에 센터 배치, 좌우 대칭으로 5인 전개",14,TX,False)],
     [("• ‘一遍 一遍’ 4박자마다 6인이 같은 동작 반복 → 더우인/숏폼 챌린지로 그대로 전환",14,TX,False)]],
    line_spacing=1.2)
footer(s, 10)

# ============================================================ 11 · BUDGET / NEXT
s = slide(); accent(s, WARM); kicker(s, "Budget & Next Steps", WARM); title(s, "초저예산 전략 & 다음 스텝")
left = rect(s, Inches(0.7), Inches(1.95), Inches(5.85), Inches(4.5), fill=PANEL, line=LINE, round_=True)
txt(s, Inches(1.0), Inches(2.2), Inches(5.3), Inches(4.1),
    [[("WHERE THE MONEY GOES (적게)", 13, WARM, True)],[("",4,MUT,False)],
     [("로케비", 13, MUT, True),("   0원 (공공 공간)", 14, GREEN, True)],
     [("미술", 13, MUT, True),("     공간이 곧 미술", 14, TX, False)],
     [("장비", 13, MUT, True),("     짐벌 1 + 헤이즈 + 자연광", 14, TX, False)],
     [("조명", 13, MUT, True),("     실내 소형 + 선셋 활용", 14, TX, False)],
     [("",6,MUT,False)],
     [("핵심 절약 = 단일 로케 + 원테이크로", 13, TX, False)],
     [("컷·세트·이동 최소화", 13, TX, False)]], line_spacing=1.25)
right = rect(s, Inches(6.75), Inches(1.95), Inches(5.85), Inches(4.5), fill=PANEL, line=LINE, round_=True)
steps = ["로케 헌팅 — 옌자오/톈퉁위안 옥상·고층 빈집 확인",
         "선셋 시간표 + 콜시트 작성 (골든아워 역산)",
         "6인 포메이션·파트 분배 콘티 확정",
         "실제 로케 사진으로 키프레임 재생성(크레딧 충전)",
         "원테이크 리허설 → 촬영"]
txt(s, Inches(7.05), Inches(2.2), Inches(5.3), Inches(0.4),
    [[("NEXT STEPS", 13, WARM, True)]])
for i,st in enumerate(steps):
    yy = 2.7 + i*0.72
    rect(s, Inches(7.05), Inches(yy), Inches(0.4), Inches(0.4), fill=PANEL2, line=WARM, round_=True)
    txt(s, Inches(7.06), Inches(yy), Inches(0.38), Inches(0.4),
        [[(str(i+1),13,WARM,True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    txt(s, Inches(7.6), Inches(yy-0.02), Inches(4.8), Inches(0.6),
        [[(st,13,TX,False)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
footer(s, 11)

# ============================================================ 12 · CLOSING
s = slide()
bw = EMU_W/3
for i,c in enumerate([RGBColor(0x2F,0x4F,0x8F), RED, WARM]):
    rect(s, Emu(int(bw*i)), Inches(7.32), Emu(int(bw)+9525), Inches(0.18), fill=c)
txt(s, Inches(0.7), Inches(2.7), Inches(12), Inches(1.4),
    [[("一遍 一遍", 60, TX, True)]])
txt(s, Inches(0.72), Inches(3.9), Inches(12), Inches(0.6),
    [[("끝없이 반복되는 도시에서, 단 한 번의 선셋.", 20, WARM, False)]])
txt(s, Inches(0.7), Inches(5.0), Inches(12), Inches(0.5),
    [[("ADD  ·  Beijing  ·  One-location performance film", 14, MUT, False)]])

prs.save("/home/user/NEW/yibian_ADD_MV_deck.pptx")
print("saved:", len(prs.slides._sldIdLst), "slides")
