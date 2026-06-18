#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MV-themed pitch deck — reads pitch_data.json, restyles the pitch-deck skill
output into a dark cinematic deck with MV-appropriate section titles + native
tone visuals (no external images needed)."""
import json
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BG    = RGBColor(0x0D,0x0E,0x12); PANEL = RGBColor(0x17,0x1A,0x22)
PANEL2= RGBColor(0x20,0x24,0x2F); TX = RGBColor(0xEC,0xEE,0xF4)
MUT   = RGBColor(0x9A,0xA0,0xB0); BLUE = RGBColor(0x5B,0x8C,0xFF)
RED   = RGBColor(0xFF,0x4D,0x5E); WARM = RGBColor(0xFF,0xB2,0x4D)
GREEN = RGBColor(0x3D,0xDC,0x84); LINE = RGBColor(0x2A,0x2E,0x3A)
DARK  = RGBColor(0x10,0x12,0x18)
FONT = "Arial"
EMU_W, EMU_H = Inches(13.333), Inches(7.5)

data = json.load(open("pitch_data.json"))
prs = Presentation(); prs.slide_width = EMU_W; prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]

def slide(bg=BG):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,EMU_W,EMU_H)
    r.fill.solid(); r.fill.fore_color.rgb=bg; r.line.fill.background()
    r.shadow.inherit=False; return s

def rect(s,x,y,w,h,fill=None,line=None,lw=1.0,rnd=False):
    shp=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rnd else MSO_SHAPE.RECTANGLE,x,y,w,h)
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb=line; shp.line.width=Pt(lw)
    shp.shadow.inherit=False; return shp

def txt(s,x,y,w,h,runs,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,sa=4,ls=1.0):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=Emu(0)
    for i,para in enumerate(runs):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; p.space_after=Pt(sa); p.space_before=Pt(0); p.line_spacing=ls
        for (t,sz,col,b) in para:
            r=p.add_run(); r.text=t; r.font.size=Pt(sz); r.font.color.rgb=col
            r.font.bold=b; r.font.name=FONT
    return tb

def tone_strip(s, y, h=Inches(0.12)):
    bw=EMU_W/3
    for i,c in enumerate([RGBColor(0x2F,0x4F,0x8F),RED,WARM]):
        rect(s,Emu(int(bw*i)),y,Emu(int(bw)+9525),h,fill=c)

def skyline(s, x, y, w, h, cols=12):
    """native repeating-concrete illustration as a keyframe stand-in."""
    rect(s,x,y,w,h,fill=RGBColor(0x33,0x36,0x3D),line=LINE,rnd=True)
    base=y+h-Inches(0.12); colw=(w-Inches(0.3))/cols
    for i in range(cols):
        bx=x+Inches(0.15)+colw*i
        bh=Inches(0.9)+Inches(0.18)*(i%4)
        rect(s,bx,base-bh,colw-Inches(0.04),bh,fill=RGBColor(0x45,0x49,0x52),
             line=RGBColor(0x55,0x59,0x62),lw=0.5)
        rows=int(bh/Inches(0.16))
        for ry in range(rows):
            for rx in range(2):
                rect(s,bx+Inches(0.03)+rx*Inches(0.07),base-bh+Inches(0.08)+ry*Inches(0.16),
                     Inches(0.04),Inches(0.07),fill=RGBColor(0x66,0x6A,0x74))
    # lone figure
    fx=x+w/2
    rect(s,Emu(int(fx)),Emu(int(base-Inches(0.22))),Inches(0.07),Inches(0.22),fill=DARK)

def footer(s,n):
    txt(s,Inches(0.7),Inches(7.04),Inches(10),Inches(0.3),
        [[("一遍一遍 · ADD · MV PITCH",9,MUT,False)]])
    txt(s,Inches(12.2),Inches(7.04),Inches(0.9),Inches(0.3),
        [[(str(n),9,MUT,False)]],align=PP_ALIGN.RIGHT)

def content_slide(n, kick, title, accent, items, visual=False):
    s=slide();
    rect(s,Inches(0.5),Inches(0.55),Inches(0.06),Inches(0.95),fill=accent)
    tone_strip(s,Inches(0.0))
    txt(s,Inches(0.7),Inches(0.55),Inches(11),Inches(0.4),[[(kick.upper(),12,accent,True)]])
    txt(s,Inches(0.7),Inches(0.92),Inches(12),Inches(0.8),[[(title,32,TX,True)]])
    pw = Inches(7.0) if visual else Inches(11.9)
    panel=rect(s,Inches(0.7),Inches(2.0),pw,Inches(4.6),fill=PANEL,line=LINE,rnd=True)
    runs=[]
    for it in items:
        runs.append([("›  ",15,accent,True),(it,15,TX,False)])
        runs.append([("",5,MUT,False)])
    txt(s,Inches(1.05),Inches(2.3),pw-Inches(0.7),Inches(4.1),runs,ls=1.1)
    if visual:
        skyline(s,Inches(8.0),Inches(2.0),Inches(4.6),Inches(3.2))
        txt(s,Inches(8.0),Inches(5.35),Inches(4.6),Inches(1.1),
            [[("KEY VISUAL",11,MUT,True)],
             [("끝없이 반복되는 콘크리트 =",12,TX,False)],
             [("‘一遍 一遍’ (실로케 사진 대체 예정)",12,MUT,False)]],ls=1.1)
    footer(s,n); return s

def twocol_slide(n, kick, title, accent, comp):
    s=slide()
    rect(s,Inches(0.5),Inches(0.55),Inches(0.06),Inches(0.95),fill=accent)
    tone_strip(s,Inches(0.0))
    txt(s,Inches(0.7),Inches(0.55),Inches(11),Inches(0.4),[[(kick.upper(),12,accent,True)]])
    txt(s,Inches(0.7),Inches(0.92),Inches(12),Inches(0.8),[[(title,32,TX,True)]])
    cols=[("OUR ADVANTAGES",comp.get("our_advantages",[]),GREEN,0.7),
          ("AVOIDING",comp.get("competitors",[]),MUT,6.75)]
    for head,its,col,xx in cols:
        rect(s,Inches(xx),Inches(2.0),Inches(5.85),Inches(4.6),fill=PANEL,line=col,rnd=True)
        txt(s,Inches(xx+0.3),Inches(2.25),Inches(5.3),Inches(0.4),[[(head,16,col,True)]])
        runs=[]
        for it in its:
            runs.append([("•  ",14,col,True),(it,14,TX,False)])
            runs.append([("",5,MUT,False)])
        txt(s,Inches(xx+0.3),Inches(2.8),Inches(5.3),Inches(3.6),runs,ls=1.12)
    footer(s,n); return s

# ---- Title ----
s=slide(); tone_strip(s,Inches(0.0)); tone_strip(s,Inches(7.32))
txt(s,Inches(0.7),Inches(2.3),Inches(12),Inches(0.5),
    [[("MUSIC VIDEO · PITCH",14,MUT,True)]])
txt(s,Inches(0.7),Inches(2.75),Inches(12),Inches(1.6),[[(data["company_name"],46,TX,True)]])
txt(s,Inches(0.72),Inches(4.0),Inches(12),Inches(0.6),[[(data.get("tagline",""),22,WARM,False)]])
txt(s,Inches(0.7),Inches(6.4),Inches(12),Inches(0.5),
    [[("Beijing · Ultra-low-budget · One-location one-take · Sunset",13,MUT,False)]])

# ---- Sections (MV-relabeled titles, rotating accents) ----
secs=[("problem","The Brief","과제 — 무엇을 풀어야 하나",RED,False),
      ("solution","The Concept","콘셉트 — 반복되는 도시",WARM,True),
      ("market","Artist & Market","ADD & 시장 맥락",BLUE,False),
      ("product","Production","연출 — 단일 로케 원테이크",BLUE,False),
      ("traction","Key Shots & Hooks","핵심 샷 & 훅",RED,False),
      ("business_model","Budget Strategy","예산 전략",GREEN,False),
      ("team","Team","ADD 6인 + 크루",BLUE,False),
      ("financials","Next Steps","다음 스텝",WARM,False)]
n=2
for key,kick,title,accent,visual in secs:
    if key in data:
        content_slide(n,kick,title,accent,data[key],visual=visual); n+=1
# competition (two-col) inserted before team for flow -> keep order: put after budget
# rebuild order properly:
# (already added team/financials; we want EDGE before team) — simplest: add edge now
if "competition" in data and isinstance(data["competition"],dict):
    twocol_slide(n,"The Edge","경쟁 우위",WARM,data["competition"]); n+=1

# ---- Closing ----
s=slide(); tone_strip(s,Inches(7.32))
txt(s,Inches(0.7),Inches(2.7),Inches(12),Inches(1.4),[[(data["company_name"].split("—")[0].strip(),52,TX,True)]])
txt(s,Inches(0.72),Inches(3.95),Inches(12),Inches(0.6),[[(data.get("tagline",""),20,WARM,False)]])
txt(s,Inches(0.7),Inches(5.0),Inches(12),Inches(0.5),
    [[("ADD · Beijing · One-location performance film",14,MUT,False)]])

out="yibian_ADD_pitch_deck_MV.pptx"
prs.save(out); print("saved",out, len(prs.slides._sldIdLst),"slides")
