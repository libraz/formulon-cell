"""Generate mac-layout.css, the Mac Format Cells dialog's frames.

Every frame is a window-relative rectangle (x, y, w, h) taken from the Mac
spreadsheet app's accessibility tree (font tab: measured from its capture).
Run `python3 scripts/gen-format-dialog-mac-layout.py [tables.json]`; the optional
argument also writes the frame tables for a position checker.
"""
import json
BX,BY=13,38   # body origin in window coordinates
# frames are window-relative (x,y,w,h); selector is relative to the tab panel unless 'abs' prefix (panel-relative, outside body)
T={}
def tab(name, items): T[name]=items

# ---- chrome (panel-relative) ----
CHROME=[
 ('.fc-fmtdlg__title',(18,8,89,16)),
 ('.fc-fmtdlg__tab[data-fc-tab="number"]',(125,43,71,24)),
 ('.fc-fmtdlg__tab[data-fc-tab="align"]',(196,43,47,24)),
 ('.fc-fmtdlg__tab[data-fc-tab="font"]',(243,43,71,24)),
 ('.fc-fmtdlg__tab[data-fc-tab="border"]',(314,43,47,24)),
 ('.fc-fmtdlg__tab[data-fc-tab="fill"]',(361,43,83,24)),
 ('.fc-fmtdlg__tab[data-fc-tab="protection"]',(444,43,47,24)),
 ('.fc-fmtdlg__footer .fc-fmtdlg__btn:not(.fc-fmtdlg__btn--primary)',(414,558,86,26)),
 ('.fc-fmtdlg__footer .fc-fmtdlg__btn--primary',(510,558,86,26)),
]

tab('align',[
 ('.fc-fmtdlg__align-select-row--h > span',(39,76,216,18)),
 ('.fc-fmtdlg__align-select-row--h .fc-select',(39,99,216,26)),
 ('.fc-fmtdlg__align-select-row--v > span',(39,132,216,18)),
 ('.fc-fmtdlg__align-select-row--v .fc-select',(40,156,216,26)),
 ('.fc-fmtdlg__justify-distributed-row .fc-fmtdlg__check',(40,200,156,12)),
 ('.fc-fmtdlg__indent-row > span',(272,76,66,18)),
 ('.fc-fmtdlg__indent-row input',(273,100,49,24)),
 ('.fc-fmtdlg__align-preview-title',(390,76,33,18)),
 ('.fc-fmtdlg__align-preview-vertical',(408,102,33,153)),
 ('.fc-fmtdlg__align-preview-dial',(448,102,95,153)),
 ('.fc-fmtdlg__align-degree > span',(408,259,18,19)),
 ('.fc-fmtdlg__align-degree input',(428,257,38,24)),
 ('.fc-fmtdlg__text-control-title',(40,362,69,18)),
 ('.fc-fmtdlg__text-control .fc-fmtdlg__check:nth-of-type(1)',(54,390,168,18)),
 ('.fc-fmtdlg__text-control .fc-fmtdlg__check:nth-of-type(2)',(54,412,168,18)),
 ('.fc-fmtdlg__text-control .fc-fmtdlg__check:nth-of-type(3)',(54,434,168,18)),
])
tab('protection',[
 ('.fc-fmtdlg__protection-hint',(38,90,524,34)),
 ('.fc-fmtdlg__protection-checks .fc-fmtdlg__check:nth-of-type(1)',(40,137,60,18)),
 ('.fc-fmtdlg__protection-checks .fc-fmtdlg__check:nth-of-type(2)',(40,159,84,18)),
])
G=(38,98)  # gallery origin
def gal(style,x,y,h): return ('.fc-fmtdlg__line-style--'+style,(x,y,110,h))
tab('border',[
 ('.fc-fmtdlg__border-style-row > span',(40,75,81,18)),
 ('.fc-fmtdlg__line-gallery',(38,98,224,197)),
 gal('none',41,102,28),gal('hair',41,128,29),gal('dotted',41,155,29),gal('dashDotDot',41,182,29),gal('dashDot',41,209,29),gal('dashed',41,236,29),gal('thin',41,263,29),
 gal('mediumDashDotDot',148,102,28),gal('slantDashDot',148,128,29),gal('mediumDashDot',148,155,29),gal('mediumDashed',148,182,29),gal('medium',148,209,29),gal('thick',148,236,29),gal('double',148,263,29),
 ('.fc-fmtdlg__border-color-row > span:first-child',(38,318,42,18)),
 ('.fc-fmtdlg__border-color-row .fc-fmtdlg__color-toggle',(38,341,194,26)),
 ('.fc-fmtdlg__border-heading:not(.fc-fmtdlg__border-title)',(314,75,69,18)),
 ('.fc-fmtdlg__border-preset--none',(323,100,56,40)),
 ('.fc-fmtdlg__border-preset--outline',(402,100,57,40)),
 ('.fc-fmtdlg__border-preset--inside',(482,98,57,40)),
 ('.fc-fmtdlg__border-preset-caption--none',(316,146,70,18)),
 ('.fc-fmtdlg__border-preset-caption--outline',(395,146,70,18)),
 ('.fc-fmtdlg__border-preset-caption--inside',(475,142,70,18)),
 ('.fc-fmtdlg__border-title',(314,196,33,18)),
 ('.fc-fmtdlg__border-preview',(358,224,132,132)),
 ('.fc-fmtdlg__border-hit--top',(321,224,27,27)),
 ('.fc-fmtdlg__border-hit--inside-h',(321,276,27,27)),
 ('.fc-fmtdlg__border-hit--bottom',(321,329,27,27)),
 ('.fc-fmtdlg__border-hit--diagonalUp.fc-fmtdlg__border-hit--left-diag',(321,370,27,27)),
 ('.fc-fmtdlg__border-hit--left',(360,370,28,27)),
 ('.fc-fmtdlg__border-hit--inside-v',(410,370,28,27)),
 ('.fc-fmtdlg__border-hit--right',(462,370,28,27)),
 ('.fc-fmtdlg__border-hit--diagonalDown.fc-fmtdlg__border-hit--right-diag',(504,370,28,27)),
 ('.fc-fmtdlg__border-hint',(40,413,524,92)),
])
tab('fill',[
 ('.fc-fmtdlg__fill-bg-row > span:first-child',(46,92,42,18)),
 ('.fc-fmtdlg__fill-bg-row .fc-fmtdlg__fill-popup',(46,118,213,23)),
 ('.fc-fmtdlg__fill-pattern-color-row > span:first-child',(333,92,78,18)),
 ('.fc-fmtdlg__fill-pattern-color-row .fc-fmtdlg__fill-popup',(333,118,213,26)),
 ('.fc-fmtdlg__fill-pattern-row > span:first-child',(333,154,105,18)),
 ('.fc-fmtdlg__fill-pattern-row .fc-fmtdlg__fill-popup',(332,180,214,26)),
 ('.fc-fmtdlg__fill-sample-label',(46,351,54,18)),
 ('.fc-fmtdlg__fill-sample-box',(51,379,508,82)),
])

tab('font',[
 ('.fc-fmtdlg__font-family-row > span',(41,77,100,18)),
 ('.fc-fmtdlg__font-family-row input',(41,99,261,26)),
 ('.fc-fmtdlg__font-list--family',(41,133,261,109)),
 ('.fc-fmtdlg__font-style-label > span',(321,77,100,18)),
 ('.fc-fmtdlg__font-style-field',(321,101,128,24)),
 ('.fc-fmtdlg__font-list--style',(321,133,128,109)),
 ('.fc-fmtdlg__font-size-row > span',(468,77,100,18)),
 ('.fc-fmtdlg__font-size-row input',(468,101,100,24)),
 ('.fc-fmtdlg__font-list--size',(468,133,100,109)),
 ('.fc-fmtdlg__font-underline--mac > span',(41,258,100,18)),
 ('.fc-fmtdlg__font-underline--mac .fc-select',(41,274,261,24)),
 ('.fc-fmtdlg__font-color-row > span',(321,258,100,18)),
 ('.fc-fmtdlg__font-color-row .fc-fmtdlg__color-control',(321,274,128,24)),
 ('.fc-fmtdlg__normal-font',(468,277,104,18)),
 ('.fc-fmtdlg__font-deco-title',(44,325,80,18)),
 ('.fc-fmtdlg__font-deco .fc-fmtdlg__check:nth-of-type(1)',(53,353,110,18)),
 ('.fc-fmtdlg__font-deco .fc-fmtdlg__check:nth-of-type(2)',(53,375,110,18)),
 ('.fc-fmtdlg__font-deco .fc-fmtdlg__check:nth-of-type(3)',(53,397,110,18)),
 ('.fc-fmtdlg__font-preview-label',(325,325,60,18)),
 ('.fc-fmtdlg__font-preview-box',(328,354,216,66)),
])

CATS=[(41,97+19*i,125,19) for i in range(12)]
ncommon=[('.fc-fmtdlg__cat-label',(38,73,30,18)),
 ('.fc-fmtdlg__cat',(40,96,127,330)),
 ('.fc-fmtdlg__preview-label',(183,73,50,18)),
 ('.fc-fmtdlg__preview-cell',(185,80,386,52))]
# children of containers are relative to the container: x-cx, y-cy
def rel(f,c): return (f[0]-c[0],f[1]-c[1],f[2],f[3])
catitems=[('.fc-fmtdlg__cat-item:nth-child(%d)'%(i+1),CATS[i]) for i in range(12)]
dec=[('.fc-fmtdlg__decimals-row > span',(183,149,102,19)),('.fc-fmtdlg__decimals-row input',(290,146,34,24))]
thou=[('.fc-fmtdlg__number-check',(184,173,160,18))]
sym=[('.fc-fmtdlg__symbol-row > span',(183,179,62,21)),('.fc-fmtdlg__symbol-row .fc-select',(250,177,322,25))]
NEGC=(185,237,386,187)
neg=[('.fc-fmtdlg__negative-label',(183,213,102,18)),('.fc-fmtdlg__negative-list',NEGC)]
negitems=lambda n:[('.fc-fmtdlg__negative-list > .fc-fmtdlg__negative-item:nth-child(%d)'%(i+1),(186,238+19*i,384,19)) for i in range(n)]
PLC=(185,174,386,148)
pat=[('.fc-fmtdlg__pattern-list-label',(183,149,30,19)),('.fc-fmtdlg__pattern-list',PLC)]
patitems=lambda n:[('.fc-fmtdlg__pattern-list > .fc-fmtdlg__pattern-item:nth-child(%d)'%(i+1),(186,175+19*i,384,19)) for i in range(n)]
loc=[('.fc-fmtdlg__locale-row > span',(183,328,67,18)),('.fc-fmtdlg__locale-row .fc-select',(184,348,388,26))]
cal=[('.fc-fmtdlg__calendar-row > span',(183,380,102,18)),('.fc-fmtdlg__calendar-row .fc-select',(184,400,388,26))]
cust=[('.fc-fmtdlg__pattern-row > span',(183,149,30,19)),('.fc-fmtdlg__pattern-row input',(184,174,388,26))]
HINT_DEFAULT=(29,522,558,18)
NUMCATS={
 'general':[], 'fixed':dec+thou+neg+negitems(7), 'currency':dec+sym+neg+negitems(5), 'accounting':dec+sym,
 'date':pat+patitems(19)+loc+cal, 'time':pat+patitems(9)+loc, 'percent':dec, 'fraction':pat+patitems(9),
 'scientific':dec, 'text':[], 'special':pat+patitems(7)+loc, 'custom':cust}
ALLNUM=ncommon+catitems+dec+thou+sym+neg+negitems(7)+pat+patitems(19)+loc+cal+cust
# per-category tables for the checker (window frames, selector scope = number tab panel)
for c,items in NUMCATS.items():
    T['number:'+c]=ncommon+catitems+items
json_num=True
import os, sys
HERE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(HERE,'..','packages','formulon-cell','src','styles','core','app','format-dialog','mac-layout.css')
if len(sys.argv)>1: json.dump({'chrome':CHROME,'tabs':T},open(sys.argv[1],'w'),ensure_ascii=False)


STRUCT = """
  .fc-fmtdlg--mac .fc-fmtdlg__panel {
    position: relative;
    display: block;
    width: 616px;
    height: 603px;
    max-width: none;
    max-height: none;
    padding: 0;
    border: 0;
    box-sizing: border-box;
    overflow: visible;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__header,
  .fc-fmtdlg--mac .fc-fmtdlg__tabs,
  .fc-fmtdlg--mac .fc-fmtdlg__footer {
    display: contents;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__close {
    display: none;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__body {
    position: absolute;
    left: 13px;
    top: 38px;
    display: block;
    width: 590px;
    height: 515px;
    margin: 0;
    padding: 0;
    border: 0;
    box-sizing: border-box;
    overflow: visible;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__body .fc-fmtdlg__panel-tab {
    display: block;
    position: static;
    margin: 0;
    padding: 0;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab::before {
    display: none;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__hintbar {
    position: absolute;
    left: 29px;
    top: 522px;
    width: 558px;
    height: 18px;
    margin: 0;
    padding: 0;
    border: 0;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__cat-label {
    display: block;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__cat {
    padding: 0;
    border: 0;
    box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-border, currentColor);
  }
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__negative-list,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__pattern-list {
    padding: 0;
    border: 0;
    box-shadow: inset 0 0 0 1px var(--fc-fmtdlg-list-border, currentColor);
  }
  /* Wrappers whose children are placed by frame lose their own box. */
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__align-select-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__justify-distributed-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__indent-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__align-preview,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__align-preview-box,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__align-degree,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__text-control,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__section,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__protection-checks,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-style-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-color-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-color-row .fc-fmtdlg__color-control,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-presets,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-visual,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__border-stage,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__fill-bg-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__fill-pattern-color-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__fill-pattern-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__number-layout,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__cat-controls,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__preview,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__decimals-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__symbol-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__pattern-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__pattern-preset-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__locale-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__calendar-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__negative,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__pattern-list-wrap,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-family-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-size-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-style-label,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-underline--mac,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-color-row,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-deco,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__font-preview,
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__fill-sample {
    display: contents;
  }
  .fc-fmtdlg--mac .fc-fmtdlg__panel-tab .fc-fmtdlg__line-gallery {
    display: block;
  }
"""

WRAP=('protection-hint','border-hint','__hint')
def rule(prefix,sel,f,origin):
    x,y,w,h=f; ox,oy=origin
    extra='' if any(k in sel for k in WRAP) else '    white-space: nowrap;\n'
    if '__tab[' in sel: extra+='    z-index: 2;\n'
    full=', '.join(prefix+s.strip() for s in sel.split(',')) if ',' in sel and False else prefix+sel
    return f"  {full} {{\n    position: absolute;\n    left: {x-ox}px;\n    top: {y-oy}px;\n    width: {w}px;\n    height: {h}px;\n    margin: 0;\n    min-width: 0;\n    min-height: 0;\n    max-width: none;\n    max-height: none;\n    box-sizing: border-box;\n{extra}  }}\n"
out=["/* Generated from the Excel Mac Format Cells accessibility frames; every frame is\n   the window-relative rectangle, shifted here to its containing block. Do not\n   edit by hand: change the frame tables and regenerate. */\n\n@layer fc.app {\n"]
out.append(STRUCT)
P='.fc-fmtdlg--mac .fc-fmtdlg__panel '
for sel,f in CHROME: out.append(rule(P,sel,f,(0,0)))
for name,items in T.items():
    if ':' in name: continue
    pre=f'.fc-fmtdlg--mac .fc-fmtdlg__panel-tab[data-fc-tab="{name}"] '
    for sel,f in items:
        origin=(BX,BY)
        if name=='border' and sel.startswith('.fc-fmtdlg__line-style--'):
            origin=(BX+G[0]-BX,BY+G[1]-BY)  # relative to gallery: window - gallery origin
            origin=(G[0],G[1])
        out.append(rule(pre,sel,f,origin))

# containers
CONT={'.fc-fmtdlg__cat-item':(40,96),'.fc-fmtdlg__negative-item':NEGC[:2],'.fc-fmtdlg__pattern-item':PLC[:2]}
for sel,f in ALLNUM:
    origin=(BX,BY)
    for k,o in CONT.items():
        if k in sel: origin=o
    out.append(rule('.fc-fmtdlg--mac .fc-fmtdlg__panel-tab[data-fc-tab="number"] ',sel,f,origin))

HINT=[('.fc-fmtdlg__hintbar',(29,522,558,18)),
 ('.fc-fmtdlg__hintbar[data-fc-cat="general"], .fc-fmtdlg--mac .fc-fmtdlg__panel .fc-fmtdlg__hintbar[data-fc-cat="text"]',(183,147,393,34)),
 ('.fc-fmtdlg__hintbar[data-fc-cat="date"], .fc-fmtdlg--mac .fc-fmtdlg__panel .fc-fmtdlg__hintbar[data-fc-cat="time"]',(29,490,558,50))]
for sel,f in HINT: out.append(rule('.fc-fmtdlg--mac .fc-fmtdlg__panel ',sel,f,(0,0)))
out.append("}\n")
open(OUT,'w').write(''.join(out))
print('ok',sum(len(v) for v in T.values()))
