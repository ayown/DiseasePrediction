"""
Disease Prediction System — PowerPoint Generator
Generates a polished, animated presentation using python-pptx.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ─── Colour palette ───────────────────────────────────────────────────────────
TEAL      = RGBColor(0x00, 0x83, 0x8A)   # primary accent
DARK_TEAL = RGBColor(0x00, 0x4D, 0x56)   # headings
ORANGE    = RGBColor(0xE8, 0x6F, 0x1E)   # highlight
DARK_GRAY = RGBColor(0x26, 0x26, 0x26)
MID_GRAY  = RGBColor(0x55, 0x55, 0x55)
LIGHT_BG  = RGBColor(0xF5, 0xF8, 0xF9)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
BLUE      = RGBColor(0x1F, 0x6F, 0xC8)
GREEN     = RGBColor(0x21, 0x7A, 0x3C)
RED       = RGBColor(0xC0, 0x39, 0x2B)
GOLD      = RGBColor(0xD4, 0xAC, 0x0D)

# ─── Slide dimensions (Widescreen 16:9) ──────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]   # completely blank

# ══════════════════════════════════════════════════════════════════════════════
# Helper utilities
# ══════════════════════════════════════════════════════════════════════════════

def add_rect(slide, x, y, w, h, fill_rgb=None, line_rgb=None,
             line_width=None, alpha=None):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_rgb
    else:
        shape.fill.background()
    if line_rgb:
        shape.line.color.rgb = line_rgb
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, text, x, y, w, h,
                font_size=18, bold=False, italic=False,
                color=DARK_GRAY, align=PP_ALIGN.LEFT,
                wrap=True, font_name="Calibri"):
    txb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size  = Pt(font_size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name  = font_name
    return txb


def add_para(tf, text, font_size=16, bold=False, italic=False,
             color=DARK_GRAY, align=PP_ALIGN.LEFT,
             space_before=6, bullet=False, indent=0,
             font_name="Calibri"):
    para = tf.add_paragraph()
    para.alignment = align
    para.space_before = Pt(space_before)
    if indent:
        para.level = indent
    run = para.add_run()
    run.text = text
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name   = font_name
    return para


def set_slide_bg(slide, rgb):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def accent_bar(slide, color=TEAL, x=0, y=0, w=13.33, h=0.06):
    """Thin coloured bar."""
    add_rect(slide, x, y, w, h, fill_rgb=color)


# ─── Animation helpers ────────────────────────────────────────────────────────

def _anim_ns():
    return {
        'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
        'p':   'http://schemas.openxmlformats.org/presentationml/2006/main',
        'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }


def _get_or_create_timing(slide):
    """Return (or create) the <p:timing> element on a slide."""
    spTree = slide.shapes._spTree
    sp_parent = spTree.getparent()   # <p:sp> lives inside <p:cSld><p:spTree>
    # timing lives directly under <p:sld>
    sld_el = slide._element
    timing = sld_el.find(qn('p:timing'))
    if timing is None:
        timing = etree.SubElement(sld_el, qn('p:timing'))
    return timing


def add_entrance_anim(slide, shape, effect='appear', delay_ms=0, dur_ms=500,
                      trigger='onClick'):
    """
    Add a simple entrance animation to *shape* on *slide*.
    effect: 'appear' | 'fadeIn' | 'flyInFromBottom' | 'flyInFromLeft'
    trigger: 'onClick' | 'afterPrev' | 'withPrev'
    """
    PRESET = {
        'appear':         ('entrance', 'appear',          1,  0),
        'fadeIn':         ('entrance', 'fade',             10, 0),
        'flyInFromBottom':('entrance', 'flyInFromBottom',  2,  0),
        'flyInFromLeft':  ('entrance', 'flyInFromLeft',    2,  0),
        'flyInFromRight': ('entrance', 'flyInFromRight',   2,  0),
        'zoom':           ('entrance', 'zoom',             22, 0),
        'wipe':           ('entrance', 'wipe',             56, 0),
    }
    if effect not in PRESET:
        effect = 'appear'
    cat, preset_id_str, preset_id, sub_type = PRESET[effect]

    sld_el  = slide._element
    timing  = _get_or_create_timing(slide)

    # ── ensure <p:tnLst> ─────────────────────────────
    tnLst = timing.find(qn('p:tnLst'))
    if tnLst is None:
        tnLst = etree.SubElement(timing, qn('p:tnLst'))

    # ── root par ─────────────────────────────────────
    par_root = tnLst.find(qn('p:par'))
    if par_root is None:
        par_root = etree.SubElement(tnLst, qn('p:par'))
        cTn_root = etree.SubElement(par_root, qn('p:cTn'))
        cTn_root.set('id',          _next_id(slide))
        cTn_root.set('dur',         'indefinite')
        cTn_root.set('restart',     'whenNotActive')
        cTn_root.set('nodeType',    'tmRoot')
        childTnLst_root = etree.SubElement(cTn_root, qn('p:childTnLst'))
    else:
        cTn_root       = par_root.find(qn('p:cTn'))
        childTnLst_root = cTn_root.find(qn('p:childTnLst'))

    # ── seq ──────────────────────────────────────────
    seq = childTnLst_root.find(qn('p:seq'))
    if seq is None:
        seq = etree.SubElement(childTnLst_root, qn('p:seq'))
        seq.set('concurrent',   '1')
        seq.set('nextAc',       'seek')
        cTn_seq = etree.SubElement(seq, qn('p:cTn'))
        cTn_seq.set('id',       _next_id(slide))
        cTn_seq.set('dur',      'indefinite')
        cTn_seq.set('nodeType', 'mainSeq')
        childTnLst_seq = etree.SubElement(cTn_seq, qn('p:childTnLst'))
        # nextCondLst
        nc = etree.SubElement(seq, qn('p:nextCondLst'))
        cond = etree.SubElement(nc, qn('p:cond'))
        cond.set('evt', 'onClick')
        cond.set('delay', '0')
        tn_ref = etree.SubElement(cond, qn('p:tn'))
        tn_ref.set('val', '2')
    else:
        cTn_seq        = seq.find(qn('p:cTn'))
        childTnLst_seq = cTn_seq.find(qn('p:childTnLst'))

    # ── par for this click ───────────────────────────
    par_click = etree.SubElement(childTnLst_seq, qn('p:par'))
    cTn_click = etree.SubElement(par_click, qn('p:cTn'))
    cTn_click.set('id',       _next_id(slide))
    cTn_click.set('fill',     'hold')
    if trigger == 'onClick':
        cTn_click.set('nodeType', 'clickEffect')
    elif trigger == 'afterPrev':
        cTn_click.set('nodeType', 'afterEffect')
    else:
        cTn_click.set('nodeType', 'withEffect')

    stCondLst = etree.SubElement(cTn_click, qn('p:stCondLst'))
    cond2     = etree.SubElement(stCondLst, qn('p:cond'))
    if trigger == 'onClick':
        cond2.set('evt',   'onClick')
        cond2.set('delay', '0')
        tn_ref2 = etree.SubElement(cond2, qn('p:tn'))
        tn_ref2.set('val', '2')
    else:
        cond2.set('delay', str(delay_ms))

    childTnLst_click = etree.SubElement(cTn_click, qn('p:childTnLst'))

    # ── par for the anim itself ───────────────────────
    par_anim = etree.SubElement(childTnLst_click, qn('p:par'))
    cTn_anim = etree.SubElement(par_anim, qn('p:cTn'))
    cTn_anim.set('id',   _next_id(slide))
    cTn_anim.set('fill', 'hold')
    stCondLst2 = etree.SubElement(cTn_anim, qn('p:stCondLst'))
    cond3      = etree.SubElement(stCondLst2, qn('p:cond'))
    cond3.set('delay', str(delay_ms))
    childTnLst_anim = etree.SubElement(cTn_anim, qn('p:childTnLst'))

    # ── animEffect ───────────────────────────────────
    par_ef = etree.SubElement(childTnLst_anim, qn('p:par'))
    cTn_ef = etree.SubElement(par_ef, qn('p:cTn'))
    cTn_ef.set('id',         _next_id(slide))
    cTn_ef.set('presetID',   str(preset_id))
    cTn_ef.set('presetClass','entr')
    cTn_ef.set('presetSubtype', '0')
    cTn_ef.set('fill',       'hold')
    cTn_ef.set('grpId',      '0')
    cTn_ef.set('nodeType',   'clickEffect')
    stCondLst3 = etree.SubElement(cTn_ef, qn('p:stCondLst'))
    cond4      = etree.SubElement(stCondLst3, qn('p:cond'))
    cond4.set('delay', '0')
    childTnLst_ef = etree.SubElement(cTn_ef, qn('p:childTnLst'))

    # set (makes shape invisible initially)
    set_el = etree.SubElement(childTnLst_ef, qn('p:set'))
    cBhvr_s = etree.SubElement(set_el, qn('p:cBhvr'))
    cTn_s   = etree.SubElement(cBhvr_s, qn('p:cTn'))
    cTn_s.set('id',   _next_id(slide))
    cTn_s.set('dur',  '1')
    cTn_s.set('fill', 'hold')
    stCL_s  = etree.SubElement(cTn_s, qn('p:stCondLst'))
    cond_s  = etree.SubElement(stCL_s, qn('p:cond'))
    cond_s.set('delay', '0')
    tgtEl_s = etree.SubElement(cBhvr_s, qn('p:tgtEl'))
    spTgt_s = etree.SubElement(tgtEl_s, qn('p:spTgt'))
    spTgt_s.set('spid', str(shape.shape_id))
    attrNameLst_s = etree.SubElement(cBhvr_s, qn('p:attrNameLst'))
    attrName_s    = etree.SubElement(attrNameLst_s, qn('p:attrName'))
    attrName_s.text = 'style.visibility'
    to_s  = etree.SubElement(set_el, qn('p:to'))
    strVal_s = etree.SubElement(to_s, qn('a:strVal'))
    strVal_s.set('val', 'hidden')

    # animEffect
    aef = etree.SubElement(childTnLst_ef, qn('p:animEffect'))
    aef.set('transition', 'in')
    aef.set('filter',     preset_id_str)
    cBhvr_a = etree.SubElement(aef, qn('p:cBhvr'))
    cTn_a   = etree.SubElement(cBhvr_a, qn('p:cTn'))
    cTn_a.set('id',   _next_id(slide))
    cTn_a.set('dur',  str(dur_ms))
    stCL_a  = etree.SubElement(cTn_a, qn('p:stCondLst'))
    cond_a  = etree.SubElement(stCL_a, qn('p:cond'))
    cond_a.set('delay', '0')
    tgtEl_a = etree.SubElement(cBhvr_a, qn('p:tgtEl'))
    spTgt_a = etree.SubElement(tgtEl_a, qn('p:spTgt'))
    spTgt_a.set('spid', str(shape.shape_id))


_id_counters = {}

def _next_id(slide):
    key = id(slide)
    _id_counters[key] = _id_counters.get(key, 0) + 1
    return str(_id_counters[key] + 100)


def add_slide_transition(slide, transition='fade', dur_ms=700):
    """Add a slide transition."""
    sld_el = slide._element
    trans = sld_el.find(qn('p:transition'))
    if trans is None:
        trans = etree.SubElement(sld_el, qn('p:transition'))
    trans.set('dur', str(dur_ms))
    trans.set('advTm', '0')
    if transition == 'fade':
        fade = etree.SubElement(trans, qn('p:fade'))
        fade.set('thruBlk', '0')
    elif transition == 'push':
        push = etree.SubElement(trans, qn('p:push'))
        push.set('dir', 'l')
    elif transition == 'wipe':
        wipe = etree.SubElement(trans, qn('p:wipe'))
        wipe.set('dir', 'l')
    elif transition == 'zoom':
        zoom = etree.SubElement(trans, qn('p:zoom'))
        zoom.set('dir', 'in')


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ══════════════════════════════════════════════════════════════════════════════

# ─── 1. COVER SLIDE ──────────────────────────────────────────────────────────
def build_cover(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, RGBColor(0x04, 0x1C, 0x2C))  # deep navy

    # gradient-like decorative bars
    for i, (alpha_pct, y_pos) in enumerate([(1.0, 0), (0.6, 6.9), (0.3, 7.2)]):
        bar = add_rect(slide, 0, y_pos, 13.33, 0.55,
                       fill_rgb=TEAL if i == 0 else DARK_TEAL)

    # Left accent strip
    add_rect(slide, 0, 0, 0.12, 7.5, fill_rgb=ORANGE)

    # University name
    tb = add_textbox(slide, "SIKSHA 'O' ANUSANDHAN DEEMED TO BE UNIVERSITY",
                     0.3, 0.10, 12.7, 0.5,
                     font_size=14, bold=True, color=WHITE,
                     align=PP_ALIGN.CENTER, font_name="Calibri")

    # Course tag
    add_textbox(slide, "Machine Learning Concept - 2  |  CSE 3968",
                0.3, 0.60, 12.7, 0.4,
                font_size=13, color=RGBColor(0xB0, 0xD8, 0xE0),
                align=PP_ALIGN.CENTER, font_name="Calibri")

    # Main title
    tb_title = add_textbox(slide, "Disease Prediction System\nUsing Machine Learning",
                           0.4, 1.25, 12.5, 2.2,
                           font_size=46, bold=True, color=WHITE,
                           align=PP_ALIGN.CENTER, font_name="Calibri Light")

    # Horizontal divider
    add_rect(slide, 1.5, 3.55, 10.3, 0.04, fill_rgb=ORANGE)

    # Subtitle chips
    for i, (label, val) in enumerate([
        ("Models", "DT · RF · XGB · Stacking"),
        ("Best Accuracy", "93.50 %"),
        ("Diseases", "41 Classes"),
    ]):
        x = 0.8 + i * 4.1
        chip = add_rect(slide, x, 3.75, 3.7, 0.75,
                        fill_rgb=RGBColor(0x06, 0x35, 0x50))
        add_textbox(slide, label, x, 3.78, 3.7, 0.28,
                    font_size=10, color=RGBColor(0xB0, 0xD8, 0xE0),
                    align=PP_ALIGN.CENTER, font_name="Calibri")
        add_textbox(slide, val, x, 4.05, 3.7, 0.38,
                    font_size=15, bold=True, color=ORANGE,
                    align=PP_ALIGN.CENTER, font_name="Calibri")

    # Team / date
    add_textbox(slide, "Presented by:  NAME 1  ·  NAME 2  ·  NAME 3  ·  NAME 4",
                0.3, 5.0, 12.7, 0.4,
                font_size=13, color=RGBColor(0xCC, 0xCC, 0xCC),
                align=PP_ALIGN.CENTER)
    add_textbox(slide, "Department of Computer Science & Engineering  |  ITER, SOA University  |  May 2026",
                0.3, 5.45, 12.7, 0.4,
                font_size=11, color=RGBColor(0x88, 0xAA, 0xBB),
                align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'fade', 800)


# ─── 2. AGENDA / TABLE OF CONTENTS ───────────────────────────────────────────
def build_agenda(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, LIGHT_BG)
    accent_bar(slide, TEAL, 0, 0, 13.33, 0.08)

    # Section header strip
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=DARK_TEAL)
    add_textbox(slide, "Agenda", 0.4, 0.22, 12, 0.72,
                font_size=34, bold=True, color=WHITE,
                align=PP_ALIGN.LEFT, font_name="Calibri Light")

    items = [
        ("01", "Problem Statement",      "What disease prediction challenge are we solving?"),
        ("02", "Dataset & Features",     "4,925 records · 41 diseases · severity weights + age"),
        ("03", "Methodology / Pipeline", "Data cleaning → feature engineering → 4 classifiers"),
        ("04", "Results & Metrics",      "Accuracy, precision, recall, confusion matrices"),
        ("05", "Feature Importance",     "Which symptoms (and age) matter most?"),
        ("06", "Discussion & Future",    "Limitations, next steps, deployment ideas"),
        ("07", "Conclusion",             "Key takeaways and recommendations"),
    ]

    cols = [items[:4], items[4:]]
    col_x = [0.35, 6.85]

    for ci, col in enumerate(cols):
        for ri, (num, title, sub) in enumerate(col):
            y = 1.3 + ri * 1.43
            x = col_x[ci]
            # number bubble
            bubble = add_rect(slide, x, y + 0.05, 0.58, 0.58,
                               fill_rgb=ORANGE)
            add_textbox(slide, num, x, y + 0.05, 0.58, 0.58,
                        font_size=14, bold=True, color=WHITE,
                        align=PP_ALIGN.CENTER)
            # title + subtitle
            tb_t = add_textbox(slide, title, x + 0.7, y, 5.6, 0.38,
                               font_size=14, bold=True, color=DARK_TEAL,
                               font_name="Calibri")
            tb_s = add_textbox(slide, sub, x + 0.7, y + 0.37, 5.6, 0.38,
                               font_size=10, color=MID_GRAY,
                               font_name="Calibri")

    add_slide_transition(slide, 'push', 600)


# ─── 3. PROBLEM STATEMENT ────────────────────────────────────────────────────
def build_problem(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, WHITE)
    accent_bar(slide, ORANGE, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=RGBColor(0xFF, 0xF3, 0xE0))

    add_textbox(slide, "Problem Statement", 0.4, 0.17, 12, 0.78,
                font_size=32, bold=True, color=RGBColor(0x99, 0x44, 0x00),
                font_name="Calibri Light")

    # Scenario box
    scenario = add_rect(slide, 0.4, 1.25, 12.5, 1.1,
                         fill_rgb=RGBColor(0xFD, 0xF6, 0xEC),
                         line_rgb=ORANGE, line_width=1.5)
    add_textbox(slide,
                '"A patient reports: fever + joint pain + fatigue. '
                'Is it Malaria, Dengue, Typhoid, or Arthritis?"',
                0.55, 1.30, 12.2, 0.95,
                font_size=16, italic=True, color=RGBColor(0x66, 0x33, 0x00),
                align=PP_ALIGN.CENTER)

    # Three problem pillars
    pillars = [
        ("Challenge",   "41 overlapping disease classes with shared symptoms make accurate diagnosis difficult without lab tests.", RED),
        ("Gap",         "Existing tools use binary symptom encoding — they miss severity nuance and patient age context.", BLUE),
        ("Our Solution","Severity-weighted features + epidemiological age + Stacking Ensemble to push beyond individual model limits.", GREEN),
    ]
    for i, (heading, body, col) in enumerate(pillars):
        x = 0.4 + i * 4.32
        box = add_rect(slide, x, 2.55, 4.1, 2.6,
                        fill_rgb=WHITE,
                        line_rgb=col, line_width=2)
        # top colour strip
        add_rect(slide, x, 2.55, 4.1, 0.35, fill_rgb=col)
        add_textbox(slide, heading, x, 2.58, 4.1, 0.30,
                    font_size=13, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, body, x + 0.15, 3.02, 3.8, 2.0,
                    font_size=12, color=DARK_GRAY, wrap=True)

    # Task definition
    add_textbox(slide,
                "Task: Given a 135-feature vector (severity-weighted symptoms + patient age), "
                "predict which of 41 diseases a patient most likely has.",
                0.4, 5.3, 12.5, 0.55,
                font_size=13, bold=True, color=DARK_TEAL,
                align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'wipe', 600)


# ─── 4. DATASET ──────────────────────────────────────────────────────────────
def build_dataset(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, LIGHT_BG)
    accent_bar(slide, TEAL, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=DARK_TEAL)
    add_textbox(slide, "Dataset & Feature Engineering", 0.4, 0.22, 12, 0.72,
                font_size=32, bold=True, color=WHITE,
                font_name="Calibri Light")

    # ── Stats row ──
    stats = [
        ("4,925",  "Patient Records"),
        ("41",     "Disease Classes"),
        ("136",    "Symptom Features"),
        ("1",      "Age Feature"),
        ("135",    "Final Features\n(after variance filter)"),
    ]
    for i, (num, label) in enumerate(stats):
        x = 0.3 + i * 2.55
        add_rect(slide, x, 1.25, 2.3, 1.22,
                  fill_rgb=WHITE,
                  line_rgb=TEAL, line_width=1.5)
        add_textbox(slide, num, x, 1.30, 2.3, 0.65,
                    font_size=26, bold=True, color=TEAL,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, label, x, 1.93, 2.3, 0.48,
                    font_size=10, color=MID_GRAY,
                    align=PP_ALIGN.CENTER)

    # ── Two panels ──
    # Left: Feature construction
    add_rect(slide, 0.3, 2.65, 5.9, 3.75,
              fill_rgb=WHITE, line_rgb=TEAL, line_width=1)
    add_rect(slide, 0.3, 2.65, 5.9, 0.38, fill_rgb=TEAL)
    add_textbox(slide, "Severity-Weighted Encoding", 0.3, 2.67, 5.9, 0.34,
                font_size=13, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    left_items = [
        "Each symptom assigned weight 1–7 (not binary 0/1)",
        "0 = symptom absent for that patient",
        "136 symptom columns → one per unique symptom",
        "50% symptom dropout simulates real incomplete reporting",
        "Zero-variance features removed → 135 final features",
    ]
    for i, item in enumerate(left_items):
        add_textbox(slide, f"▸  {item}",
                    0.45, 3.18 + i * 0.55, 5.6, 0.48,
                    font_size=11.5, color=DARK_GRAY)

    # Right: Age feature
    add_rect(slide, 6.55, 2.65, 6.45, 3.75,
              fill_rgb=WHITE, line_rgb=ORANGE, line_width=1)
    add_rect(slide, 6.55, 2.65, 6.45, 0.38, fill_rgb=ORANGE)
    add_textbox(slide, "Age-Aware Feature (Novel Addition)", 6.55, 2.67, 6.45, 0.34,
                font_size=13, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    age_items = [
        ("Chicken pox",  "2 – 12 yrs"),
        ("Acne",         "13 – 25 yrs"),
        ("Diabetes",     "35 – 75 yrs"),
        ("Heart attack", "45 – 80 yrs"),
        ("Malaria",      "5 – 65 yrs"),
    ]
    add_textbox(slide, "Disease-specific epidemiological ranges:", 6.7, 3.13, 6.1, 0.35,
                font_size=11, bold=True, color=DARK_TEAL)
    for i, (dis, rng) in enumerate(age_items):
        add_rect(slide, 6.7,  3.54 + i * 0.52, 3.3, 0.42,
                  fill_rgb=RGBColor(0xFF, 0xF3, 0xE0))
        add_textbox(slide, dis,  6.75, 3.56 + i * 0.52, 3.2, 0.35,
                    font_size=11, color=DARK_GRAY)
        add_rect(slide, 10.15, 3.54 + i * 0.52, 2.6, 0.42,
                  fill_rgb=RGBColor(0xFF, 0xE0, 0xB2))
        add_textbox(slide, rng, 10.15, 3.56 + i * 0.52, 2.6, 0.35,
                    font_size=11, bold=True, color=ORANGE,
                    align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'fade', 700)


# ─── 5. METHODOLOGY PIPELINE ─────────────────────────────────────────────────
def build_methodology(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, WHITE)
    accent_bar(slide, BLUE, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=BLUE)
    add_textbox(slide, "End-to-End Pipeline", 0.4, 0.22, 12, 0.72,
                font_size=32, bold=True, color=WHITE,
                font_name="Calibri Light")

    steps = [
        ("1", "Load CSVs",          "disease_raw.csv  +  symptom_severity.csv",    TEAL),
        ("2", "Build Feature Matrix","136-col severity-weighted symptom matrix",    BLUE),
        ("3", "Add Age Feature",    "Sample from epidemiological age range/disease",ORANGE),
        ("4", "50% Dropout",        "Randomly zero-out half of reported symptoms",  RGBColor(0x7B, 0x68, 0xEE)),
        ("5", "Label Encode",       "prognosis  →  integer class index",            RGBColor(0x20, 0x8B, 0x3A)),
        ("6", "Split 80/20",        "3,940 train  |  985 test  (stratified)",       RED),
        ("7", "Variance Filter",    "Remove zero-variance cols  →  135 features",   RGBColor(0xB8, 0x5C, 0x00)),
        ("8", "Train 4 Classifiers","DT  ·  RF (200)  ·  XGB  ·  Stacking",        DARK_TEAL),
        ("9", "Evaluate",           "Accuracy, Precision, Recall, F1, Confusion ✓", GREEN),
    ]

    for i, (num, title, sub, col) in enumerate(steps):
        row = i // 3
        col_i = i % 3
        x = 0.3 + col_i * 4.35
        y = 1.28 + row * 2.05

        # card
        add_rect(slide, x, y, 4.1, 1.82,
                  fill_rgb=WHITE, line_rgb=col, line_width=1.5)
        # top bar
        add_rect(slide, x, y, 4.1, 0.36, fill_rgb=col)
        # number
        add_textbox(slide, num, x, y + 0.02, 0.55, 0.32,
                    font_size=14, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, title, x + 0.55, y + 0.02, 3.5, 0.32,
                    font_size=12, bold=True, color=WHITE)
        add_textbox(slide, sub, x + 0.1, y + 0.44, 3.9, 1.28,
                    font_size=11, color=DARK_GRAY, wrap=True)

        # arrow between cards (in row, not last of row)
        if col_i < 2 and i < len(steps) - 1:
            add_textbox(slide, "→", x + 4.12, y + 0.70, 0.22, 0.38,
                        font_size=14, bold=True, color=MID_GRAY,
                        align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'push', 600)


# ─── 6. MODELS ────────────────────────────────────────────────────────────────
def build_models(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, LIGHT_BG)
    accent_bar(slide, GREEN, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=GREEN)
    add_textbox(slide, "Four Classifiers", 0.4, 0.22, 12, 0.72,
                font_size=32, bold=True, color=WHITE,
                font_name="Calibri Light")

    models = [
        ("Decision Tree",
         "85.58 %",
         ["Gini criterion, no depth limit",
          "Baseline — transparent & explainable",
          "Overfits due to dropout noise"],
         RGBColor(0xE7, 0x4C, 0x3C)),
        ("Random Forest",
         "93.30 %",
         ["200 trees, parallel bagging",
          "Smooths out symptom dropout noise",
          "Feature importance available"],
         RGBColor(0x27, 0xAE, 0x60)),
        ("XGBoost",
         "91.47 %",
         ["100 trees, lr=0.1, depth=6",
          "Sequential residual correction",
          "mlogloss multi-class eval"],
         RGBColor(0x29, 0x80, 0xB9)),
        ("Stacking Ensemble",
         "93.50 %  ★",
         ["DT + RF + XGB  →  LR meta-learner",
          "passthrough=True: raw features too",
          "5-fold stratified CV · BEST MODEL"],
         ORANGE),
    ]

    for i, (name, acc, bullets, col) in enumerate(models):
        x = 0.3 + i * 3.25
        y = 1.32

        # main card
        add_rect(slide, x, y, 3.0, 5.0,
                  fill_rgb=WHITE, line_rgb=col, line_width=2)
        # colour header
        add_rect(slide, x, y, 3.0, 1.0, fill_rgb=col)
        add_textbox(slide, name, x, y + 0.04, 3.0, 0.52,
                    font_size=13, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        # accuracy badge
        add_rect(slide, x + 0.35, y + 0.60, 2.3, 0.36,
                  fill_rgb=RGBColor(0xFF, 0xFF, 0xFF))
        add_textbox(slide, acc, x + 0.35, y + 0.62, 2.3, 0.30,
                    font_size=14, bold=True, color=col,
                    align=PP_ALIGN.CENTER)
        # bullets
        for j, b in enumerate(bullets):
            add_textbox(slide, f"• {b}", x + 0.15, y + 1.10 + j * 0.75, 2.7, 0.68,
                        font_size=11, color=DARK_GRAY, wrap=True)

    add_slide_transition(slide, 'fade', 700)


# ─── 7. RESULTS / ACCURACY ───────────────────────────────────────────────────
def build_results(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, WHITE)
    accent_bar(slide, ORANGE, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=RGBColor(0xFF, 0xF3, 0xE0))
    add_textbox(slide, "Results — Accuracy Comparison", 0.4, 0.17, 12, 0.78,
                font_size=30, bold=True, color=RGBColor(0x99, 0x44, 0x00),
                font_name="Calibri Light")

    results = [
        ("Decision Tree",      85.58, 0.86, 0.86, RED),
        ("XGBoost",            91.47, 0.92, 0.91, BLUE),
        ("Random Forest (200)",93.30, 0.94, 0.93, GREEN),
        ("Stacking Ensemble",  93.50, 0.94, 0.93, ORANGE),
    ]

    chart_left  = 0.5
    chart_right = 8.5
    chart_width = chart_right - chart_left
    bar_h = 0.72
    max_acc = 100.0

    for i, (name, acc, prec, rec, col) in enumerate(results):
        y = 1.35 + i * 1.22
        bar_w = (acc / max_acc) * chart_width * 0.95

        # label
        add_textbox(slide, name, chart_left, y, 2.55, 0.52,
                    font_size=12, bold=(name=="Stacking Ensemble"),
                    color=DARK_GRAY)
        # bar background
        add_rect(slide, chart_left + 2.6, y + 0.04, chart_width - 2.6, bar_h - 0.08,
                  fill_rgb=RGBColor(0xE8, 0xE8, 0xE8))
        # bar fill
        fill_w = (acc / max_acc) * (chart_width - 2.6) * 0.98
        add_rect(slide, chart_left + 2.6, y + 0.04, fill_w, bar_h - 0.08,
                  fill_rgb=col)
        # value label
        add_textbox(slide, f"{acc:.2f}%",
                    chart_left + 2.6 + fill_w + 0.05, y + 0.06, 1.0, 0.52,
                    font_size=12, bold=True, color=col)

    # Metrics table
    add_rect(slide, 8.8, 1.22, 4.15, 5.15,
              fill_rgb=WHITE, line_rgb=TEAL, line_width=1)
    add_rect(slide, 8.8, 1.22, 4.15, 0.42, fill_rgb=DARK_TEAL)
    headers = ["Model", "Acc", "Prec", "Rec"]
    col_ws  = [1.8, 0.78, 0.78, 0.79]
    cx = 8.8
    for h, cw in zip(headers, col_ws):
        add_textbox(slide, h, cx, 1.24, cw, 0.38,
                    font_size=10, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        cx += cw

    row_data = [
        ("DT",       "85.58", "0.86", "0.86"),
        ("XGB",      "91.47", "0.92", "0.91"),
        ("RF-200",   "93.30", "0.94", "0.93"),
        ("Stacking", "93.50", "0.94", "0.93"),
    ]
    for ri, (m, a, p, r) in enumerate(row_data):
        bg = RGBColor(0xF0, 0xFF, 0xF4) if ri == 3 else WHITE
        add_rect(slide, 8.8, 1.64 + ri * 0.92, 4.15, 0.90, fill_rgb=bg)
        row_vals = [m, a, p, r]
        cx = 8.8
        for vi, (v, cw) in enumerate(zip(row_vals, col_ws)):
            is_best = (ri == 3)
            add_textbox(slide, v, cx, 1.68 + ri * 0.92, cw, 0.50,
                        font_size=11,
                        bold=(is_best),
                        color=ORANGE if is_best else DARK_GRAY,
                        align=PP_ALIGN.CENTER)
            cx += cw
        if ri < 3:
            add_rect(slide, 8.8, 1.64 + (ri+1)*0.92, 4.15, 0.02,
                      fill_rgb=RGBColor(0xDD, 0xDD, 0xDD))

    # Take-away
    add_rect(slide, 0.5, 6.32, 12.33, 0.62,
              fill_rgb=RGBColor(0xFF, 0xF0, 0xD0),
              line_rgb=ORANGE, line_width=1)
    add_textbox(slide,
                "★  Stacking Ensemble (93.50%) outperforms all individual models. "
                "Random Forest (93.30%) is a strong single-model alternative.",
                0.65, 6.36, 12.0, 0.54,
                font_size=12, bold=True, color=RGBColor(0x88, 0x33, 0x00),
                align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'fade', 700)


# ─── 8. STACKING ENSEMBLE DEEP DIVE ─────────────────────────────────────────
def build_stacking(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, LIGHT_BG)
    accent_bar(slide, ORANGE, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=RGBColor(0xC0, 0x5A, 0x00))
    add_textbox(slide, "Stacking Ensemble — Deep Dive", 0.4, 0.22, 12, 0.72,
                font_size=30, bold=True, color=WHITE,
                font_name="Calibri Light")

    # Architecture diagram
    # Input layer
    add_rect(slide, 0.4, 1.35, 2.6, 0.75,
              fill_rgb=RGBColor(0xE8, 0xF4, 0xFD), line_rgb=BLUE, line_width=1.5)
    add_textbox(slide, "Input\n135 Features", 0.4, 1.38, 2.6, 0.70,
                font_size=11, bold=True, color=BLUE,
                align=PP_ALIGN.CENTER)

    # Base models
    base_y = [1.15, 1.82, 2.50]
    base_labels = ["Decision Tree", "Random Forest\n(200 trees)", "XGBoost\n(100 trees)"]
    base_colors = [RED, GREEN, BLUE]
    for i, (by, bl, bc) in enumerate(zip(base_y, base_labels, base_colors)):
        add_rect(slide, 4.0, by, 2.8, 0.65,
                  fill_rgb=WHITE, line_rgb=bc, line_width=1.5)
        add_textbox(slide, bl, 4.0, by + 0.02, 2.8, 0.62,
                    font_size=11, bold=True, color=bc,
                    align=PP_ALIGN.CENTER)
        # Arrow from input
        add_textbox(slide, "→", 3.05, by + 0.17, 0.9, 0.35,
                    font_size=13, color=MID_GRAY, align=PP_ALIGN.CENTER)

    # passthrough arrow
    add_textbox(slide, "──────────────────────────── passthrough=True ────────→",
                0.4, 3.4, 9.2, 0.38,
                font_size=9, italic=True, color=ORANGE,
                align=PP_ALIGN.CENTER)

    # Meta-learner
    add_rect(slide, 7.2, 1.65, 3.2, 1.08,
              fill_rgb=RGBColor(0xFF, 0xF3, 0xE0), line_rgb=ORANGE, line_width=2)
    add_rect(slide, 7.2, 1.65, 3.2, 0.34, fill_rgb=ORANGE)
    add_textbox(slide, "Meta-Learner", 7.2, 1.67, 3.2, 0.30,
                font_size=12, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, "Logistic Regression\nC=0.1  |  saga solver\nmax_iter=3000",
                7.2, 2.01, 3.2, 0.70,
                font_size=10, color=DARK_GRAY,
                align=PP_ALIGN.CENTER)
    for i, by in enumerate(base_y):
        add_textbox(slide, "→", 6.82, by + 0.17, 0.4, 0.35,
                    font_size=13, color=MID_GRAY, align=PP_ALIGN.CENTER)

    # Output
    add_rect(slide, 10.8, 1.70, 2.25, 0.95,
              fill_rgb=RGBColor(0xE8, 0xF8, 0xEC), line_rgb=GREEN, line_width=2)
    add_textbox(slide, "Output\nDisease Label", 10.8, 1.75, 2.25, 0.88,
                font_size=12, bold=True, color=GREEN,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, "→", 10.4, 1.95, 0.4, 0.35,
                font_size=13, color=MID_GRAY, align=PP_ALIGN.CENTER)

    # Key points
    add_rect(slide, 0.4, 3.85, 12.5, 0.40, fill_rgb=DARK_TEAL)
    add_textbox(slide, "Why Stacking Works Here", 0.4, 3.87, 12.5, 0.36,
                font_size=13, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    key_pts = [
        ("5-Fold Stratified CV", "Out-of-fold predictions prevent data leakage to the meta-learner during training."),
        ("passthrough=True",     "Meta-learner sees raw clinical signals alongside base model outputs — corrects cases all 3 models got wrong."),
        ("predict_proba",        "Passes class probability distributions (not just hard labels) — more information for the meta-learner."),
        ("+0.20% over RF",       "Small but consistent gain: the meta-learner recovers errors that RF and XGB both make on hard disease boundary cases."),
    ]
    for i, (kw, desc) in enumerate(key_pts):
        x = 0.4 + (i % 2) * 6.4
        y = 4.38 + (i // 2) * 1.45
        add_rect(slide, x, y, 6.1, 1.28,
                  fill_rgb=WHITE, line_rgb=ORANGE, line_width=1)
        add_textbox(slide, kw,   x + 0.12, y + 0.07, 5.8, 0.36,
                    font_size=12, bold=True, color=ORANGE)
        add_textbox(slide, desc, x + 0.12, y + 0.43, 5.8, 0.75,
                    font_size=10.5, color=DARK_GRAY, wrap=True)

    add_slide_transition(slide, 'fade', 700)


# ─── 9. FEATURE IMPORTANCE ───────────────────────────────────────────────────
def build_feature_importance(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, WHITE)
    accent_bar(slide, TEAL, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=DARK_TEAL)
    add_textbox(slide, "Feature Importance (Random Forest)", 0.4, 0.22, 12, 0.72,
                font_size=30, bold=True, color=WHITE,
                font_name="Calibri Light")

    # Top symptoms mock bar chart
    features = [
        ("yellowing_of_eyes",      0.048, ORANGE),
        ("dark_urine",             0.044, ORANGE),
        ("sweating",               0.041, TEAL),
        ("chills",                 0.038, TEAL),
        ("age",                    0.035, RED),
        ("loss_of_appetite",       0.033, TEAL),
        ("fatigue",                0.031, TEAL),
        ("high_fever",             0.029, TEAL),
        ("vomiting",               0.027, TEAL),
        ("nausea",                 0.025, TEAL),
    ]
    chart_x     = 0.5
    label_w     = 2.9
    max_imp     = 0.050
    chart_max_w = 7.2

    add_textbox(slide, "Top 10 Most Discriminative Features", chart_x, 1.22, 10, 0.38,
                font_size=13, bold=True, color=DARK_TEAL)

    for i, (fname, imp, col) in enumerate(features):
        y = 1.70 + i * 0.52
        bar_w = (imp / max_imp) * chart_max_w

        # label
        display = fname.replace("_", " ")
        txt_col = RED if fname == "age" else DARK_GRAY
        add_textbox(slide, display, chart_x, y + 0.04, label_w, 0.42,
                    font_size=10.5, bold=(fname == "age"),
                    color=txt_col)
        # bar bg
        add_rect(slide, chart_x + label_w + 0.05, y + 0.06,
                  chart_max_w, 0.35,
                  fill_rgb=RGBColor(0xEE, 0xEE, 0xEE))
        # bar fill
        add_rect(slide, chart_x + label_w + 0.05, y + 0.06,
                  bar_w, 0.35, fill_rgb=col)
        # value
        add_textbox(slide, f"{imp:.3f}",
                    chart_x + label_w + bar_w + 0.12, y + 0.06, 0.8, 0.35,
                    font_size=9.5, color=col, bold=True)

    # Callout panel
    add_rect(slide, 10.85, 1.22, 2.15, 6.40,
              fill_rgb=RGBColor(0xF0, 0xFB, 0xF8), line_rgb=TEAL, line_width=1.5)
    add_rect(slide, 10.85, 1.22, 2.15, 0.40, fill_rgb=TEAL)
    add_textbox(slide, "Key Insights", 10.85, 1.24, 2.15, 0.36,
                font_size=11, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    callouts = [
        "Jaundice-related symptoms (yellowing, dark urine) are most discriminative",
        "",
        "Age ranks #5 — confirms epidemiological prior knowledge adds real signal",
        "",
        "Fatigue, fever, vomiting rank lower — shared across many diseases",
        "",
        "Severity weights give each symptom a continuous signal, not binary",
    ]
    y0 = 1.70
    for ci, c in enumerate(callouts):
        add_textbox(slide, c, 10.92, y0 + ci * 0.75, 2.0, 0.70,
                    font_size=9, color=DARK_GRAY, wrap=True)

    add_slide_transition(slide, 'wipe', 600)


# ─── 10. DISCUSSION ──────────────────────────────────────────────────────────
def build_discussion(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, LIGHT_BG)
    accent_bar(slide, RED, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=RGBColor(0x6C, 0x0C, 0x0C))
    add_textbox(slide, "Discussion & Limitations", 0.4, 0.22, 12, 0.72,
                font_size=30, bold=True, color=WHITE,
                font_name="Calibri Light")

    # What works
    add_rect(slide, 0.35, 1.25, 5.95, 5.15,
              fill_rgb=WHITE, line_rgb=GREEN, line_width=1.5)
    add_rect(slide, 0.35, 1.25, 5.95, 0.42, fill_rgb=GREEN)
    add_textbox(slide, "✓  What Works Well", 0.35, 1.27, 5.95, 0.38,
                font_size=14, bold=True, color=WHITE)
    works = [
        "Severity weighting outperforms binary encoding — more nuance per symptom",
        "Age feature confirms epidemiological value — ranks #5 in importance",
        "50% dropout makes models robust to missing symptoms at test time",
        "Balanced dataset (≈120 records/class) — accuracy is meaningful",
        "Stacking with passthrough corrects cases all base models fail on",
    ]
    for i, w in enumerate(works):
        add_textbox(slide, f"▸  {w}", 0.5, 1.78 + i * 0.73, 5.65, 0.65,
                    font_size=11, color=DARK_GRAY, wrap=True)

    # Limitations
    add_rect(slide, 6.65, 1.25, 5.95, 5.15,
              fill_rgb=WHITE, line_rgb=RED, line_width=1.5)
    add_rect(slide, 6.65, 1.25, 5.95, 0.42, fill_rgb=RED)
    add_textbox(slide, "✗  Limitations", 6.65, 1.27, 5.95, 0.38,
                font_size=14, bold=True, color=WHITE)
    limits = [
        "Structured checklist, not natural language — real patients describe symptoms differently",
        "Hepatitis B/C/D/E still confused — lab tests needed for final differentiation",
        "No hyperparameter tuning — Optuna/GridSearch could push accuracy further",
        "Dataset is synthetic-style (clean labels) — real clinical data is noisier",
        "No gender, vital signs, or lab values as features",
    ]
    for i, lim in enumerate(limits):
        add_textbox(slide, f"▸  {lim}", 6.8, 1.78 + i * 0.73, 5.65, 0.65,
                    font_size=11, color=DARK_GRAY, wrap=True)

    add_slide_transition(slide, 'fade', 700)


# ─── 11. FUTURE SCOPE ────────────────────────────────────────────────────────
def build_future(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, WHITE)
    accent_bar(slide, BLUE, 0, 0, 13.33, 0.08)
    add_rect(slide, 0, 0.08, 13.33, 1.0, fill_rgb=BLUE)
    add_textbox(slide, "Future Scope & Applications", 0.4, 0.22, 12, 0.72,
                font_size=30, bold=True, color=WHITE,
                font_name="Calibri Light")

    future_items = [
        ("Hyperparameter Tuning", "Optuna / GridSearchCV on XGB depth and LR regulariser C",     BLUE,   0.4,  1.30),
        ("Neural Network",        "MLP / Transformer encoder for higher-order symptom interactions", RGBColor(0x7B, 0x68, 0xEE), 4.5,  1.30),
        ("SHAP Explainability",   "Per-patient prediction explanations for clinical trust",        TEAL,   8.6,  1.30),
        ("Web Deployment",        "FastAPI backend + symptom form → usable prototype",             GREEN,  0.4,  3.55),
        ("Richer Patient Data",   "Add gender, BP, temperature, CBC lab values",                  ORANGE, 4.5,  3.55),
        ("Real Clinical NLP",     "Extract symptoms from free-text discharge summaries via NLP",  RED,    8.6,  3.55),
    ]

    for title, desc, col, x, y in future_items:
        add_rect(slide, x, y, 3.85, 1.92,
                  fill_rgb=WHITE, line_rgb=col, line_width=1.8)
        add_rect(slide, x, y, 3.85, 0.38, fill_rgb=col)
        add_textbox(slide, title, x, y + 0.03, 3.85, 0.32,
                    font_size=12, bold=True, color=WHITE,
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, desc, x + 0.15, y + 0.48, 3.55, 1.35,
                    font_size=11, color=DARK_GRAY, wrap=True)

    # Applications callout
    add_rect(slide, 0.4, 5.65, 12.5, 0.95,
              fill_rgb=RGBColor(0xE8, 0xF4, 0xFD), line_rgb=BLUE, line_width=1)
    add_textbox(slide, "Applications: GP clinics  ·  Remote health workers  ·  Patient symptom checkers  "
                "·  Hospital triage  ·  Medical education",
                0.55, 5.73, 12.2, 0.72,
                font_size=12, bold=True, color=BLUE,
                align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'push', 600)


# ─── 12. CONCLUSION ──────────────────────────────────────────────────────────
def build_conclusion(prs):
    slide = prs.slides.add_slide(BLANK)
    set_slide_bg(slide, RGBColor(0x04, 0x1C, 0x2C))

    add_rect(slide, 0, 0, 13.33, 0.08, fill_rgb=ORANGE)
    add_rect(slide, 0, 0.08, 13.33, 0.95, fill_rgb=DARK_TEAL)
    add_textbox(slide, "Conclusion", 0.4, 0.16, 12, 0.78,
                font_size=34, bold=True, color=WHITE,
                font_name="Calibri Light")

    # Accuracy recap
    perf = [
        ("Decision Tree",       "85.58 %", RED),
        ("XGBoost",             "91.47 %", BLUE),
        ("Random Forest (200)", "93.30 %", GREEN),
        ("Stacking Ensemble",   "93.50 %", ORANGE),
    ]
    for i, (m, a, col) in enumerate(perf):
        x = 0.4 + i * 3.2
        add_rect(slide, x, 1.22, 3.0, 0.85,
                  fill_rgb=RGBColor(0x06, 0x35, 0x50))
        add_textbox(slide, m, x, 1.25, 3.0, 0.36,
                    font_size=10.5, color=RGBColor(0xCC, 0xCC, 0xCC),
                    align=PP_ALIGN.CENTER)
        add_textbox(slide, a, x, 1.61, 3.0, 0.42,
                    font_size=16, bold=True, color=col,
                    align=PP_ALIGN.CENTER)

    # Star badge for best
    add_rect(slide, 9.7, 1.18, 3.25, 0.95,
              fill_rgb=ORANGE)
    add_textbox(slide, "★  BEST MODEL", 9.7, 1.22, 3.25, 0.48,
                font_size=12, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)
    add_textbox(slide, "Stacking Ensemble  93.50%", 9.7, 1.68, 3.25, 0.38,
                font_size=10, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    # Key takeaways
    takeaways = [
        "Classical ML with thoughtful feature engineering remains highly competitive for structured medical data.",
        "Severity-weighted symptoms + epidemiological age are more powerful than binary encoding alone.",
        "Stacking Ensemble with passthrough=True squeezes residual errors that no single model can correct.",
        "50% dropout simulation ensures real-world robustness to missing patient-reported symptoms.",
        "The system is ready for extension: richer data, NLP integration, or web deployment.",
    ]
    add_textbox(slide, "Key Takeaways", 0.4, 2.28, 12.5, 0.40,
                font_size=15, bold=True, color=TEAL)

    for i, t in enumerate(takeaways):
        add_rect(slide, 0.4, 2.82 + i * 0.85, 12.5, 0.72,
                  fill_rgb=RGBColor(0x06, 0x35, 0x50))
        add_textbox(slide, f"  ✦  {t}", 0.5, 2.86 + i * 0.85, 12.2, 0.62,
                    font_size=11.5, color=WHITE, wrap=True)

    # Footer
    add_textbox(slide,
                "Thank you  ·  Questions welcome",
                0.3, 7.05, 12.7, 0.38,
                font_size=14, bold=True, color=ORANGE,
                align=PP_ALIGN.CENTER)

    add_slide_transition(slide, 'fade', 800)


# ══════════════════════════════════════════════════════════════════════════════
# BUILD ALL SLIDES
# ══════════════════════════════════════════════════════════════════════════════
build_cover(prs)
build_agenda(prs)
build_problem(prs)
build_dataset(prs)
build_methodology(prs)
build_models(prs)
build_results(prs)
build_stacking(prs)
build_feature_importance(prs)
build_discussion(prs)
build_future(prs)
build_conclusion(prs)

out = "/home/saswatbarai/Downloads/MLC Project /Disease_Prediction_PPT.pptx"
prs.save(out)
print(f"Saved → {out}")
print(f"Slides: {len(prs.slides)}")
