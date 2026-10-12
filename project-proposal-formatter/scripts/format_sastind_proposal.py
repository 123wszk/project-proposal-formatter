# -*- coding: utf-8 -*-
"""SASTIND 项目建议书格式调整（离线 python-docx 路径）—— 参考实现 / 起手模板。

用法: python format_sastind_proposal.py <in.docx> <out.docx> [--check]
（--check = 只诊断不打改动；设置环境变量 STAGES=1 会逐步另存 stage_XX_*.docx，便于二分定位损坏点）

对应 references/format-proposal-sastind.md 的 § 13 流程 2→12。**这是 D3000 实跑通过的实现，但含少量稿件特定硬编码，移植时必须改：**
  - 样式名（`建议书任务书一级标题` / `建议书正文` / `项目二级标题` 等）随机构模板而变，先跑 diagnose_docx.py 摸清实际样式名；
  - STEP 11 里"手打重复题注 `表 3` → 顺延为 `表 16`"以及随后的缓存顺延，是 D3000 特有的编号冲突，换稿必须重算（判据见 § 3.2.1）；
  - 交付前记得再跑一次 § 11 的"用 Word 回存刷新域"（本脚本不负责，否则图目录/表目录会留占位文字、主目录页码过期）。

关键防坑（都踩过，别删）：
  - § 7.1 分节符插入时的 footerReference/headerReference **去重**（不去重 Word 报"文件可能已经损坏"）；
  - 标题 run 级 `w:sz/w:b` 必须清掉，否则覆盖样式定义；
  - 正文小四必须落到段落级 `pPr/rPr/sz=24`，只改 run 级不彻底；
  - 表格居中要写 `tblPr/w:jc`，且只处理正文区顶层表（前置「项目基本信息表」不适用）；
  - STEP 12 插完分节符段落，要回头清掉首个一级标题的 `pageBreakBefore`（STEP 5 给它补过，
    当时前方只有目录 sdt 判不出分节；叠加有出空白页风险 —— V4.3）。
  - § 1.1 **全文关闭"与文档网格对齐"**（STEP 6，段落级 `w:snapToGrid w:val="0"`）：必须显式写
    `val="0"`（该元素 schema 默认值是 true=勾选，删元素 ≠ 取消）；用 `body.iter` 递归设，
    否则漏掉表格内段落；STEP 12 新建的段落要单独补（它们不在 STEP 6 的遍历里）。
"""
import sys, re, copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Twips
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

INPUT = sys.argv[1]
OUTPUT = sys.argv[2]
CHECK = '--check' in sys.argv
STAGES = __import__('os').environ.get('STAGES') == '1'
_stage_n = [0]
def snap(name):
    if STAGES:
        _stage_n[0] += 1
        fn = 'stage_%02d_%s.docx' % (_stage_n[0], name)
        doc.save(fn)
        print('SNAP', fn)

doc = Document(INPUT)
body = doc.element.body

CN_MAP = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,'十':10}
CN_REV = {v:k for k,v in CN_MAP.items()}
CN_DIGITS = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9}
def cn2num(s):
    if not s: return None
    if s == '十': return 10
    if '十' in s:
        a, _, b = s.partition('十')
        tens = CN_DIGITS.get(a, 1) if a else 1
        ones = CN_DIGITS.get(b, 0) if b else 0
        v = tens * 10 + ones
        return v if v > 0 else None
    if len(s) == 1:
        return CN_DIGITS.get(s)
    return None
def num2cn(n):
    if n <= 10: return CN_REV[n]
    if n < 20: return '十' + CN_REV[n-10]
    if n == 20: return '二十'
    if n < 30: return '二十' + CN_REV[n-20]
    return str(n)

LOG = []
def log(*a):
    s = ' '.join(str(x) for x in a)
    LOG.append(s)
    print(s)

# ---------------- low level xml helpers ----------------
PPR_ORDER = ['w:pStyle','w:keepNext','w:keepLines','w:pageBreakBefore','w:framePr','w:widowControl',
 'w:numPr','w:suppressLineNumbers','w:pBdr','w:shd','w:tabs','w:suppressAutoHyphens','w:kinsoku',
 'w:wordWrap','w:overflowPunct','w:topLinePunct','w:autoSpaceDE','w:autoSpaceDN','w:bidi',
 'w:adjustRightInd','w:snapToGrid','w:spacing','w:ind','w:contextualSpacing','w:mirrorIndents',
 'w:suppressOverlap','w:jc','w:textDirection','w:textAlignment','w:textboxTightWrap','w:outlineLvl',
 'w:divId','w:cnfStyle','w:rPr','w:sectPr','w:pPrChange']

RPR_ORDER = ['w:rStyle','w:rFonts','w:b','w:bCs','w:i','w:iCs','w:caps','w:smallCaps','w:strike',
 'w:dstrike','w:outline','w:shadow','w:emboss','w:imprint','w:noProof','w:snapToGrid','w:vanish',
 'w:webHidden','w:color','w:spacing','w:w','w:kern','w:position','w:sz','w:szCs','w:highlight','w:u',
 'w:effect','w:bdr','w:shd','w:fitText','w:vertAlign','w:rtl','w:cs','w:em','w:lang',
 'w:eastAsianLayout','w:specVanish','w:oMath']

def _ordered_set(parent, tag, attrs, order):
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    e = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        e.set(qn(k), str(v))
    idx = order.index(tag)
    for child in parent:
        cfull = 'w:' + etree.QName(child).localname
        if cfull in order and order.index(cfull) > idx:
            child.addprevious(e)
            return e
    parent.append(e)
    return e

def ppr_set(pPr, tag, attrs=None):
    return _ordered_set(pPr, tag, attrs, PPR_ORDER)

def ppr_no_grid(pPr):
    """§1.1 关闭「如果定义了文档网格，则与网格对齐」（= Word 段落对话框里那个勾选，取消）。
    OOXML: w:pPr/w:snapToGrid w:val="0"。
    ⚠️ 必须显式写 val="0" —— 该元素 schema 默认值是 true(=勾选)，把元素删掉等于没做。
    ⚠️ 位置由 PPR_ORDER 保证（紧跟 w:adjustRightInd、在 w:spacing / w:ind 之前），
       乱序插入 Word 会报"文件可能已经损坏"。
    ⚠️ 别和 run 级同名元素 w:rPr/w:snapToGrid（字符对字符网格）混淆，这里只动段落级那个。"""
    return ppr_set(pPr, 'w:snapToGrid', {'w:val': '0'})

def rpr_set(rPr, tag, attrs=None):
    return _ordered_set(rPr, tag, attrs, RPR_ORDER)

def rpr_drop(rPr, tags):
    for t in tags:
        e = rPr.find(qn(t))
        if e is not None:
            rPr.remove(e)

def rpr_font(rPr, ea='宋体', latin='Times New Roman'):
    # 去掉 hint，避免亚洲字体与西文混淆
    old = rPr.find(qn('w:rFonts'))
    if old is not None and old.get(qn('w:hint')) is not None:
        rpr_set(rPr, 'w:rFonts', {'w:ascii': latin, 'w:hAnsi': latin, 'w:eastAsia': ea, 'w:cs': latin})
    else:
        rpr_set(rPr, 'w:rFonts', {'w:ascii': latin, 'w:hAnsi': latin, 'w:eastAsia': ea, 'w:cs': latin})

def rpr_sz(rPr, val):
    rpr_set(rPr, 'w:sz', {'w:val': val})
    rpr_set(rPr, 'w:szCs', {'w:val': val})

def run_rpr(r):
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    return rPr

def para_rpr(pPr):
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        sect = pPr.find(qn('w:sectPr'))
        if sect is not None:
            sect.addprevious(rPr)
        else:
            pPr.append(rPr)
    return rPr

def get_pPr(p):
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)
    return pPr

# ---------------- locate boundary / children ----------------
children = list(body.iterchildren())
boundary = None
for i, ch in enumerate(children):
    if ch.tag == qn('w:p'):
        pp = Paragraph(ch, doc)
        st = pp.style.name or ''
        t = pp.text.strip()
        if '一级标题' in st and re.match(r'^[一二三四五六七八九十]+、', t):
            boundary = i
            break
if boundary is None:
    sys.exit('未找到正文边界（第一个一级标题）')
log('边界 body child index =', boundary, '::', Paragraph(children[boundary], doc).text.strip())
region = set(id(e) for e in children[boundary:])

H1_STYLE_OLD = '建议书任务书一级标题'
H1_STYLE = '建议书一级标题'
H2_STYLE = '建议书二级标题'
H3_STYLE = '建议书三级标题'

def is_para(e):
    return e.tag == qn('w:p')

def style_id_of(p):
    pPr = p.find(qn('w:pPr'))
    if pPr is None: return None
    e = pPr.find(qn('w:pStyle'))
    return e.get(qn('w:val')) if e is not None else None

def style_name_to_id(name):
    try:
        st = doc.styles[name]
        if st is not None:
            return st.style_id
    except Exception:
        pass
    for st in doc.styles.element.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        if nm is not None and nm.get(qn('w:val')) == name:
            return st.get(qn('w:styleId'))
    return None

SID = {n: style_name_to_id(n) for n in ['Normal', H1_STYLE_OLD, H2_STYLE, H3_STYLE, 'Caption', '建议书正文', '建议书任务书二级标题', '项目二级标题']}
log('styleIds:', SID)

# ============ STEP 2a: 空标题段落降级为 Normal ============
empty_headling = []
for e in children[boundary:]:
    if not is_para(e): continue
    pp = Paragraph(e, doc)
    st = pp.style.name or ''
    if st in (H1_STYLE_OLD, H2_STYLE, H3_STYLE, '建议书任务书二级标题', '项目二级标题') and not pp.text.strip():
        empty_headling.append((children.index(e), st))
if not CHECK:
    for e in children[boundary:]:
        if not is_para(e): continue
        pp = Paragraph(e, doc)
        st = pp.style.name or ''
        if st in (H1_STYLE_OLD, H2_STYLE, H3_STYLE, '建议书任务书二级标题', '项目二级标题') and not pp.text.strip():
            ppr_set(get_pPr(e), 'w:pStyle', {'w:val': SID['Normal']})
log('空标题段落降级为 Normal:', empty_headling)

# ============ STEP 2b: 章节漏号（numPr 假阳性） ============
def first_text_run(p):
    for r in p.iter(qn('w:r')):
        t = r.find(qn('w:t'))
        if t is not None:
            return r, t
    return None, None

chapter_paras = []   # (elem, num or None)
for e in children[boundary:]:
    if not is_para(e): continue
    st = Paragraph(e, doc).style.name or ''
    if st == H1_STYLE_OLD:
        t = Paragraph(e, doc).text.strip()
        if t:
            m = re.match(r'^([一二三四五六七八九十]+)、', t)
            chapter_paras.append((e, cn2num(m.group(1)) if m else None, t))

fixed = []
expected = 0
for e, num, t in chapter_paras:
    expected += 1
    if num is None:
        pPr = get_pPr(e)
        had_num = pPr.find(qn('w:numPr')) is not None
        if not CHECK:
            np_ = pPr.find(qn('w:numPr'))
            if np_ is not None:
                pPr.remove(np_)
            r, te = first_text_run(e)
            if te is not None:
                te.text = num2cn(expected) + '、' + (te.text or '')
                te.set(qn('xml:space'), 'preserve')
        fixed.append((expected, t, 'had numPr' if had_num else 'no numPr'))
    else:
        if num != expected:
            fixed.append((expected, t, '序号不符(实际%s)' % num2cn(num)))
log('章节编号修正:', fixed)
log('章节总数:', expected)
snap('02_chapters')

# ============ STEP 3: 标题样式定义 ============
def find_style(name):
    for st in doc.styles.element.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        if nm is not None and nm.get(qn('w:val')) == name:
            return st
    return None

def rename_style(old, new):
    st = find_style(old)
    if st is None: return False
    st.find(qn('w:name')).set(qn('w:val'), new)
    return True

def style_ppr(st):
    pPr = st.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        rPr = st.find(qn('w:rPr'))
        if rPr is not None: rPr.addprevious(pPr)
        else: st.append(pPr)
    return pPr

def style_rpr(st):
    rPr = st.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr'); st.append(rPr)
    return rPr

if not CHECK:
    # 一级：重命名 + 定义
    ok = rename_style(H1_STYLE_OLD, H1_STYLE)
    rename_style(H1_STYLE_OLD + ' Char', H1_STYLE + ' Char')
    log('重命名', H1_STYLE_OLD, '->', H1_STYLE, ok)

def define_heading_style(name, sz, bold, firstline, outline):
    st = find_style(name)
    if st is None:
        log('!! 样式不存在:', name); return
    pPr = style_ppr(st)
    rPr = style_rpr(st)
    rpr_font(rPr); 
    if bold:
        rpr_set(rPr, 'w:b', None); rpr_set(rPr, 'w:bCs', None)
    else:
        rpr_drop(rPr, ('w:b', 'w:bCs'))
    rpr_sz(rPr, sz)
    ppr_no_grid(pPr)   # §1.1：样式级兜底，之后新录入的段落自动继承
    _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
    _ordered_set(pPr, 'w:ind', {'w:firstLine': str(firstline), 'w:firstLineChars': '200'}, PPR_ORDER)
    _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
    _ordered_set(pPr, 'w:outlineLvl', {'w:val': outline}, PPR_ORDER)

if not CHECK:
    define_heading_style(H1_STYLE, 28, False, 560, '0')
    define_heading_style(H2_STYLE, 24, True, 480, '1')
    define_heading_style(H3_STYLE, 24, True, 480, '2')
    # 正文样式字号兜底（建议书正文 未定义 sz，会回落 Normal 五号）
    st = find_style('建议书正文')
    if st is not None:
        rpr_sz(style_rpr(st), 24)
    st = find_style('建议书任务书正文')
    if st is not None:
        rpr_sz(style_rpr(st), 24)
    # §1.1 兜底：Normal 与正文样式也写一份，让之后新录入的段落自动继承（段落级仍是主力）
    for _sn in ('Normal', '建议书正文', '建议书任务书正文'):
        _st = find_style(_sn)
        if _st is not None:
            ppr_no_grid(style_ppr(_st))
    log('标题样式定义完成')
snap('03_styles')

# ============ STEP 4: 标题段落识别 + 重映射 ============
L2_RE = re.compile(r'^(\d+)\.(\d+)(?![\d.])[\s\u4e00-\u9fff（(【]')
L3_RE = re.compile(r'^(\d+)\.(\d+)\.(\d+)[\s\u4e00-\u9fff（(【]')
CAP_RE = re.compile(r'^(图|表)\s*\d+')
END_RE = re.compile(r'[。；！？：]\s*$')
HEAD_END_RE = re.compile(r'[。；！？]\s*$')

headings = {'1': [], '2': [], '3': []}
for e in children[boundary:]:
    if not is_para(e): continue
    pp = Paragraph(e, doc)
    st = pp.style.name or ''
    t = pp.text.strip()
    if not t: continue
    if '一级标题' in st and re.match(r'^[一二三四五六七八九十]+、', t):
        headings['1'].append(e); continue
    if CAP_RE.match(t) and len(t) < 80:
        continue
    if L3_RE.match(t) and len(t) < 80 and not HEAD_END_RE.search(t):
        headings['3'].append(e); continue
    if L2_RE.match(t) and len(t) < 80 and not HEAD_END_RE.search(t):
        headings['2'].append(e); continue
log('标题数量  L1=%d L2=%d L3=%d' % (len(headings['1']), len(headings['2']), len(headings['3'])))
for k in ('2', '3'):
    log('  L%s:' % k, [Paragraph(e, doc).text.strip()[:28] for e in headings[k]])

if not CHECK:
    for lvl, sname, sz, bold, fl in (
        ('1', H1_STYLE, 28, False, 560),
        ('2', H2_STYLE, 24, True, 480),
        ('3', H3_STYLE, 24, True, 480)):
        sid = style_name_to_id(sname)
        for e in headings[lvl]:
            pPr = get_pPr(e)
            ppr_set(pPr, 'w:pStyle', {'w:val': sid})
            # 清理 run 级直接格式
            for r in e.iter(qn('w:r')):
                rPr = run_rpr(r)
                rpr_drop(rPr, ('w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u', 'w:highlight'))
                rpr_font(rPr)
            # 段落级字体/字号（让样式的 b 生效 → 不设 b）
            prPr = para_rpr(pPr)
            rpr_drop(prPr, ('w:sz', 'w:szCs', 'w:b', 'w:bCs', 'w:i', 'w:iCs'))
            rpr_font(prPr)
            rpr_sz(prPr, sz)
            _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
            _ordered_set(pPr, 'w:ind', {'w:firstLine': str(fl), 'w:firstLineChars': '200'}, PPR_ORDER)
            _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
            _ordered_set(pPr, 'w:outlineLvl', {'w:val': {'1': '0', '2': '1', '3': '2'}[lvl]}, PPR_ORDER)
    log('标题段落已重映射 + run 级格式清理')
snap('04_headings')

# ============ STEP 5: 一级标题分页 ============
def prev_has_break(e):
    pv = e.getprevious()
    if pv is None: return False
    if pv.tag == qn('w:p'):
        pPr = pv.find(qn('w:pPr'))
        if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
            return True
        for b in pv.iter(qn('w:br')):
            if b.get(qn('w:type')) == 'page':
                return True
    if pv.tag == qn('w:sdt'):
        ps = pv.findall('.//' + qn('w:p'))
        if ps:
            pPr = ps[-1].find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
                return True
    return False

pb_added = []
for e in headings['1']:
    pp = Paragraph(e, doc)
    if prev_has_break(e):
        continue
    pb_added.append(pp.text.strip()[:20])
    if not CHECK:
        ParagraphFormat(e).page_break_before = True
log('补充章节分页:', pb_added)
snap('05_pagebreak')

# ============ STEP 6: 全文 1.5 倍行距 + 关闭"与文档网格对齐" ============
# §1.1：全文所有段落取消「如果定义了文档网格，则与网格对齐」。
#   - 必须显式写 val="0"（见 ppr_no_grid 注释：默认值是 true，删元素 ≠ 取消）；
#   - body.iter 是**递归**的，表格 cell 内的段落也一并覆盖（这是"表格除外"最易翻车的地方）。
if not CHECK:
    n = 0
    for p in body.iter(qn('w:p')):
        ParagraphFormat(p).line_spacing = 1.5
        ppr_no_grid(get_pPr(p))
        n += 1
    log('设置 1.5 倍行距段落数:', n, '（同时关闭网格对齐 —— 同一批段落）')
snap('06_linespacing')

# ============ STEP 7: 题注识别 ============
captions = []
for e in children[boundary:]:
    if not is_para(e): continue
    pp = Paragraph(e, doc)
    t = pp.text
    ts = t.strip()
    if not ts: continue
    if CAP_RE.match(ts) and len(ts) < 80 and ('\t' not in t) and not END_RE.search(ts):
        captions.append(e)
log('题注数量:', len(captions))

# ============ STEP 8: 正文段落格式（边界之后，非标题非题注） ============
hset = set(id(x) for lv in headings.values() for x in lv)
cset = set(id(x) for x in captions)
body_count = 0
_body_targets = [e for e in children[boundary:] if is_para(e) and id(e) not in hset and id(e) not in cset]
log('正文段落（边界后、非标题/题注）总数:', len(_body_targets))
if not CHECK:
    for e in _body_targets:
        pp = Paragraph(e, doc)
        pPr = get_pPr(e)
        prPr = para_rpr(pPr)
        rpr_drop(prPr, ('w:sz', 'w:szCs', 'w:b', 'w:bCs'))
        rpr_font(prPr)
        rpr_sz(prPr, 24)
        for r in e.iter(qn('w:r')):
            rPr = run_rpr(r)
            rpr_font(rPr)
            rpr_sz(rPr, 24)
        _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
        _ordered_set(pPr, 'w:ind', {'w:firstLine': '480', 'w:firstLineChars': '200'}, PPR_ORDER)
        _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
        body_count += 1
log('正文段落格式化数:', body_count)
snap('08_body')

# ============ STEP 9: 题注统一 ============
if not CHECK:
    cap_sid = style_name_to_id('Caption')
    for e in captions:
        pPr = get_pPr(e)
        ppr_set(pPr, 'w:pStyle', {'w:val': cap_sid})
        prPr = para_rpr(pPr)
        rpr_drop(prPr, ('w:sz', 'w:szCs', 'w:b', 'w:bCs'))
        rpr_font(prPr)
        rpr_sz(prPr, 21)
        ind = pPr.find(qn('w:ind'))
        if ind is not None: pPr.remove(ind)
        _ordered_set(pPr, 'w:jc', {'w:val': 'center'}, PPR_ORDER)
        _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
        for r in e.iter(qn('w:r')):
            rPr = run_rpr(r)
            rpr_drop(rPr, ('w:b', 'w:bCs'))
            rpr_font(rPr)
            rpr_sz(rPr, 21)
    log('题注统一完成:', len(captions))
snap('09_captions')

# ============ STEP 10: 表格（正文区）宋体五号 ============
tbl_count = 0
_tbls = [e for e in children[boundary:] if e.tag == qn('w:tbl')]
log('正文区表格数:', len(_tbls))
TBLPR_ORDER = ['w:tblStyle','w:tblpPr','w:tblOverlap','w:bidiVisual','w:tblStyleRowBandSize',
    'w:tblStyleColBandSize','w:tblW','w:jc','w:tblCellSpacing','w:tblInd','w:tblBorders','w:shd',
    'w:tblLayout','w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange']
if not CHECK:
    for e in _tbls:
        tbl_count += 1
        # 表格居中（§4）
        tblPr = e.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr'); e.insert(0, tblPr)
        _ordered_set(tblPr, 'w:jc', {'w:val': 'center'}, TBLPR_ORDER)
        for p in e.iter(qn('w:p')):
            if p.find(qn('w:pPr')) is None:
                pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
            else:
                pPr = p.find(qn('w:pPr'))
            para_rpr(pPr)  # noop ensure
            prPr = para_rpr(pPr)
            rpr_drop(prPr, ('w:sz', 'w:szCs'))
            rpr_font(prPr)
            rpr_sz(prPr, 21)
            for r in p.iter(qn('w:r')):
                rPr = run_rpr(r)
                rpr_font(rPr)
                rpr_sz(rPr, 21)
log('正文区表格处理数:', tbl_count)
snap('10_tables')

if CHECK:
    sys.exit(0)

# ============ STEP 11: SEQ 题注修正（重复编号 表3 -> 表16） ============
def has_seq(p):
    return any('SEQ' in (x.text or '') for x in p.iter(qn('w:instrText')))

plain = [e for e in captions if not has_seq(e)]
log('纯文本题注:', [Paragraph(e, doc).text.strip() for e in plain])

existing_bm = set(b.get(qn('w:name')) for b in body.iter(qn('w:bookmarkStart')))
max_bmid = 0
for b in body.iter(qn('w:bookmarkStart')):
    try: max_bmid = max(max_bmid, int(b.get(qn('w:id'))))
    except Exception: pass

def new_bm_name(prefix):
    global _bmc
    while True:
        _bmc += 1
        nm = '%s9%05d' % (prefix, _bmc)
        if nm not in existing_bm:
            existing_bm.add(nm); return nm
_bmc = 0

def mk_rpr_el(sz=21):
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for k, v in (('w:ascii','Times New Roman'),('w:hAnsi','Times New Roman'),('w:eastAsia','宋体'),('w:cs','Times New Roman')):
        rf.set(qn(k), v)
    rPr.append(rf)
    b = OxmlElement('w:b'); b.set(qn('w:val'), '0'); rPr.append(b)
    bcs = OxmlElement('w:bCs'); bcs.set(qn('w:val'), '0'); rPr.append(bcs)
    s = OxmlElement('w:sz'); s.set(qn('w:val'), str(sz)); rPr.append(s)
    sc = OxmlElement('w:szCs'); sc.set(qn('w:val'), str(sz)); rPr.append(sc)
    return rPr

def mk_run(text, sz=21):
    r = OxmlElement('w:r'); r.append(mk_rpr_el(sz))
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t); return r

def mk_fld(kind):
    r = OxmlElement('w:r'); r.append(mk_rpr_el())
    fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), kind)
    r.append(fc); return r

def mk_instr(txt):
    r = OxmlElement('w:r'); r.append(mk_rpr_el())
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = txt
    r.append(it); return r

def mk_bm_start(name):
    global max_bmid
    max_bmid += 1
    b = OxmlElement('w:bookmarkStart'); b.set(qn('w:id'), str(max_bmid)); b.set(qn('w:name'), name)
    return b, max_bmid

def mk_bm_end(bid):
    b = OxmlElement('w:bookmarkEnd'); b.set(qn('w:id'), str(bid)); return b

# 表 1..15 保持；光罩表变 表16；其后 表16..21 -> 17..22
for e in plain:
    txt = Paragraph(e, doc).text.strip()
    m = re.match(r'^(表|图)\s*\d+\s*(.*)$', txt)
    label, desc = m.group(1), m.group(2)
    pPr = e.find(qn('w:pPr'))
    for ch in list(e):
        if ch is not pPr: e.remove(ch)
    # 清除 ind
    if pPr is not None:
        ind = pPr.find(qn('w:ind'))
        if ind is not None: pPr.remove(ind)
    refn = new_bm_name('_Ref')
    tocn = new_bm_name('_Toc')
    bs1, bid1 = mk_bm_start(refn)
    e.append(bs1)
    e.append(mk_run(label + ' '))
    e.append(mk_fld('begin'))
    e.append(mk_instr(' SEQ %s \\* ARABIC ' % label))
    e.append(mk_fld('separate'))
    e.append(mk_run('16'))
    e.append(mk_fld('end'))
    e.append(mk_bm_end(bid1))
    bs2, bid2 = mk_bm_start(tocn)
    e.append(bs2)
    e.append(mk_run(' ' + desc))
    e.append(mk_bm_end(bid2))
    log('已把纯文本题注转换为 SEQ 域:', txt, '->', '%s 16 %s' % (label, desc))

# 顺延后续 表 SEQ 缓存编号
shift_count = 0
seen_plain = False
for e in children[boundary:]:
    if not is_para(e): continue
    if id(e) in set(id(x) for x in plain):
        seen_plain = True; continue
    if not seen_plain: continue
    p = e
    if not any('SEQ 表' in (x.text or '') for x in p.iter(qn('w:instrText'))): continue
    # 找 separate 后第一个含文本的 run
    ran = list(p.iter(qn('w:r')))
    state = 0
    for r in ran:
        if r.find(qn('w:fldChar')) is not None:
            fc = r.find(qn('w:fldChar')).get(qn('w:fldCharType'))
            if fc == 'separate': state = 1
            elif fc == 'end': state = 0
            continue
        if state == 1:
            t = r.find(qn('w:t'))
            if t is not None and (t.text or '').strip().isdigit():
                t.text = str(int(t.text.strip()) + 1)
                shift_count += 1
            break
log('表 SEQ 缓存顺延数:', shift_count)

# REF 缓存顺延（表17 -> 表18）
ref_shift = 0
for p in body.iter(qn('w:p')):
    instr = [ (x.text or '').strip() for x in p.iter(qn('w:instrText')) ]
    if not any(s.startswith('REF') for s in instr): continue
    ran = list(p.iter(qn('w:r')))
    state = 0
    for r in ran:
        if r.find(qn('w:fldChar')) is not None:
            fc = r.find(qn('w:fldChar')).get(qn('w:fldCharType'))
            if fc == 'separate': state = 1
            elif fc == 'end': state = 0
            continue
        if state == 1:
            t = r.find(qn('w:t'))
            if t is not None:
                mm = re.match(r'^表\s*(\d+)$', (t.text or '').strip())
                if mm and int(mm.group(1)) >= 16:
                    t.text = '表 %d' % (int(mm.group(1)) + 1)
                    ref_shift += 1
            state = 0
log('REF 缓存顺延数:', ref_shift)
snap('11_seq')

# ============ STEP 12: 插入 图目录 / 表目录 / 目录-正文分节符 ============
def mk_title_para(text):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    ppr_no_grid(pPr)   # §1.1：新建段落不在 STEP 6 的遍历里，必须单独补
    _ordered_set(pPr, 'w:spacing', {'w:before':'0','w:after':'0','w:line':'360','w:lineRule':'auto'}, PPR_ORDER)
    _ordered_set(pPr, 'w:ind', {'w:left':'0','w:leftChars':'0','w:right':'0','w:rightChars':'0','w:firstLine':'0','w:firstLineChars':'0'}, PPR_ORDER)
    _ordered_set(pPr, 'w:jc', {'w:val':'center'}, PPR_ORDER)
    pPr.append(mk_rpr_el(24))
    p.append(pPr)
    p.append(mk_run(text, 24))
    return p

def mk_toc_field_para(instr, placeholder):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    ppr_no_grid(pPr)   # §1.1：同上
    _ordered_set(pPr, 'w:spacing', {'w:line':'360','w:lineRule':'auto'}, PPR_ORDER)
    p.append(pPr)
    p.append(mk_fld('begin'))
    r2 = OxmlElement('w:r'); r2.append(mk_rpr_el(24))
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text = instr
    r2.append(it); p.append(r2)
    p.append(mk_fld('separate'))
    p.append(mk_run(placeholder, 24))
    p.append(mk_fld('end'))
    return p

# 找到 TOC sdt 元素
toc_sdt = None
for sdt in body.iter(qn('w:sdt')):
    if any('TOC' in (x.text or '') for x in sdt.iter(qn('w:instrText'))):
        toc_sdt = sdt; break
if toc_sdt is None:
    log('!! 未找到主目录 sdt')
else:
    # 复制 sectPr #6（终止目录所在分节的那个）
    sects = list(body.iter(qn('w:sectPr')))
    src_sect = None
    idx_sdt = list(body.iterchildren()).index(toc_sdt)
    for sp in sects:
        # sectPr 所属段落的位置
        par = sp.getparent()
        if par.tag == qn('w:pPr'):
            par = par.getparent()
        if par.tag != qn('w:p'):
            continue
        try:
            ipos = list(body.iterchildren()).index(par)
        except ValueError:
            continue
        if ipos > idx_sdt:
            src_sect = sp; break
    new_sect = copy.deepcopy(src_sect)
    # pgNumType -> lowerRoman start 1（注意 OOXML 子元素顺序：必须在 pgSz/pgMar/cols 之间）
    SECTPR_ORDER = ['w:footnotePr','w:endnotePr','w:type','w:pgSz','w:pgMar','w:paperSrc',
        'w:pgBorders','w:lnNumType','w:pgNumType','w:cols','w:formProt','w:vAlign','w:noEndnote',
        'w:titlePg','w:textDirection','w:bidi','w:rtlGutter','w:docGrid','w:printerSettings',
        'w:sectPrChange']
    old = new_sect.find(qn('w:pgNumType'))
    if old is not None: new_sect.remove(old)
    _ordered_set(new_sect, 'w:pgNumType', {'w:fmt':'lowerRoman','w:start':'1'}, SECTPR_ORDER)
    # ---- 关键修复（V4.2）：同一 footer/header rId 不能被两个 section 同时引用 ----
    # 实测：把带 footerReference 的 sectPr 复制到新分节（原分节保留同一 rId）会让 Word 报"文件可能已经损坏"。
    # 修法：新分节独享 footer/header 引用，原 sectPr 移除这些引用，改为继承（紧随其后，视觉完全一致）。
    removed_refs = []
    for _tg in ('w:footerReference', 'w:headerReference'):
        for _ref in list(src_sect.findall(qn(_tg))):
            removed_refs.append((_tg.split('}')[-1], _ref.get(qn('w:type'))))
            src_sect.remove(_ref)
    log('分节引用去重（新分节独享，原分节改为继承）:', removed_refs)
    sect_para = OxmlElement('w:p')
    spPr = OxmlElement('w:pPr')
    ppr_no_grid(spPr)   # §1.1：分节符载体段同样是新建段落
    _ordered_set(spPr, 'w:spacing', {'w:line':'240','w:lineRule':'auto'}, PPR_ORDER)
    spPr.append(new_sect)
    sect_para.append(spPr)
    # 先插分节段落
    toc_sdt.addnext(sect_para)
    # 再在 sdt 与 sect_para 之间插 图目录/表目录
    tbl = mk_toc_field_para(' TOC \\h \\z \\c "表" ', '（表目录：更新域后生成）')
    bkl = mk_toc_field_para(' TOC \\h \\z \\c "图" ', '（图目录：更新域后生成）')
    t2 = mk_title_para('表目录'); t2p = t2.find(qn('w:pPr'))
    _ordered_set(t2p, 'w:pageBreakBefore', None, PPR_ORDER)
    t1 = mk_title_para('图目录'); t1p = t1.find(qn('w:pPr'))
    _ordered_set(t1p, 'w:pageBreakBefore', None, PPR_ORDER)
    cur = toc_sdt
    for el in (t1, bkl, t2, tbl):
        cur.addnext(el); cur = el
    log('已插入 图目录 / 表目录 / 目录-正文分节符')

    # ---- 关键修复（V4.3）：分节符已经让正文起新页，首个一级标题的 pageBreakBefore 要清掉 ----
    # 步骤 5.5 给「一、立项的必要性」补过 pageBreakBefore（当时它前面只有目录 sdt，判不出"已分页"），
    # 步骤 12 又在它前面插了分节符段落。两者叠加有出空白页的风险（§2 自动化注意第 7 条），
    # 这里保守处理：只要是"开启态 + 前方现在确有分节符"，就把该标题的 PBB 移除。
    if headings['1']:
        _e = headings['1'][0]
        _pv = _e.getprevious()
        _pv_sect = (is_para(_pv) and _pv.find(qn('w:pPr')) is not None
                    and _pv.find(qn('w:pPr')).find(qn('w:sectPr')) is not None)
        _pPr = get_pPr(_e)
        _pbb = _pPr.find(qn('w:pageBreakBefore'))
        if _pv_sect and _pbb is not None and (_pbb.get(qn('w:val')) in (None, '1', 'true', 'on')):
            _pPr.remove(_pbb)
            log('清除首个一级标题 pageBreakBefore（前方已插分节符）:',
                Paragraph(_e, doc).text.strip()[:20])
snap('12_toc')

# ============ STEP 13: 前置页脚核查（前置内容本就无页码） ============
# 注：原稿封面 / 填写说明 / 承诺书 / 项目基本信息表 各节的页脚本就为空（无 PAGE 域），无需清空。
# 旧版曾用 `part._blob = ...` 覆写页脚，但 FooterPart 是 XmlPart，其 blob 由 _element 序列化，
# 对 _blob 赋值会被静默忽略（无效），故移除该步骤，改为核查并记录。
for _i, _s in enumerate(body.iter(qn('w:sectPr'))):
    _ftrs = [r.get(qn('r:id')) for r in _s.findall(qn('w:footerReference'))]
    if _ftrs:
        _info = []
        for _rid in _ftrs:
            try:
                _xml = doc.part.rels[_rid].target_part.blob.decode('utf-8', 'ignore')
                _info.append('%s:%s' % (_rid, 'PAGE' if 'PAGE' in _xml else 'empty'))
            except Exception:
                _info.append('%s:?' % _rid)
        log('  sec%d 页脚 %s' % (_i, _info))

snap('13_footer')
doc.save(OUTPUT)
log('已保存:', OUTPUT)
print('\n'.join(LOG[-6:]))
