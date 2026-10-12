# -*- coding: utf-8 -*-
"""SASTIND 项目任务书格式调整（离线 python-docx 路径）—— 参考实现 / 起手模板。

用法: python format_sastind_taskbook.py <in.docx> <out.docx>   (STAGES=1 分步另存便于二分定位)

对应 references/format-taskbook-sastind.md 的 § 13 流程。**D3000 家族任务书实跑通过，但含稿件特定假设，移植前先读下面"必须按稿调整"清单。**

与建议书实现（format_sastind_proposal.py）的关键差异：
  - 任务书稿**可能完全没有标题样式**（D3000 任务书稿标题都是 Normal/Body Text），故边界与标题
    识别改为"文本模式 + 位置 + 排除 toc 样式 + 短句不以句读结尾"，而不是"是否套一级标题样式"；
  - 四个级别标题都是 宋体 小四 加粗（一级不再四号）；需**新建** 4 个样式；
  - 正文区行距 **1.25**（w:line=300）**仅限"正文文字段落"**；而 各级标题 / 表格内 / 前置 / 题注 /
    目录条目 保持 1.5（360）—— 用户口径是"正文 1.25（表格除外），其他部分默认与建议书一致"，
    故标题**不**参与 1.25 收紧（STEP 7 把 headings 排除在 body_targets 之外）；
  - **一级标题不强制分页**（V1.2 口径，与建议书**相反**）：STEP 4 把一级标题的
    `w:pageBreakBefore` **一律删除**（开态关态都删），并顺带删掉紧邻前方的"纯分页符段落"；
    建议书脚本 STEP 5 则是"缺分页就补上"。**两边的分页逻辑不可互换。**
  - 题注默认**改为连续编号 图1..图n / 表1..表n** 并构造成 SEQ 域（_Ref/_Toc 双书签）——
    这是"题注必须连续编号"口径下的做法；若该稿要保留章节式编号（图3.1），须改 STEP 9；
  - 本脚本**不插**图目录/表目录、**不动**分节结构（是否加图表目录按用户决定）。

必须按稿调整的硬编码：
  - STEP 2 的"编号异常订正"（`N.N.N.` 数字后多一个点）是 D3000 稿实测的 2 处，换稿要重扫；
  - STEP 3 的边界/标题识别阈值（长度 < 60、不以句读结尾）要与新稿实际排版习惯核对；
  - STEP 9 的题注编号策略（连续 vs 章节式）；
  - STEP 10 会**清掉主目录的过期缓存条目**改为占位提示——若该稿主目录本来是对的，跳过本步。

关键防坑（都踩过，别删）：
  - 内建样式在 styles.xml 里 `w:name` 常是**小写**（`"caption"`）：按 `=='Caption'` 精确匹配会找不到，
    题注就套不上 Caption 样式。用大小写无关匹配（或 python-docx 的 `styles['Caption']`）。
  - **TOC 域的闭合不能用"见到 end 就停"**：TOC 段内嵌 HYPERLINK/PAGEREF 域，首段就有多个 end。
    必须用**域深度**走查（begin +1 / end −1，回到 0 才是 TOC 真正闭合）。
  - 标题 run 级 `w:sz/w:b` 必须清掉，否则覆盖样式定义（任务书四级别同字号，残留最刺眼）；
  - 正文小四必须落到段落级 `pPr/rPr/sz=24`，只改 run 级不彻底；
  - 行距顺序：**先全文打底 1.5 → 再只把正文区"非表格/非题注"段落改 1.25**；按边界整片刷会把表格带成 1.25；
  - § 1.1 **全文关闭"与文档网格对齐"**（STEP 5，段落级 `w:snapToGrid w:val="0"`）：1.25 倍（`w:line=300`）
    通常不是文档网格行高的整数倍，勾着网格对齐会被吸附成 1 倍/1.5 倍 —— **"设了 1.25 却没生效"多半是它**。
    必须显式写 `val="0"`（元素 schema 默认值是 true=勾选，删元素 ≠ 取消）；用 `body.iter` 递归设，
    否则漏掉表格内段落。
  - 表格居中要写 `tblPr/w:jc`，且只处理正文区顶层表（前置「项目基本信息表」不适用）。
  - **清章节分页时别误伤两样东西**：① `w:sectPr` 分节符（承担页码体系，删了页码就乱）；
    ② 与正文同段的手工分页符（`w:br` 和文本可能在同一 `w:r` 里）。只删"纯分页符段落"。
"""
import sys, os, re, copy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.text.parfmt import ParagraphFormat
from docx.enum.style import WD_STYLE_TYPE
from lxml import etree

INPUT, OUTPUT = sys.argv[1], sys.argv[2]
STAGES = os.environ.get('STAGES') == '1'
_n = [0]
doc = Document(INPUT)
body = doc.element.body

def snap(name):
    if STAGES:
        _n[0] += 1
        fn = 'stage_%02d_%s.docx' % (_n[0], name)
        doc.save(fn); print('SNAP', fn)

def log(*a):
    print(' '.join(str(x) for x in a))

# ---------------- OOXML 子元素顺序（必须遵守，否则 Word 报"文件可能已经损坏"）----------------
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
TBLPR_ORDER = ['w:tblStyle','w:tblpPr','w:tblOverlap','w:bidiVisual','w:tblStyleRowBandSize',
 'w:tblStyleColBandSize','w:tblW','w:jc','w:tblCellSpacing','w:tblInd','w:tblBorders','w:shd',
 'w:tblLayout','w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange']

def _ordered_set(parent, tag, attrs, order):
    old = parent.find(qn(tag))
    if old is not None:
        parent.remove(old)
    e = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        e.set(qn(k), str(v))
    idx = order.index(tag)
    for child in parent:
        cf = 'w:' + etree.QName(child).localname
        if cf in order and order.index(cf) > idx:
            child.addprevious(e); return e
    parent.append(e); return e

def ppr_set(pPr, tag, attrs=None):  return _ordered_set(pPr, tag, attrs, PPR_ORDER)
def rpr_set(rPr, tag, attrs=None):  return _ordered_set(rPr, tag, attrs, RPR_ORDER)
def tblpr_set(tPr, tag, attrs=None): return _ordered_set(tPr, tag, attrs, TBLPR_ORDER)

def ppr_no_grid(pPr):
    """§1.1 关闭「如果定义了文档网格，则与网格对齐」（= Word 段落对话框里那个勾选，取消）。
    OOXML: w:pPr/w:snapToGrid w:val="0"。
    ⚠️ 必须显式写 val="0" —— 该元素 schema 默认值是 true(=勾选)，把元素删掉等于没做。
    ⚠️ 位置由 PPR_ORDER 保证（紧跟 w:adjustRightInd、在 w:spacing / w:ind 之前），乱序会让 Word 报损坏。
    ⚠️ 别和 run 级同名元素 w:rPr/w:snapToGrid（字符对字符网格）混淆，这里只动段落级那个。
    ⚠️ 任务书尤其重要：正文 1.25 倍（w:line=300）通常不是网格行高的整数倍，勾着网格对齐会被吸附，
       表现为"行距设了却没生效"。"""
    return ppr_set(pPr, 'w:snapToGrid', {'w:val': '0'})

def rpr_drop(rPr, tags):
    for t in tags:
        e = rPr.find(qn(t))
        if e is not None: rPr.remove(e)

def rpr_font(rPr, ea='宋体', latin='Times New Roman'):
    rpr_set(rPr, 'w:rFonts', {'w:ascii': latin, 'w:hAnsi': latin, 'w:eastAsia': ea, 'w:cs': latin})

def rpr_sz(rPr, val):
    rpr_set(rPr, 'w:sz', {'w:val': val}); rpr_set(rPr, 'w:szCs', {'w:val': val})

def get_pPr(p):
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr'); p.insert(0, pPr)
    return pPr

def run_rpr(r):
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr'); r.insert(0, rPr)
    return rPr

def para_rpr(pPr):
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        sect = pPr.find(qn('w:sectPr'))
        if sect is not None: sect.addprevious(rPr)
        else: pPr.append(rPr)
    return rPr

def mk_rpr_el(sz=24, bold=None):
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    for k, v in (('w:ascii','Times New Roman'),('w:hAnsi','Times New Roman'),
                 ('w:eastAsia','宋体'),('w:cs','Times New Roman')):
        rf.set(qn(k), v)
    rPr.append(rf)
    if bold is not None:
        b = OxmlElement('w:b'); b.set(qn('w:val'), '1' if bold else '0'); rPr.append(b)
        bc = OxmlElement('w:bCs'); bc.set(qn('w:val'), '1' if bold else '0'); rPr.append(bc)
    s = OxmlElement('w:sz'); s.set(qn('w:val'), str(sz)); rPr.append(s)
    sc = OxmlElement('w:szCs'); sc.set(qn('w:val'), str(sz)); rPr.append(sc)
    return rPr

def style_name_of(e):
    return Paragraph(e, doc).style.name or ''

# =====================================================================================
# STEP 0：定位边界（第一个"汉字序号一级标题"，且不是目录条目）
# =====================================================================================
children = list(body.iterchildren())
H1_RE = re.compile(r'^([一二三四五六七八九十]+)、(.+)$')
boundary = None
for i, ch in enumerate(children):
    if ch.tag != qn('w:p'): continue
    st = style_name_of(ch)
    if st.lower().startswith('toc'): continue
    t = (Paragraph(ch, doc).text or '').strip()
    m = H1_RE.match(t)
    if m and len(t) < 40 and not t.endswith(('。','；','，')):
        boundary = i; break
if boundary is None:
    sys.exit('未找到正文边界（第一个一级标题）')
log('边界 idx =', boundary, '::', Paragraph(children[boundary], doc).text.strip())
region = children[boundary:]

# =====================================================================================
# STEP 1：新建 4 个标题样式（宋体 小四 加粗；首行缩进 2 字符；两端对齐；outlineLvl 0-3）
# =====================================================================================
STYLES = doc.styles.element
HEAD_DEF = [('任务书一级标题', 'TBHeading1', '0'),
            ('任务书二级标题', 'TBHeading2', '1'),
            ('任务书三级标题', 'TBHeading3', '2'),
            ('任务书四级标题', 'TBHeading4', '3')]

def find_style_el(name):
    for st in STYLES.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        if nm is not None and nm.get(qn('w:val')) == name:
            return st
    return None

def ensure_heading_style(name, sid, outline):
    st = find_style_el(name)
    if st is None:
        st = OxmlElement('w:style')
        st.set(qn('w:type'), 'paragraph'); st.set(qn('w:styleId'), sid)
        nm = OxmlElement('w:name'); nm.set(qn('w:val'), name); st.append(nm)
        bs = OxmlElement('w:basedOn'); bs.set(qn('w:val'), 'Normal'); st.append(bs)
        ns = OxmlElement('w:next'); ns.set(qn('w:val'), 'Normal'); st.append(ns)
        STYLES.append(st)
    # pPr
    pPr = st.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr'); st.append(pPr)
    ppr_no_grid(pPr)   # §1.1：样式级兜底，之后新录入的段落自动继承
    _ordered_set(pPr, 'w:keepNext', None, PPR_ORDER)
    # 标题行距 = 1.5 倍（w:line=360）：§ 13 步骤 4.5 收紧到 1.25 的范围是"正文文字段落"，
    # 标题不参与分档（用户口径："正文 1.25（表格除外），其他部分与建议书一致"）。
    # 这里必须与 STEP 5 打底写入段落级的 360 保持一致 —— 曾误写 300，
    # 造成"样式说 1.25、段落实际 1.5"的定义/渲染不一致（见 § 2 行距说明）。
    _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
    _ordered_set(pPr, 'w:ind', {'w:firstLine': '480', 'w:firstLineChars': '200'}, PPR_ORDER)
    _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
    _ordered_set(pPr, 'w:outlineLvl', {'w:val': outline}, PPR_ORDER)
    # rPr
    rPr = st.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr'); st.append(rPr)
    rpr_font(rPr)
    rpr_set(rPr, 'w:b', None); rpr_set(rPr, 'w:bCs', None)
    rpr_sz(rPr, 24)
    return st.get(qn('w:styleId'))

HEAD_SID = {}
for nm, sid, ol in HEAD_DEF:
    HEAD_SID[nm] = ensure_heading_style(nm, sid, ol)

# §1.1 兜底：Normal 样式也写一份，让之后新录入的普通段落自动继承（段落级仍是主力）
def style_ppr(st):
    pPr = st.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        rPr = st.find(qn('w:rPr'))
        if rPr is not None: rPr.addprevious(pPr)
        else: st.append(pPr)
    return pPr
for st in STYLES.findall(qn('w:style')):
    _nm = st.find(qn('w:name'))
    if _nm is not None and (_nm.get(qn('w:val')) or '').lower() == 'normal':
        ppr_no_grid(style_ppr(st)); break

log('标题样式就绪:', HEAD_SID)
snap('01_styles')

# =====================================================================================
# STEP 2：异常字符/编号订正（2.2.5. 多了个点）
# =====================================================================================
for e in region:
    if e.tag != qn('w:p'): continue
    t = (Paragraph(e, doc).text or '').strip()
    if re.match(r'^\d+\.\d+\.\d+\.\D', t):     # 形如 2.2.5.xxx
        for r in e.iter(qn('w:r')):
            te = r.find(qn('w:t'))
            if te is not None and te.text and re.match(r'^\d+\.\d+\.\d+\.', te.text):
                te.text = re.sub(r'^(\d+\.\d+\.\d+)\.', r'\1', te.text)
                log('订正编号异常:', t[:26], '->', te.text[:26])
                break

# =====================================================================================
# STEP 3：标题识别与归位
# =====================================================================================
L4_RE = re.compile(r'^(\d+)\.(\d+)\.(\d+)\.(\d+)(?![\d.])')
L3_RE = re.compile(r'^(\d+)\.(\d+)\.(\d+)(?![\d.])')
L2_RE = re.compile(r'^(\d+)\.(\d+)(?![\d.])')
CAP_RE = re.compile(r'^(图|表)\s*\d')
END_RE = re.compile(r'[。；！？：]\s*$')

headings = {'1': [], '2': [], '3': [], '4': []}
for e in region:
    if e.tag != qn('w:p'): continue
    t = (Paragraph(e, doc).text or '').strip()
    if not t or len(t) > 60 or END_RE.search(t): continue
    if CAP_RE.match(t): continue
    if 'toc' in style_name_of(e).lower(): continue
    if H1_RE.match(t):
        headings['1'].append(e); continue
    if t == '附件':
        headings['1'].append(e); continue
    if L4_RE.match(t): headings['4'].append(e); continue
    if L3_RE.match(t): headings['3'].append(e); continue
    if L2_RE.match(t): headings['2'].append(e); continue

log('标题数量 L1=%d L2=%d L3=%d L4=%d' % tuple(len(headings[k]) for k in '1234'))
log('  L1:', [Paragraph(e, doc).text.strip()[:14] for e in headings['1']])

LVL_NAME = {'1': '任务书一级标题', '2': '任务书二级标题', '3': '任务书三级标题', '4': '任务书四级标题'}
LVL_OUT = {'1': '0', '2': '1', '3': '2', '4': '3'}
for lvl in '1234':
    sid = HEAD_SID[LVL_NAME[lvl]]
    for e in headings[lvl]:
        pPr = get_pPr(e)
        ppr_set(pPr, 'w:pStyle', {'w:val': sid})
        # 清理 run 级直接格式（否则覆盖样式定义）
        for r in e.iter(qn('w:r')):
            rPr = run_rpr(r)
            rpr_drop(rPr, ('w:sz','w:szCs','w:b','w:bCs','w:i','w:iCs','w:u','w:highlight'))
            rpr_font(rPr)
        # 段落级字体/字号（不设 b，让样式的加粗生效）
        prPr = para_rpr(pPr)
        rpr_drop(prPr, ('w:sz','w:szCs','w:b','w:bCs','w:i','w:iCs'))
        rpr_font(prPr); rpr_sz(prPr, 24)
        _ordered_set(pPr, 'w:spacing', {'w:line': '300', 'w:lineRule': 'auto'}, PPR_ORDER)
        _ordered_set(pPr, 'w:ind', {'w:firstLine': '480', 'w:firstLineChars': '200'}, PPR_ORDER)
        _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
        _ordered_set(pPr, 'w:outlineLvl', {'w:val': LVL_OUT[lvl]}, PPR_ORDER)
log('标题段落已归位 + run 级格式清理')
snap('03_headings')

# =====================================================================================
# STEP 4：清除一级标题的章节分页（V1.2 口径：任务书**不强制**分页）
#   规范：references/format-taskbook-sastind.md § 2「章节起始（任务书专属口径）」+ § 13 步骤 5.5
#   ⚠️ 与建议书**相反**：建议书是"前方没有分节符/分页符就把 pageBreakBefore 补成开启态"；
#      任务书是"**有就删掉**"——一级标题随正文自然流动，不要求每个章节各起新页。
#   ⚠️ 只删 `w:pageBreakBefore`，**绝不碰 `w:sectPr`**（分节符承担目录罗马/正文阿拉伯的页码体系）；
#      也**不要**删"与正文同段的手工分页符"（`w:br` 可能和实际文本混在同一个 `w:r` 里）。
# =====================================================================================
def is_pure_pagebreak_para(e):
    """是否为可安全删除的"纯分页符段落"：无文本、无绘图、无表格、自身不带 sectPr，且含手工分页符。
    只有这种段落才删；分页符若与正文混在同段则一律不动。"""
    if e.tag != qn('w:p'):
        return False
    pPr = e.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:sectPr')) is not None:
        return False                                    # 分节符段落：保留（页码体系靠它）
    if e.find('.//' + qn('w:drawing')) is not None:
        return False
    if e.find('.//' + qn('w:tbl')) is not None:
        return False
    if (Paragraph(e, doc).text or '').strip():
        return False
    return any(b.get(qn('w:type')) == 'page' for b in e.iter(qn('w:br')))

pb_removed, br_removed = [], []
for e in headings['1']:
    pPr = e.find(qn('w:pPr'))
    if pPr is not None:
        pbb = pPr.find(qn('w:pageBreakBefore'))
        if pbb is not None:
            # 开态（val 缺省/1/true）与关态（val="0"）**一律移除**，让段落属性干净
            # 注意：用 python-docx 的 paragraph_format.page_break_before = False
            #       只会写成 <w:pageBreakBefore w:val="0"/>，元素仍在——不够彻底。
            pb_removed.append('%s[%s]' % (Paragraph(e, doc).text.strip()[:16],
                                          pbb.get(qn('w:val')) or 'on'))
            pPr.remove(pbb)
    pv = e.getprevious()
    if pv is not None and is_pure_pagebreak_para(pv):
        br_removed.append(Paragraph(e, doc).text.strip()[:16])
        pv.getparent().remove(pv)
log('清除一级标题 w:pageBreakBefore %d 处:' % len(pb_removed), pb_removed)
log('删除一级标题前方"纯分页符段落" %d 个:' % len(br_removed), br_removed)
log('（校验：改完仍带 w:pageBreakBefore 的一级标题数应为 0，见 § 13.1 第 9 项）')
snap('04_no_pagebreak')

# =====================================================================================
# STEP 5：全文打底 1.5 倍行距 + 关闭"与文档网格对齐"（前置 + 正文 + 表格内 + 题注 + 目录条目）
#   §1.1：全文所有段落取消「如果定义了文档网格，则与网格对齐」。
#     - 必须显式写 val="0"（见 ppr_no_grid 注释：schema 默认值是 true，删元素 ≠ 取消）；
#     - body.iter 是**递归**的，表格 cell 内的段落也一并覆盖；
#     - 任务书正文 1.25 倍（w:line=300）通常不是网格行高的整数倍，不关掉会被吸附成 1 倍/1.5 倍，
#       表现为"行距设了却没生效" —— 这是任务书最常见的"行距改不动"原因。
# =====================================================================================
n = 0
for p in body.iter(qn('w:p')):
    ParagraphFormat(p).line_spacing = 1.5
    ppr_no_grid(get_pPr(p))
    n += 1
log('打底 1.5 倍行距段落数:', n, '（同时关闭网格对齐 —— 同一批段落）')

# =====================================================================================
# STEP 6：题注识别（§3.1 六条件）
# =====================================================================================
captions = []
for e in region:
    if e.tag != qn('w:p'): continue
    pp = Paragraph(e, doc); t = pp.text; ts = t.strip()
    if not ts: continue
    if CAP_RE.match(ts) and len(ts) < 80 and '\t' not in t and not END_RE.search(ts):
        captions.append(e)
log('题注数量:', len(captions), [Paragraph(e, doc).text.strip()[:20] for e in captions])

# =====================================================================================
# STEP 7：正文区 1.25 倍 + 宋体小四 + 首行缩进 2 字符 + 两端对齐
#        排除：标题、题注、表格内（顶层不是 w:p）、图片段落（居中不缩进）
# =====================================================================================
hset = set(id(x) for lv in headings.values() for x in lv)
cset = set(id(x) for x in captions)
body_targets = [e for e in region if e.tag == qn('w:p') and id(e) not in hset and id(e) not in cset]

img_paras = []
cnt_body = 0
for e in body_targets:
    pPr = get_pPr(e)
    has_img = e.find('.//' + qn('w:drawing')) is not None
    txt = (Paragraph(e, doc).text or '').strip()
    if has_img and not txt:
        img_paras.append(e)
        _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
        _ordered_set(pPr, 'w:jc', {'w:val': 'center'}, PPR_ORDER)
        ind = pPr.find(qn('w:ind'))
        if ind is not None: pPr.remove(ind)
        continue
    prPr = para_rpr(pPr)
    rpr_drop(prPr, ('w:sz','w:szCs','w:b','w:bCs'))
    rpr_font(prPr); rpr_sz(prPr, 24)
    for r in e.iter(qn('w:r')):
        rPr = run_rpr(r)
        rpr_font(rPr); rpr_sz(rPr, 24)
    _ordered_set(pPr, 'w:spacing', {'w:line': '300', 'w:lineRule': 'auto'}, PPR_ORDER)
    _ordered_set(pPr, 'w:ind', {'w:firstLine': '480', 'w:firstLineChars': '200'}, PPR_ORDER)
    _ordered_set(pPr, 'w:jc', {'w:val': 'both'}, PPR_ORDER)
    cnt_body += 1
log('正文段落(1.25/小四/缩进2/两端) 数:', cnt_body, ' 图片段落(居中):', len(img_paras))
snap('07_body')

# =====================================================================================
# STEP 8：表格（正文区顶层表）居中 + 宋体五号 + 1.5 倍行距
# =====================================================================================
tbls = [e for e in region if e.tag == qn('w:tbl')]
for e in tbls:
    tblPr = e.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); e.insert(0, tblPr)
    tblpr_set(tblPr, 'w:jc', {'w:val': 'center'})
    for p in e.iter(qn('w:p')):
        pPr = get_pPr(p)
        prPr = para_rpr(pPr)
        rpr_drop(prPr, ('w:sz','w:szCs'))
        rpr_font(prPr); rpr_sz(prPr, 21)
        for r in p.iter(qn('w:r')):
            rPr = run_rpr(r)
            rpr_font(rPr); rpr_sz(rPr, 21)
        _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
log('正文区表格处理数:', len(tbls))
snap('08_tables')

# =====================================================================================
# STEP 9：题注统一为 Caption 样式 + 宋体五号 + 不加粗 + 居中 + 1.5 + 去缩进
#         并按文档顺序重编号为连续 图1..n / 表1..n，构造成 SEQ 域（_Ref/_Toc 双书签）
# =====================================================================================
def _caption_style_el():
    # 内建样式在 XML 里 w:name 常为小写 "caption"，故做大小写无关匹配
    for st in STYLES.findall(qn('w:style')):
        nm = st.find(qn('w:name'))
        if nm is not None and (nm.get(qn('w:val')) or '').lower() == 'caption':
            return st
    return None

_capel = _caption_style_el()
cap_sid = _capel.get(qn('w:styleId')) if _capel is not None else None
# 兜底改 Caption 样式本身（宋体五号、不加粗）
if _capel is not None:
    _cr = _capel.find(qn('w:rPr'))
    if _cr is None:
        _cr = OxmlElement('w:rPr'); _capel.append(_cr)
    rpr_font(_cr); rpr_drop(_cr, ('w:b','w:bCs')); rpr_sz(_cr, 21)
log('Caption styleId =', cap_sid)

# 收集已用书签名 / 最大 id
existing_bm = set(b.get(qn('w:name')) for b in body.iter(qn('w:bookmarkStart')))
max_bmid = 0
for b in body.iter(qn('w:bookmarkStart')):
    try: max_bmid = max(max_bmid, int(b.get(qn('w:id'))))
    except Exception: pass
_bmc = [0]
def new_bm_name(prefix):
    while True:
        _bmc[0] += 1
        nm = '%s9%05d' % (prefix, _bmc[0])
        if nm not in existing_bm:
            existing_bm.add(nm); return nm

def mk_run(text, sz=21):
    r = OxmlElement('w:r'); r.append(mk_rpr_el(sz, bold=False))
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t); return r
def mk_fld(kind):
    r = OxmlElement('w:r'); r.append(mk_rpr_el(21, bold=False))
    fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), kind); r.append(fc); return r
def mk_instr(txt):
    r = OxmlElement('w:r'); r.append(mk_rpr_el(21, bold=False))
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = txt
    r.append(it); return r
def mk_bm_start(name):
    max_bmid_ = max_bmid + 1
    b = OxmlElement('w:bookmarkStart'); b.set(qn('w:id'), str(max_bmid_)); b.set(qn('w:name'), name)
    globals()['max_bmid'] = max_bmid_
    return b, max_bmid_
def mk_bm_end(bid):
    b = OxmlElement('w:bookmarkEnd'); b.set(qn('w:id'), str(bid)); return b

seq_counter = {'图': 0, '表': 0}
for e in captions:
    txt = (Paragraph(e, doc).text or '').strip()
    m = re.match(r'^(图|表)\s*\d+(?:[\.．]\d+)?\s*(.*)$', txt)
    label, desc = m.group(1), m.group(2).strip()
    seq_counter[label] += 1
    num = seq_counter[label]
    # 样式 + 段落格式
    pPr = get_pPr(e)
    if cap_sid: ppr_set(pPr, 'w:pStyle', {'w:val': cap_sid})
    prPr = para_rpr(pPr)
    rpr_drop(prPr, ('w:sz','w:szCs','w:b','w:bCs'))
    rpr_font(prPr); rpr_sz(prPr, 21)
    ind = pPr.find(qn('w:ind'))
    if ind is not None: pPr.remove(ind)
    _ordered_set(pPr, 'w:jc', {'w:val': 'center'}, PPR_ORDER)
    _ordered_set(pPr, 'w:spacing', {'w:line': '360', 'w:lineRule': 'auto'}, PPR_ORDER)
    # 重建 run 结构：label + SEQ 域（_Ref 书签）+ 说明（_Toc 书签）
    for ch in list(e):
        if ch is not pPr: e.remove(ch)
    refn, tocn = new_bm_name('_Ref'), new_bm_name('_Toc')
    bs1, bid1 = mk_bm_start(refn)
    e.append(bs1)
    e.append(mk_run(label))
    e.append(mk_fld('begin'))
    e.append(mk_instr(' SEQ %s \\* ARABIC ' % label))
    e.append(mk_fld('separate'))
    e.append(mk_run(str(num)))
    e.append(mk_fld('end'))
    e.append(mk_bm_end(bid1))
    bs2, bid2 = mk_bm_start(tocn)
    e.append(bs2)
    e.append(mk_run(' ' + desc))
    e.append(mk_bm_end(bid2))
    log('题注 -> %s%d %s   (原: %s)' % (label, num, desc, txt))
snap('09_captions')

# =====================================================================================
# STEP 10：主目录（保留 TOC 域，清掉过期的缓存条目，放占位提示）
# =====================================================================================
toc_start = None
for e in children[:boundary]:
    if e.tag != qn('w:p'): continue
    instrs = [(x.text or '').strip() for x in e.iter(qn('w:instrText'))]
    if any(s.startswith('TOC') for s in instrs):
        toc_start = e; break

if toc_start is None:
    log('!! 未找到主目录 TOC 指令段')
else:
    ci = children.index(toc_start)
    # 用"域深度"找 TOC 域真正闭合的段落（段内含嵌套 HYPERLINK/PAGEREF 域，不能见 end 就停）
    depth = 0; started = False; toc_end_idx = None
    for j in range(ci, len(children)):
        e = children[j]
        if e.tag != qn('w:p'): continue
        for r in e.iter(qn('w:r')):
            fc = r.find(qn('w:fldChar'))
            if fc is None: continue
            k = fc.get(qn('w:fldCharType'))
            if k == 'begin':
                depth += 1; started = True
            elif k == 'end':
                depth -= 1
                if started and depth <= 0:
                    toc_end_idx = j; break
        if toc_end_idx is not None: break
    # 清掉 TOC 域 separate 之后的缓存 run，放占位提示
    first_sep = None
    for r in toc_start.iter(qn('w:r')):
        fc = r.find(qn('w:fldChar'))
        if fc is not None and fc.get(qn('w:fldCharType')) == 'separate':
            first_sep = r; break
    if first_sep is not None:
        rm = False
        for ch in list(toc_start):
            if ch is first_sep:
                rm = True; continue
            if rm and ch.tag != qn('w:pPr'):
                toc_start.remove(ch)
        toc_start.append(mk_run('（目录：请在 WPS 中 Ctrl+A → F9 更新域后生成）', 24))
    # 删除 sep 与 end 之间的过期缓存段落（保留闭合段，它带分节符）
    removed = 0
    if toc_end_idx is not None:
        for e in children[ci+1:toc_end_idx]:
            if e.tag == qn('w:p'):
                e.getparent().remove(e); removed += 1
    log('主目录：清掉过期缓存段落 %d 个，改为占位提示（保留 TOC 域，闭合于 idx %s）' % (removed, toc_end_idx))
snap('10_toc')

# =====================================================================================
# STEP 11：保存
# =====================================================================================
doc.save(OUTPUT)
log('已保存:', OUTPUT)
