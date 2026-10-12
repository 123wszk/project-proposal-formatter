#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""离线诊断 DOCX 文档（不需要 editor_sdk），输出格式体检报告。

配合 project-proposal-formatter skill 的 Workflow 第 0 步（诊断）使用。
纯 python-docx 实现，可在文档未打开 / 无编辑器时先跑一遍摸底。

Usage:
    python diagnose_docx.py <path-to.docx> [--json]

Checks:
    1. 基本规模：段落数 / 表格数 / 内嵌图片数 / 字符数
    2. 样式清单：实际用到的段落样式 + 是否存在常见自定义标题样式
    3. 章节结构：汉字序号一级标题（一、…十五、）的连续性与漏号
    4. 标题层级：二级（N.N）/ 三级（N.N.N）标题数量
    5. 异常字符：控制字符（\\b \\u0005 等）所在段落
    6. 已有目录：TOC 字段 / 疑似目录条目
    7. 题注候选：形如「图N」「表N」开头的段落
    8. 行距统计：段落 line_spacing 分布
    9. 网格对齐：段落「如果定义了文档网格，则与网格对齐」开关覆盖情况
       （SASTIND 规范 § 1.1 要求全文取消，即段落级 w:snapToGrid w:val="0"；
        未关闭的段落行距会被吸附到 w:sectPr/w:docGrid 的网格行上，表现为"行距设了没生效"）
"""
import sys
import json
import re
from collections import Counter

try:
    from docx import Document
    from docx.shared import Pt
except ImportError:
    sys.exit("需要 python-docx：pip install python-docx")

# 汉字序号 -> 阿拉伯数字（覆盖 1..30，够 SASTIND 的"一…十五"用）
CN_NUM = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
    "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15,
    "十六": 16, "十七": 17, "十八": 18, "十九": 19, "二十": 20,
    "二十一": 21, "二十二": 22, "二十三": 23, "二十四": 24, "二十五": 25,
    "二十六": 26, "二十七": 27, "二十八": 28, "二十九": 29, "三十": 30,
}

# 需要优先报警的控制字符（除 \t \n \r 这些合法空白）
CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# 常见自定义标题样式名（各机构可能不同，这里给 SASTIND 一组 + 泛化关键词）
CUSTOM_STYLE_HINTS = ("建议书一级标题", "建议书二级标题", "建议书三级标题")
STYLE_KEYWORD_HINTS = ("标题", "Heading")


def _iter_paragraphs(doc):
    """产出 (scope, paragraph)：scope 为 'body' 或 'table'。"""
    for p in doc.paragraphs:
        yield "body", p
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield "table", p


def _heading_level_from_style(style_name):
    """从样式名推断 heading level（Heading 1 / 标题 1 / 建议书N级标题）。"""
    if not style_name:
        return 0
    m = re.search(r"(?:Heading|标题)\s*(\d)", style_name)
    if m:
        return int(m.group(1))
    return 0


def _is_heading_style(style_name):
    """样式名是否像"标题"样式（Heading / 标题 / XX一级标题…）。"""
    if not style_name:
        return False
    return bool(re.search(r"(Heading\s*\d|标题|heading)", style_name, re.I))


def diagnose(path):
    doc = Document(path)
    report = {}

    paras = list(_iter_paragraphs(doc))

    # ---- 1. 基本规模 ----
    body_paras = [p for s, p in paras if s == "body"]
    total_chars = sum(len(p.text) for _, p in paras)
    n_images = 0
    try:
        n_images = len(doc.inline_shapes)
    except Exception:
        pass
    report["scale"] = {
        "paragraphs_body": len(body_paras),
        "paragraphs_total": len(paras),
        "tables": len(doc.tables),
        "inline_images": n_images,
        "chars": total_chars,
    }

    # ---- 2. 样式清单 ----
    style_counter = Counter(
        (p.style.name if p.style is not None else "(None)") for _, p in paras
    )
    report["styles"] = {
        "used": dict(style_counter.most_common()),
        "has_sastind_title_styles": {
            name: (name in doc.styles) for name in CUSTOM_STYLE_HINTS
        },
    }

    # ---- 3. 章节结构（汉字序号一级标题 + 漏号）----
    # 关键：只有"汉字序号开头"且"用标题样式"的段落才算真章节，
    # 否则承诺书条款（一、二、三… 用 Normal 样式）会被误判。
    cn_chapters = []        # (order, num, text)
    cn_not_heading = []     # 形如章节但未套标题样式 —— 本身就是需要报的格式问题
    for _, p in paras:
        t = p.text.strip()
        m = re.match(r"^([一二三四五六七八九十]+)[、\.．]", t)
        if not (m and m.group(1) in CN_NUM):
            continue
        style_name = p.style.name if p.style is not None else ""
        if _is_heading_style(style_name):
            cn_chapters.append((CN_NUM[m.group(1)], t[:40]))
        else:
            cn_not_heading.append({"num": CN_NUM[m.group(1)], "style": style_name, "text": t[:40]})
    seen_nums = [n for n, _ in cn_chapters]
    missing = [n for n in range(1, (max(seen_nums) + 1) if seen_nums else 1) if n not in seen_nums]
    dupes = [n for n, c in Counter(seen_nums).items() if c > 1]
    # V3.6 漏号假阳性检测：文本无汉字序号但带 w:numPr 自动编号的标题样式段落
    # （Word 渲染时序号由编号引擎生成，p.text 读不到 —— 直接补文本会叠加成"七、七、X"）
    WNS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    numpr_note = []
    for _, p in paras:
        t = p.text.strip()
        if not t or len(t) > 40 or re.match(r"^[一二三四五六七八九十]+[、]", t):
            continue
        style_name = p.style.name if p.style is not None else ""
        if not _is_heading_style(style_name):
            continue
        pPr = p._p.find(WNS + "pPr")
        if pPr is not None and pPr.find(WNS + "numPr") is not None:
            numpr_note.append({"style": style_name, "text": t[:40]})
    report["chapters"] = {
        "count": len(cn_chapters),
        "order": [{"num": n, "text": t} for n, t in cn_chapters],
        "missing_nums": missing,
        "duplicate_nums": sorted(dupes),
        "cn_like_but_not_heading": cn_not_heading,
        "numpr_autonumber_candidates": numpr_note,
    }

    # ---- 4. 标题层级（N.N / N.N.N）----
    lvl2 = [p.text.strip() for p in body_paras
            if re.match(r"^\d+\.\d+\s", p.text.strip())]
    lvl3 = [p.text.strip() for p in body_paras
            if re.match(r"^\d+\.\d+\.\d+\s", p.text.strip())]
    report["heading_levels"] = {
        "level2_count": len(lvl2),
        "level3_count": len(lvl3),
        "level2_samples": lvl2[:5],
        "level3_samples": lvl3[:5],
    }

    # ---- 5. 异常字符（paragraph.text + raw XML 双扫）----
    anomalies = []
    for i, (scope, p) in enumerate(paras):
        bad = CTRL_RE.findall(p.text)
        if bad:
            codes = sorted({"U+%04X" % ord(c) for c in bad})
            anomalies.append({
                "index": i,
                "scope": scope,
                "codes": codes,
                "text": p.text[:50],
            })
    # raw XML：有些控制字符藏在字段 / 超链接 / 批注里，paragraph.text 读不到
    raw_xml = doc.element.body.xml
    xml_ctrl = sorted({"U+%04X" % ord(c) for c in CTRL_RE.findall(raw_xml)})
    report["control_chars"] = {
        "count": len(anomalies),
        "items": anomalies[:30],
        "xml_only_codes": xml_ctrl,
    }

    # ---- 6. 已有目录 ----
    xml = doc.element.body.xml
    has_toc_field = ("TOC \\" in xml) or ("fldSimple" in xml and "TOC" in xml)
    toc_like = [p.text.strip()[:60] for p in body_paras
                if re.search(r"(\.{3,}|…{2,})\s*\d+\s*$", p.text.strip())]
    report["existing_toc"] = {
        "has_toc_field": has_toc_field,
        "toc_like_paragraphs": len(toc_like),
        "samples": toc_like[:5],
    }

    # ---- 7. 题注候选 ----
    fig_caps = [p.text.strip()[:60] for _, p in paras
                if re.match(r"^图\s*\d+", p.text.strip())]
    tbl_caps = [p.text.strip()[:60] for _, p in paras
                if re.match(r"^表\s*\d+", p.text.strip())]
    report["captions"] = {
        "figure_caption_like": len(fig_caps),
        "table_caption_like": len(tbl_caps),
        "figure_samples": fig_caps[:5],
        "table_samples": tbl_caps[:5],
    }

    # ---- 8. 行距统计 ----
    ls_counter = Counter()
    for scope, p in paras:
        if scope != "body":
            continue
        ls = p.paragraph_format.line_spacing
        if ls is None:
            ls_counter["(未设置，继承)"] += 1
        elif isinstance(ls, float):
            ls_counter[f"{ls:g} 倍"] += 1
        else:
            # Length 对象：固定行距（pt）
            try:
                ls_counter[f"固定 {ls.pt:g}pt"] += 1
            except Exception:
                ls_counter[str(ls)] += 1
    report["line_spacing"] = dict(ls_counter.most_common())

    # ---- 9. 网格对齐（§1.1：全文应关闭「如果定义了文档网格，则与网格对齐」）----
    # OOXML: 段落级 w:pPr/w:snapToGrid w:val="0" 才是"取消勾选"；
    # 元素缺省 / val 为空 / val 为 1,true,on —— 都按"勾选"处理（schema 默认值是 true）。
    snap_off = 0
    snap_bad = []
    for scope, p in paras:
        pPr = p._p.find(WNS + "pPr")
        el = pPr.find(WNS + "snapToGrid") if pPr is not None else None
        val = el.get(WNS + "val") if el is not None else None
        if val in ("0", "false", "off"):
            snap_off += 1
        else:
            snap_bad.append({
                "scope": scope,
                "val": val if val is not None else "(缺省=勾选)",
                "text": (p.text or "").strip()[:40],
            })
    report["grid_alignment"] = {
        "paragraphs_total": len(paras),
        "snap_disabled": snap_off,
        "snap_enabled_or_unset": len(snap_bad),
        "items": snap_bad[:30],
    }

    return report


def render_text(rep):
    L = []
    s = rep["scale"]
    L.append("=" * 60)
    L.append("DOCX 格式体检报告")
    L.append("=" * 60)
    L.append(f"[规模] 段落 {s['paragraphs_total']}（正文段 {s['paragraphs_body']}）· "
             f"表格 {s['tables']} · 内嵌图 {s['inline_images']} · 字符 {s['chars']}")

    st = rep["styles"]
    L.append("\n[样式] 使用最多的段落样式 Top 12:")
    for name, cnt in list(st["used"].items())[:12]:
        L.append(f"    {cnt:>5}  {name}")
    L.append("    自定义标题样式是否存在（SASTIND 命名）:")
    for name, ok in st["has_sastind_title_styles"].items():
        L.append(f"      {'✅' if ok else '❌'}  {name}")

    ch = rep["chapters"]
    L.append(f"\n[章节] 汉字序号一级标题 {ch['count']} 个（已套标题样式）")
    if ch["missing_nums"]:
        L.append(f"    ⚠️  漏号: {ch['missing_nums']}（对照: 一=1 二=2 …）")
        if ch.get("numpr_autonumber_candidates"):
            L.append("    ⚠️  以下段落带 w:numPr 自动编号（渲染时序号由 Word 生成，p.text 读不到）——")
            L.append("        上述漏号可能是假阳性！补号前先移除 numPr 再手写序号，否则渲染成'七、七、X':")
            for it in ch["numpr_autonumber_candidates"]:
                L.append(f"        [{it['style']}] {it['text']}")
    if ch["duplicate_nums"]:
        L.append(f"    ⚠️  重号: {ch['duplicate_nums']}")
    if not ch["missing_nums"] and not ch["duplicate_nums"]:
        L.append("    ✅ 编号连续")
    for item in ch["order"]:
        L.append(f"      {item['num']:>2}. {item['text']}")
    if ch["cn_like_but_not_heading"]:
        L.append(f"    ℹ️  另有 {len(ch['cn_like_but_not_heading'])} 段形如章节但未套标题样式"
                 f"（可能是承诺书条款，也可能是漏套样式的真章节）:")
        for it in ch["cn_like_but_not_heading"][:10]:
            L.append(f"       [{it['style']}] {it['text']}")

    hl = rep["heading_levels"]
    L.append(f"\n[标题层级] 二级(N.N) {hl['level2_count']} 个 · 三级(N.N.N) {hl['level3_count']} 个")
    for t in hl["level2_samples"]:
        L.append(f"      二级样例: {t}")
    for t in hl["level3_samples"]:
        L.append(f"      三级样例: {t}")

    cc = rep["control_chars"]
    L.append(f"\n[异常字符] 命中段落 {cc['count']} 个")
    if cc["count"] == 0:
        L.append("    ✅ paragraph.text 中未发现控制字符")
    for it in cc["items"]:
        L.append(f"      idx={it['index']} [{it['scope']}] {it['codes']}  {it['text']!r}")
    if cc.get("xml_only_codes"):
        L.append(f"    ⚠️  raw XML 中另含控制字符 {cc['xml_only_codes']}"
                 f"（多在字段 / 超链接 / 批注内，paragraph.text 读不到；编辑器中可能表现为异常字符）")

    toc = rep["existing_toc"]
    L.append(f"\n[已有目录] TOC 字段: {'✅ 有' if toc['has_toc_field'] else '❌ 无'} · "
             f"目录样式段落 {toc['toc_like_paragraphs']} 个")
    for t in toc["samples"]:
        L.append(f"      {t}")

    cap = rep["captions"]
    L.append(f"\n[题注候选] 图题 {cap['figure_caption_like']} 个 · 表题 {cap['table_caption_like']} 个")
    for t in cap["figure_samples"]:
        L.append(f"      图: {t}")
    for t in cap["table_samples"]:
        L.append(f"      表: {t}")

    L.append(f"\n[行距] 正文段落 line_spacing 分布: {rep['line_spacing']}")

    ga = rep["grid_alignment"]
    L.append(f"\n[网格对齐] 段落 {ga['paragraphs_total']} 个 · 已关闭 {ga['snap_disabled']} 个 · "
             f"仍勾选/缺省 {ga['snap_enabled_or_unset']} 个（规范 §1.1 要求全部关闭）")
    if ga["snap_enabled_or_unset"] == 0:
        L.append("    ✅ 全部段落已取消「如果定义了文档网格，则与网格对齐」")
    else:
        L.append("    ⚠️  以下段落未关闭网格对齐（行距会被吸附到文档网格行，表现为\"行距设了没生效\"）:")
        for it in ga["items"]:
            L.append(f"      [{it['scope']}] val={it['val']}  {it['text']!r}")
    L.append("=" * 60)
    return "\n".join(L)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    path = sys.argv[1]
    as_json = "--json" in sys.argv[2:]
    rep = diagnose(path)
    if as_json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        print(render_text(rep))


if __name__ == "__main__":
    main()
