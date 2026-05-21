#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Extract text, references, chapter structure, and citations from a Chinese thesis .docx file."""

import docx
import re
import sys

FILE_PATH = "/Users/shiberlin/Desktop/毕业论文/毕业论文_修改版.docx"
OUTPUT_PATH = "/Users/shiberlin/Desktop/毕业论文/thesis_extraction_result.txt"

doc = docx.Document(FILE_PATH)

lines = []
def out(s=""):
    lines.append(s)

# ============================================================
# Part 1: Extract ALL paragraph text with index positions
# ============================================================
out("=" * 80)
out("PART 1: ALL NON-EMPTY PARAGRAPHS WITH INDEX POSITIONS")
out("=" * 80)

paragraphs_info = []
for i, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    style_name = para.style.name if para.style else "None"
    has_superscript = any(run.font.superscript for run in para.runs if run.font.superscript)
    paragraphs_info.append({
        "index": i,
        "text": text,
        "style": style_name,
        "has_superscript": has_superscript,
        "text_length": len(text),
    })

for p in paragraphs_info:
    if p["text_length"] > 0:
        marker = " [SUPERSCRIPT]" if p["has_superscript"] else ""
        out(f"[{p['index']:4d}] ({p['style']:30s}){marker} {p['text'][:200]}")

non_empty = sum(1 for p in paragraphs_info if p["text_length"] > 0)
out(f"\nTotal paragraphs: {len(paragraphs_info)}, Non-empty: {non_empty}")

# ============================================================
# Part 2: Identify chapter structure
# ============================================================
out("\n" + "=" * 80)
out("PART 2: CHAPTER STRUCTURE (参考文献 / 章节)")
out("=" * 80)

chapter_patterns = [
    (r"^参考文献$", "参考文献"),
    (r"^第[一二三四五六七八九十]+章", "章节标题"),
    (r"^附录", "附录"),
    (r"^摘要", "摘要"),
    (r"^Abstract", "Abstract"),
    (r"^致谢", "致谢"),
    (r"^\d+\.\d+\s+\S", "节标题"),
]

chapters = []
ref_start_idx = None

for p in paragraphs_info:
    if p["text_length"] == 0:
        continue
    for pattern, label in chapter_patterns:
        if re.search(pattern, p["text"]):
            chapters.append({
                "index": p["index"],
                "label": label,
                "text": p["text"],
            })
            if label == "参考文献":
                ref_start_idx = p["index"]
            break

for ch in chapters:
    out(f"  [{ch['index']:4d}] {ch['label']:8s} — {ch['text']}")

# ============================================================
# Part 3: Extract references
# ============================================================
out("\n" + "=" * 80)
out("PART 3: REFERENCE LIST (参考文献)")
out("=" * 80)

references = []
if ref_start_idx is not None:
    for p in paragraphs_info[ref_start_idx + 1:]:
        text = p["text"]
        if not text:
            continue
        # Stop at next major section
        if re.match(r"^(附录|致谢|结论|第[一二三四五六七八九十]+章)", text):
            break
        references.append(text)

    out(f"Found {len(references)} references:\n")
    for i, ref in enumerate(references, 1):
        out(f"  [{i}] {ref}")
else:
    out("WARNING: '参考文献' section not found! Searching all paragraphs...")
    for p in paragraphs_info:
        if re.match(r"^\[\d+\]", p["text"]):
            references.append(p["text"])
    if references:
        out(f"Found {len(references)} references by [N] pattern:")
        for ref in references:
            out(f"  {ref}")
    else:
        out("No references found.")

# ============================================================
# Part 4: Citation analysis
# ============================================================
out("\n" + "=" * 80)
out("PART 4: CITATION ANALYSIS (引文标注)")
out("=" * 80)

citation_pattern = re.compile(r"\[\d+([,\-—]\d+)*\]")

citations_found = []
for p in paragraphs_info:
    if p["text_length"] == 0:
        continue
    matches = citation_pattern.findall(p["text"])
    if matches:
        belonging_chapter = "未知"
        for ch in chapters:
            if ch["index"] <= p["index"]:
                belonging_chapter = ch["text"]
        citations_found.append({
            "para_index": p["index"],
            "chapter": belonging_chapter,
            "citations": matches,
            "snippet": p["text"][:200],
        })

out(f"Total paragraphs containing bracket citations: {len(citations_found)}")
out(f"Total citation instances: {sum(len(c['citations']) for c in citations_found)}")

all_cite_nums = set()
for c in citations_found:
    for cite in c["citations"]:
        nums = re.findall(r"\d+", cite)
        all_cite_nums.update(int(n) for n in nums)

out(f"Unique citation numbers referenced: {sorted(all_cite_nums)}")

# Superscript analysis
superscript_paras = [p for p in paragraphs_info if p["has_superscript"] and p["text_length"] > 0]
out(f"\nParagraphs with superscript runs: {len(superscript_paras)}")
for p in superscript_paras[:50]:
    super_texts = [run.text for run in doc.paragraphs[p["index"]].runs if run.font.superscript and run.text.strip()]
    if super_texts:
        out(f"  [{p['index']:4d}] Superscript content: {super_texts} | Context: {p['text'][:120]}")

if citations_found:
    out("\n--- All paragraphs with citations ---")
    for c in citations_found:
        out(f"  [{c['para_index']:4d}] ({c['chapter'][:15]}) Citations: {c['citations']}")
        out(f"          Text: {c['snippet']}")
        out()

# ============================================================
# Part 5: Chapter-to-paragraph-range mapping
# ============================================================
out("=" * 80)
out("PART 5: CHAPTER-TO-PARAGRAPH MAPPING")
out("=" * 80)

for idx, ch in enumerate(chapters):
    start = ch["index"]
    end = chapters[idx + 1]["index"] if idx + 1 < len(chapters) else len(paragraphs_info)
    content_paras = [p for p in paragraphs_info if start < p["index"] < end and p["text_length"] > 0]
    cite_count = sum(1 for c in citations_found if start < c["para_index"] < end)
    out(f"  {ch['text']}")
    out(f"    Paragraph range: [{start+1} - {end-1}]")
    out(f"    Content paragraphs: {len(content_paras)}")
    out(f"    Paragraphs with citations: {cite_count}")
    out()

# ============================================================
# SUMMARY
# ============================================================
out("=" * 80)
out("SUMMARY")
out("=" * 80)
out(f"  Total paragraphs in document: {len(paragraphs_info)}")
out(f"  Non-empty paragraphs: {non_empty}")
out(f"  Chapters/sections found: {len(chapters)}")
out(f"  References extracted: {len(references)}")
out(f"  Paragraphs with bracket citations: {len(citations_found)}")
out(f"  Paragraphs with superscript runs: {len(superscript_paras)}")
out(f"  Unique citation numbers: {sorted(all_cite_nums) if all_cite_nums else 'None'}")

if references:
    out(f"\n  REFERENCE NUMBER RANGE: [{min(len(references), 1)}..{len(references)}]")
    out(f"  CITATION NUMBER RANGE:   [{min(all_cite_nums) if all_cite_nums else 'N/A'}..{max(all_cite_nums) if all_cite_nums else 'N/A'}]")
    if all_cite_nums:
        missing = set(range(1, max(all_cite_nums)+1)) - all_cite_nums
        if missing:
            out(f"  Citations NOT matching references: {sorted(missing)}")

# Write to file
with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Results written to: {OUTPUT_PATH}")
print(f"Total lines: {len(lines)}")
