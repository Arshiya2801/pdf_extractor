import fitz  # PyMuPDF
import os
import json
import re
from collections import Counter, defaultdict
from statistics import mean

INPUT_DIR = r"C:/Users/sisod/Desktop/adobe/p1/input"
OUTPUT_DIR = r"C:/Users/sisod/Desktop/adobe/p1/output"

class TextElement:
    def __init__(self, text, font_size, font_name, is_bold, is_italic, x0, y0, w, h, page_num, block_no, line_no):
        self.text = text.strip()
        self.font_size = font_size
        self.font_name = font_name
        self.is_bold = is_bold
        self.is_italic = is_italic
        self.x0 = x0
        self.y0 = y0
        self.w = w
        self.h = h
        self.page_num = page_num  # physical page number (1-based from PyMuPDF)
        self.block_no = block_no
        self.line_no = line_no

    def __repr__(self):
        return f"TextElement(text={self.text!r}, fs={self.font_size}, b={self.is_bold}, i={self.is_italic}, x={self.x0}, y={self.y0}, pg={self.page_num}, block={self.block_no}, line={self.line_no})"

def extract_text_elements(doc):
    """
    Extracts styled lines as TextElement (font, position, block/line grouping info).
    """
    text_elements = []
    for pg in range(len(doc)):
        page = doc[pg]
        blocks = page.get_text("dict")["blocks"]
        for block_no, b in enumerate(blocks):
            if b['type'] != 0:
                continue
            for line_no, line in enumerate(b["lines"]):
                # Aggregate spans within a line by their font/size/style, merging only if they are the same
                line_text, font_size, font_name, is_bold, is_italic, x0, y0, w, h = [], None, None, False, False, None, None, None, None
                for span in line["spans"]:
                    t = span["text"].strip()
                    if not t:
                        continue
                    # Take the dominant attributes from the first non-empty span in the line
                    if font_size is None:
                        font_size = span["size"]
                        font_name = span["font"]
                        is_bold = "Bold" in font_name or "bold" in font_name.lower()
                        is_italic = "Italic" in font_name or "Oblique" in font_name or "italic" in font_name.lower()
                        x0 = span["bbox"][0]
                        y0 = span["bbox"][1]
                        w = span["bbox"][2]-span["bbox"][0]
                        h = span["bbox"][3]-span["bbox"][1]
                    line_text.append(t)
                if line_text:
                    txt = " ".join(line_text)
                    el = TextElement(txt, font_size, font_name, is_bold, is_italic, x0, y0, w, h, pg + 1, block_no, line_no)
                    text_elements.append(el)
    return text_elements

def cluster_font_sizes(text_elements, tolerance=0.5):
    """
    Cluster unique font sizes, returns cluster means sorted descending.
    """
    sizes = sorted(set(round(el.font_size, 2) for el in text_elements if el.font_size), reverse=True)
    clusters = []
    while sizes:
        base = sizes[0]
        cluster = [s for s in sizes if abs(s - base) <= tolerance]
        clusters.append(mean(cluster))
        for s in cluster:
            sizes.remove(s)
    return sorted(clusters, reverse=True)

def remove_header_footer(text_elements, per_page, y_top=50, y_bottom=750):
    """
    Remove elements likely belonging to recurring header/footer.
    """
    num_pages = max(el.page_num for el in text_elements)
    headers = defaultdict(int)
    footers = defaultdict(int)
    for el in text_elements:
        if el.y0 < y_top:
            headers[el.text] += 1
        if el.y0 > y_bottom:
            footers[el.text] += 1
    header_texts = {k for k, v in headers.items() if v > 0.7 * num_pages}
    footer_texts = {k for k, v in footers.items() if v > 0.7 * num_pages}
    return [
        el for el in text_elements
        if not (el.text in header_texts and el.y0 < y_top) \
        and not (el.text in footer_texts and el.y0 > y_bottom)
    ]

def merge_multiline_blocks(text_elements, tolerance=2):
    """
    Merge consecutive lines with the same font/size/style/indent/block, forming multi-line titles/headings.
    """
    merged = []
    buf = None
    for el in sorted(text_elements, key=lambda x: (x.page_num, x.block_no, x.line_no)):
        if buf is None:
            buf = el
            continue
        # If exactly same font, size, bold, italic and block and x0 are very close
        # and vertically adjacent, treat as same entity (merge text)
        close_same = (
            el.page_num == buf.page_num and
            el.block_no == buf.block_no and
            abs(el.font_size - buf.font_size) < 0.2 and
            el.font_name == buf.font_name and
            el.is_bold == buf.is_bold and
            el.is_italic == buf.is_italic and
            abs(el.x0 - buf.x0) < 2 and
            0 < (el.y0 - (buf.y0 + buf.h)) < 15  # within 15 pts vertically stacked
        )
        if close_same:
            buf.text += " " + el.text
            buf.h += el.h  # extend block height
        else:
            merged.append(buf)
            buf = el
    if buf:
        merged.append(buf)
    return merged

def detect_title(text_elements, single_page):
    """
    Combine all largest-size lines at top/center on cover (or single page) as title (multi-line).
    """
    page_num = 1
    page_elems = [el for el in text_elements if el.page_num == page_num]
    if not page_elems:
        return "", None
    # Find largest font size
    max_size = max(el.font_size for el in page_elems)
    candidates = [el for el in page_elems if abs(el.font_size - max_size) < 0.2]
    if not candidates:
        return "", None
    # Collect all vertically grouped/collapsed candidates at top~center
    top_candidates = sorted([el for el in candidates if el.y0 < 250], key=lambda x: x.y0)
    title_lines = [el.text for el in top_candidates]
    title = " ".join(title_lines).strip()
    if not title and single_page:
        # Fallback: assign largest font text(s) as H1, not title
        return "", max_size
    return title, max_size

def find_numbering_level(text):
    text = text.strip()
    m = re.match(r'^(\d+(\.\d+){0,2})[ \t\-:]+', text)
    if m:
        level = m.group(1).count('.') + 1
        if level > 3:
            level = 3
        return level
    return None

def assign_headings(text_elements, font_clusters, title_font_size, title_lines, page_offset, is_single_page):
    """
    Assign heading levels, robustly handling page offset, multi-line, missing title, number patterns.
    """
    # Exclude title from candidates
    heading_clusters = [fs for fs in font_clusters if not (title_font_size and abs(fs-title_font_size) < 0.3)]
    levels = {}
    idx = 0
    for tag in ["H1", "H2", "H3"]:
        if idx < len(heading_clusters):
            levels[heading_clusters[idx]] = tag
            idx += 1

    outline = []
    prev_order = {"H1":1, "H2":2, "H3":3}
    last_level = 0

    # Sort for stable processing order
    text_elements = sorted(text_elements, key=lambda x: (x.page_num, x.y0, x.x0))
    for el in text_elements:
        page_idx = (el.page_num-1) if not is_single_page else 0
        # For multi-page docs, treat cover as page 0, others page_offset=1
        logical_page = (page_idx) - page_offset
        if is_single_page:
            logical_page = 0
        if logical_page < 0:
            continue  # skip processing on cover

        # Skip title lines
        if el.page_num == 1 and title_lines and el.text in title_lines:
            continue

        # Skip short/junk text
        if len(el.text) < 2:
            continue
        # Skip repeated header/footer (should be already removed)
        # Match against body: want standalone, not inline
        candidates = [(fs, abs(el.font_size - fs)) for fs in levels]
        if not candidates:
            continue
        best_cluster = min(candidates, key=lambda t: t[1])
        cluster_val, delta = best_cluster
        if delta > 0.3:  # Not sufficiently matching any cluster
            continue
        level = levels[cluster_val]

        # Numbering validation: allow numbering to override heuristic level
        num_level = find_numbering_level(el.text)
        if num_level is not None:
            if num_level == 1:
                level = "H1"
            elif num_level == 2:
                level = "H2"
            elif num_level == 3:
                level = "H3"
        # Skip if ends in fullstop
        if el.text.rstrip().endswith('.'):
            continue
        # (Optional) skip if looks like a full sentence
        if el.text.split() and len(el.text.split()) > 12:
            continue
        # Do not allow H2/H3 before H1 on a new page (structure order)
        if prev_order[level] < last_level:
            continue
        last_level = prev_order[level]
        # Use vertical whitespace as cue (TODO: can use if we wish)

        outline.append({
            "level": level,
            "text": el.text,
            "page": int(logical_page)+1  # JSON output page 1-based (after offset)
        })
    return outline

def process_pdf_file(filepath):
    doc = fitz.open(filepath)
    num_pages = doc.page_count
    single_page = num_pages == 1

    # Step 1: Extract all styled lines
    raw_elements = extract_text_elements(doc)
    per_page = defaultdict(list); [per_page[el.page_num].append(el) for el in raw_elements]

    # Step 2: Remove likely headers/footers
    cleaned = remove_header_footer(raw_elements, per_page)

    # Step 3: Merge multi-line blocks
    elements_merged = merge_multiline_blocks(cleaned)

    # Step 4: Cluster font sizes
    font_clusters = cluster_font_sizes(elements_merged)

    # Step 5: Detect title (smart multi-line, center/top, fallback logic)
    title, title_size = detect_title(elements_merged, single_page)
    title_lines = [l for l in elements_merged if abs(l.font_size-title_size)<0.2 and l.page_num==1] if title_size else []

    # Step 6: Decide page offsets for logical numbering (cover = 0)
    page_offset = 1 if not single_page else 0

    # Step 7: Assign hierarchical headings using robust mapping
    outline = assign_headings(elements_merged, font_clusters, title_size, [el.text for el in title_lines], page_offset, single_page)

    # Final output
    return {
        "title": title,
        "outline": outline
    }

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    pdf_files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".pdf")]
    for pdf_file in pdf_files:
        file_path = os.path.join(INPUT_DIR, pdf_file)
        print(f"Processing {pdf_file}...")
        outline_data = process_pdf_file(file_path)
        base_name = os.path.splitext(pdf_file)[0]
        output_path = os.path.join(OUTPUT_DIR, f"{base_name}.json")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(outline_data, f, indent=2, ensure_ascii=False)
        print(f"Output written to {output_path}")

if __name__ == "__main__":
    main()
