import fitz  # PyMuPDF
import os
import json
import re
from collections import defaultdict, Counter
from statistics import mean

INPUT_DIR = "C:/Users/sisod/Desktop/adobe/p1/input"
OUTPUT_DIR = "C:/Users/sisod/Desktop/adobe/p1/output"

# Helper dataclass for text elements with style
class TextElement:
    def __init__(self, text, font_size, font_name, is_bold, is_italic, x0, y0, page_num):
        self.text = text.strip()
        self.font_size = font_size
        self.font_name = font_name
        self.is_bold = is_bold
        self.is_italic = is_italic
        self.x0 = x0  # horizontal position for indentation
        self.y0 = y0  # vertical position
        self.page_num = page_num

    def __repr__(self):
        return f"TextElement(text={self.text}, size={self.font_size}, bold={self.is_bold}, italic={self.is_italic}, x0={self.x0}, y0={self.y0}, page={self.page_num})"


def extract_text_elements(doc):
    """
    Extract text elements (lines with styles, font info, position) from all pages.
    Returns list of TextElement.
    """
    text_elements = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if b['type'] != 0:  # skip images etc.
                continue
            for line in b["lines"]:
                spans = line["spans"]
                # To handle same line with multiple spans and different fonts/sizes,
                # treat each span as a separate element for now.
                for span in spans:
                    text = span["text"]
                    if not text.strip():
                        continue
                    font_size = span["size"]
                    font_name = span["font"]
                    is_bold = "Bold" in font_name or "bold" in font_name.lower()
                    is_italic = "Italic" in font_name or "Oblique" in font_name or "italic" in font_name.lower()
                    x0 = span["bbox"][0]
                    y0 = span["bbox"][1]
                    el = TextElement(text, font_size, font_name, is_bold, is_italic, x0, y0, page_num + 1)
                    text_elements.append(el)
    return text_elements

def cluster_font_sizes(text_elements, tolerance=0.5):
    """
    Cluster font sizes into groups (to identify unique font sizes representing Title, H1, H2, H3)
    Returns sorted list of unique sizes (largest to smallest).
    """
    sizes = sorted(set([round(el.font_size, 2) for el in text_elements]), reverse=True)
    clusters = []
    while sizes:
        base = sizes[0]
        cluster = [s for s in sizes if abs(s - base) <= tolerance]
        clusters.append(mean(cluster))
        for s in cluster:
            sizes.remove(s)
    clusters.sort(reverse=True)
    return clusters

def detect_title(text_elements):
    """
    Heuristic: Title is largest font on page 1, near top or center.
    If multiple candidates, pick the earliest (smallest y0) and/or centered text.
    Returns title text string.
    """
    first_page_elements = [el for el in text_elements if el.page_num == 1]
    if not first_page_elements:
        return ""
    # Get largest font size on page 1
    max_font_size = max(el.font_size for el in first_page_elements)
    candidates = [el for el in first_page_elements if abs(el.font_size - max_font_size) < 0.2]

    # Among candidates pick one near top (min y0) and near center (x0 around page center)
    # Approximate page width for centering
    page_width = 595  # A4 approx width in points; adjust if you want to use actual doc.page.rect[2]
    def score(el):
        # Lower score = better; y distance and distance from center weighted
        center_x = page_width / 2
        return el.y0 + abs(el.x0 - center_x)*0.5
    candidates.sort(key=score)
    title = candidates[0].text
    return title

def remove_header_footer(text_elements, y_threshold_top=50, y_threshold_bottom=750, page_height=842):
    """
    Remove elements likely to be in header/footer based on vertical position and repetitive content.
    We judge header around top 50 pts, footer around bottom 92 pts (page height - approx)
    """
    filtered_elements = []
    # Identify frequent recurring texts at top or bottom pos (possible headers/footers)
    top_texts = [el.text for el in text_elements if el.y0 < y_threshold_top]
    bottom_texts = [el.text for el in text_elements if el.y0 > y_threshold_bottom]
    top_freq = Counter(top_texts)
    bottom_freq = Counter(bottom_texts)
    # Threshold: texts appearing on >70% pages in header/footer area likely headers/footers
    num_pages = len(set(el.page_num for el in text_elements))
    header_candidates = set([txt for txt, cnt in top_freq.items() if cnt > 0.7*num_pages])
    footer_candidates = set([txt for txt, cnt in bottom_freq.items() if cnt > 0.7*num_pages])
    for el in text_elements:
        if el.y0 < y_threshold_top and el.text in header_candidates:
            continue  # Skip header
        if el.y0 > y_threshold_bottom and el.text in footer_candidates:
            continue  # Skip footer
        filtered_elements.append(el)
    return filtered_elements

def find_numbering_level(text):
    """
    Regex to detect numbering patterns like '1.', '1.2', '2.1.3' indicating heading level depth.
    Returns level (1,2,3) based on number of segments.
    """
    text = text.strip()
    # Match patterns like '1.', '2.1', '3.1.4 '
    match = re.match(r'^(\d+(\.\d+){0,2})[ \t\-:]*', text)
    if match:
        numbering = match.group(1)
        level = numbering.count(".") + 1  # count dots + 1, e.g. '1' -> level 1, '1.2' -> level 2
        if level > 3:
            level = 3  # limit to H3 max
        return level
    return None

def assign_headings(text_elements, font_clusters, title_font_size):
    """
    Assign heading levels based on font size clusters & heuristics.
    - Title font size given separately.
    - Next largest = H1, then H2, then H3 (font size in clusters).
    Apply page rules: skip page 1 for headings (except title).
    Use numbering and indentation as additional signals.
    """

    # Map font size cluster to level, ignoring title font size
    heading_font_sizes = [fs for fs in font_clusters if abs(fs - title_font_size) > 0.3]
    heading_font_sizes.sort(reverse=True)  # Larger -> higher level

    font_size_to_level = {}
    if heading_font_sizes:
        font_size_to_level[heading_font_sizes[0]] = "H1"
    if len(heading_font_sizes) > 1:
        font_size_to_level[heading_font_sizes[1]] = "H2"
    if len(heading_font_sizes) > 2:
        font_size_to_level[heading_font_sizes[2]] = "H3"

    outline = []
    prev_level_order = {"H1": 1, "H2": 2, "H3": 3}
    last_heading_level = 0  # For structural validation

    # Sort elements by page, then vertical position
    text_elements = sorted(text_elements, key=lambda x: (x.page_num, x.y0, x.x0))

    for el in text_elements:
        # Skip page 1 elements except title
        if el.page_num == 1:
            continue

        # Ignore empty or very short text
        if len(el.text) < 2:
            continue

        # Skip text likely to be body based on very small fonts or non-heading fonts
        # Use clustering with tolerance
        clusters_diffs = [abs(el.font_size - fs) for fs in font_clusters]
        min_diff = min(clusters_diffs) if clusters_diffs else 1000
        closest_cluster = font_clusters[clusters_diffs.index(min_diff)] if clusters_diffs else None

        if closest_cluster not in font_size_to_level:
            continue

        # Assign heading level by mapped font size
        level = font_size_to_level[closest_cluster]

        # Numbering pattern check: if numbering indicates different level, adjust level accordingly
        num_level = find_numbering_level(el.text)
        if num_level is not None:
            # Upgrade/downgrade level based on numbering detected
            # e.g., if numbering level mismatch font size level, prefer numbering hierarchy
            if num_level == 1:
                level = "H1"
            elif num_level == 2:
                level = "H2"
            elif num_level >= 3:
                level = "H3"

        # Structural validation of heading levels
        current_level_order = prev_level_order[level]
        if current_level_order < last_heading_level:
            # For example, we should not jump back to H2 after H3 without H1
            # But allow going deeper or same level
            # To keep it simple, accept it but you could choose to skip or adjust
            pass
        last_heading_level = current_level_order

        # Check whitespace above (y0 difference with previous element on same page)
        # Helps confirm heading is visually separated
        # (Here skipped for simplicity but can be added if access to previous element Y positions)

        # Check indentation difference - simple (if x0 is bigger than typical H1, then likely H2 etc.)
        # Could be added later for refinement

        # Semantic check - ignore if text ends with full stop (likely paragraph)
        if el.text.endswith('.'):
            continue

        # Confidence increase - numberings, bold or italic styles support heading classification
        # (Implicit here)

        outline.append({
            "level": level,
            "text": el.text,
            "page": el.page_num
        })

    return outline

def process_pdf_file(filepath):
    """
    Process single PDF file; extract title and outline.
    """
    doc = fitz.open(filepath)
    text_els = extract_text_elements(doc)
    # Remove headers and footers
    text_els = remove_header_footer(text_els)

    # Detect title
    title = detect_title(text_els)

    # Cluster font sizes to determine font hierarchy
    font_clusters = cluster_font_sizes(text_els)

    # Assign headings based on font clusters and heuristics
    outline = assign_headings(text_els, font_clusters, title_font_size=max(font_clusters) if font_clusters else 0)

    result = {
        "title": title,
        "outline": outline
    }
    return result

def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    pdf_files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".pdf")]

    for pdf_file in pdf_files:
        file_path = os.path.join(INPUT_DIR, pdf_file)
        print(f"Processing {pdf_file}...")
        outline_data = process_pdf_file(file_path)

        # Compose output JSON filename
        base_name = os.path.splitext(pdf_file)[0]
        output_path = os.path.join(OUTPUT_DIR, f"{base_name}.json")

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(outline_data, f, indent=2, ensure_ascii=False)

        print(f"Output written to {output_path}")

if __name__ == "__main__":
    main()
