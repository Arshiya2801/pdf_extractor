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

def cluster_font_sizes(text_elements, tolerance=0.3):
    """
    Cluster font sizes into groups (to identify unique font sizes representing Title, H1, H2, H3)
    Returns sorted list of unique sizes (largest to smallest).
    """
    sizes = sorted(set([round(el.font_size, 1) for el in text_elements]), reverse=True)
    clusters = []
    while sizes:
        base = sizes[0]
        cluster = [s for s in sizes if abs(s - base) <= tolerance]
        clusters.append(mean(cluster))
        for s in cluster:
            sizes.remove(s)
    clusters.sort(reverse=True)
    return clusters


import re
from collections import defaultdict

def normalize_header_footer(text):
    text = re.sub(r'Page \d+ of \d+', '', text)
    text = re.sub(r'Version\s?\d+', '', text)
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.lower().strip()

def find_repeated_lines(elements, total_pages, y_top=80, y_bottom=770, min_frac=0.6):
    header_footer = defaultdict(set)
    for el in elements:
        norm = normalize_header_footer(el.text)
        if (el.y0 < y_top or el.y1 > y_bottom) and len(norm) > 0:
            header_footer[norm].add(el.page_num)
    return {txt for txt, pages in header_footer.items() if len(pages)/total_pages > min_frac}

def filter_outline_repeats(outline, repeated_set):
    # Remove heading if normalized text matches a known header/footer
    filtered = []
    for h in outline:
        norm = normalize_header_footer(h['text'])
        if norm in repeated_set:
            continue
        filtered.append(h)
    return filtered

# --- in your main process_pdf_file flow ---
# After you extract elements and before you assign headings:

# 1. Find repetitive header/footer lines
repeated_set = find_repeated_lines(elements, total_pages=doc.page_count)

# 2. [after heading assignment]
outline = assign_headings(merged_elements, font_clusters, title_font_size, title_lines_texts,
                          page_offset, single_page, avg_body_gap)

# 3. Remove repeated header/footer artifacts from outline
outline = filter_outline_repeats(outline, repeated_set)

# The rest of your output logic continues as before.


def detect_title(text_elements):
    """
    Detect title from page 1 elements, combining multiple elements if they form the title.
    """
    first_page_elements = [el for el in text_elements if el.page_num == 1]
    if not first_page_elements:
        return ""
    
    # Look for elements that could be part of the title
    # Title elements are typically larger fonts and positioned near the top
    max_font_size = max(el.font_size for el in first_page_elements)
    title_candidates = [el for el in first_page_elements if abs(el.font_size - max_font_size) < 1.0]
    
    # Sort by vertical position (y0) to get elements in order
    title_candidates.sort(key=lambda x: x.y0)
    
    # Take the first few elements that could form the title
    # Filter out elements that are clearly not title (like page numbers, footers)
    title_parts = []
    for el in title_candidates[:5]:  # Look at first 5 candidates
        # Skip elements that look like page numbers or copyright
        if re.match(r'^\d+$', el.text) or 'copyright' in el.text.lower() or 'page' in el.text.lower():
            continue
        if len(el.text) > 2:  # Minimum length for title parts
            title_parts.append(el.text)
    
    # Combine title parts
    title = "  ".join(title_parts) if title_parts else (title_candidates[0].text if title_candidates else "")
    return title

def remove_header_footer(text_elements, y_threshold_top=80, y_threshold_bottom=750):
    """
    Remove elements likely to be in header/footer based on repetitive content.
    More conservative approach - only remove clearly repetitive elements.
    """
    filtered_elements = []
    
    # Count text occurrences across pages
    text_counts = Counter([el.text for el in text_elements])
    num_pages = len(set(el.page_num for el in text_elements))
    
    # Only remove text that appears on most pages and looks like header/footer
    repetitive_texts = set()
    for text, count in text_counts.items():
        if count >= max(3, num_pages * 0.5):  # Appears on at least 3 pages or 50% of pages
            # Check if it looks like header/footer content
            if (any(keyword in text.lower() for keyword in ['page', 'version', '©', 'copyright']) or
                re.match(r'.*\d+\s+of\s+\d+.*', text.lower()) or
                len(text) < 10):  # Very short repetitive text
                repetitive_texts.add(text)
    
    for el in text_elements:
        # Remove only clearly repetitive header/footer content
        if el.text in repetitive_texts:
            continue
        filtered_elements.append(el)
    
    return filtered_elements

def find_numbering_level(text):
    """
    Detect numbering patterns like '1.', '1.2', '2.1.3' indicating heading level depth.
    Returns level (1,2,3) based on number of segments.
    """
    text = text.strip()
    
    # More comprehensive patterns for different numbering styles
    patterns = [
        r'^(\d+)\.?\s+',  # "1. " or "1 "
        r'^(\d+\.\d+)\.?\s+',  # "2.1 " or "2.1. "
        r'^(\d+\.\d+\.\d+)\.?\s+',  # "3.1.4 " or "3.1.4. "
    ]
    
    for i, pattern in enumerate(patterns):
        match = re.match(pattern, text)
        if match:
            return i + 1  # Return 1, 2, or 3 for the different patterns
    
    return None

def is_heading_like(text):
    """
    Check if text looks like a heading based on content patterns.
    """
    text = text.strip()
    
    # Skip if it's clearly not a heading
    if len(text) < 3:
        return False
    
    # Skip page numbers, versions, copyright info
    if (re.match(r'^\d+$', text) or 
        'version' in text.lower() or 
        'page' in text.lower() or 
        '©' in text or 
        'copyright' in text.lower()):
        return False
    
    # Positive indicators for headings
    # Has numbering (1., 2.1, etc.)
    if re.match(r'^\d+[\.\d\s]*[A-Z]', text):
        return True
    
    # Common heading words
    heading_words = ['introduction', 'overview', 'content', 'references', 'acknowledgements', 
                     'history', 'outcomes', 'requirements', 'structure', 'audience', 'objectives']
    if any(word in text.lower() for word in heading_words):
        return True
    
    # Starts with capital letter and has reasonable length
    if text[0].isupper() and 5 <= len(text) <= 100:
        return True
    
    return False

def assign_headings(text_elements, font_clusters, title_font_size):
    """
    Assign heading levels based on font size clusters, numbering, and content analysis.
    """
    # Filter out title font size from heading consideration
    heading_font_sizes = [fs for fs in font_clusters if abs(fs - title_font_size) > 0.5]
    heading_font_sizes.sort(reverse=True)  # Larger -> higher level
    
    print(f"Debug: Title font size: {title_font_size}")
    print(f"Debug: All font clusters: {font_clusters}")
    print(f"Debug: Heading font sizes: {heading_font_sizes}")
    
    # Map font sizes to heading levels - be more explicit
    font_size_to_level = {}
    
    # Assign levels based on font size hierarchy
    for i, fs in enumerate(heading_font_sizes):
        if i == 0:
            font_size_to_level[fs] = "H1"
        elif i == 1:
            font_size_to_level[fs] = "H2"
        else:
            font_size_to_level[fs] = "H3"
    
    print(f"Debug: Font size to level mapping: {font_size_to_level}")
    
    outline = []
    
    # Sort elements by page, then vertical position
    text_elements = sorted(text_elements, key=lambda x: (x.page_num, x.y0, x.x0))
    
    for el in text_elements:
        # Skip page 1 elements (title page)
        if el.page_num == 1:
            continue
        
        # Skip if text is too short or doesn't look like a heading
        if not is_heading_like(el.text):
            continue
        
        print(f"Debug: Processing element: {el.text[:50]}..., font_size: {el.font_size}")
        
        # Find closest font cluster
        closest_font_size = None
        min_diff = float('inf')
        for fs in heading_font_sizes:
            diff = abs(el.font_size - fs)
            if diff < min_diff:
                min_diff = diff
                closest_font_size = fs
        
        print(f"Debug: Closest font size: {closest_font_size}, diff: {min_diff}")
        
        # Skip if font size doesn't match any heading font size closely
        if closest_font_size is None or min_diff > 1.0:
            print(f"Debug: Skipping - no close font match")
            continue
        
        # Get initial level from font size
        level = font_size_to_level.get(closest_font_size, "H3")
        print(f"Debug: Initial level from font: {level}")
        
        # Override level based on numbering if present - this is key!
        num_level = find_numbering_level(el.text)
        if num_level is not None:
            if num_level == 1:
                level = "H1"
            elif num_level == 2:
                level = "H2"
            elif num_level >= 3:
                level = "H3"
            print(f"Debug: Numbering detected, level overridden to: {level}")
        
        print(f"Debug: Final level: {level}")
        print("---")
        
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
    
    # Detect title first (before removing headers/footers)
    title = detect_title(text_els)
    
    # Remove headers and footers (but be more conservative)
    text_els = remove_header_footer(text_els)
    
    # Cluster font sizes to determine font hierarchy
    font_clusters = cluster_font_sizes(text_els)
    
    # Get title font size
    first_page_elements = [el for el in text_els if el.page_num == 1]
    title_font_size = max([el.font_size for el in first_page_elements]) if first_page_elements else 0
    
    # Assign headings based on font clusters and heuristics
    outline = assign_headings(text_els, font_clusters, title_font_size)
    
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
        
        # Add debug output
        doc = fitz.open(file_path)
        text_els = extract_text_elements(doc)
        print(f"Debug: Total text elements: {len(text_els)}")
        
        # Show some sample elements and their font sizes
        sample_elements = [el for el in text_els if el.page_num <= 3 and len(el.text) > 5][:10]
        for el in sample_elements:
            print(f"Debug sample: '{el.text[:30]}...' - Font: {el.font_size} - Page: {el.page_num}")
        
        outline_data = process_pdf_file(file_path)

        # Compose output JSON filename
        base_name = os.path.splitext(pdf_file)[0]
        output_path = os.path.join(OUTPUT_DIR, f"{base_name}.json")

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(outline_data, f, indent=2, ensure_ascii=False)

        print(f"Output written to {output_path}")
        print("="*50)

if __name__ == "__main__":
    main()