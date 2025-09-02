import pdfplumber
import re
import pandas as pd
from collections import defaultdict

months = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
cols = ["actual_month", "budget_month", "variance_month", "description", "actual_year", 
        "budget_year", "variance_year", "budget_total", "fcast_total", "actual_lastyear"]

def is_description_continuation(line):
    """Checks if the line is a continuation of a description. No digits"""
    return not re.search(r'\d', line)

def clean_desc_line(line):
    """Cleans the line by removing numbers, parentheses, and unnecessary whitespace."""
    sline = re.split(r'[0-9,()]+', line) #removes all numerical,() regexp
    if sline[0] == 'Report executed at ' or 'Order of' in sline[0] or 'SA' in sline[0] or any(month in sline[0] for month in months):
        return None
    onlywords_sline = [item for item in sline if item.strip()]
    if len(onlywords_sline) == 1:
        return onlywords_sline[0].strip()
    elif len(onlywords_sline) > 1:
        return ''.join(onlywords_sline).strip()
    return None
    

def description_grabbing():
    """Grabs all descriptions from the PDF, character by character."""
    descs = []
    for page in pdf.pages:
        chars = page.chars
        lines_by_y = defaultdict(list)

        for c in chars:
            lines_by_y[round(c["top"])].append(c)

        stitched_lines = []
        prev_line = ""
        prev_bold = False

        for y, line_chars in sorted(lines_by_y.items()):
            line_text = ''.join(c["text"] for c in sorted(line_chars, key=lambda x: x["x0"]))
            is_bold = any("Bold" in c["fontname"] for c in line_chars) #flag for 'Helvetica-Bold' font on characters - indicating sum
            if prev_line and is_description_continuation(line_text):
                prev_line += line_text
                prev_bold = prev_bold or is_bold
            else:
                if prev_line:
                    stitched_lines.append((prev_line, prev_bold))
                prev_line = line_text
                prev_bold = is_bold

        if prev_line:
            stitched_lines.append((prev_line, prev_bold))

        for line, is_bold in stitched_lines:
            cleaned = clean_desc_line(line)
            if cleaned: #not null
                if is_bold:
                    descs.append((cleaned, is_bold))
                else:
                    descs.append((cleaned, is_bold))
    return descs

def merge_alpha_components(line):
    """Merges components of a line that contain alphabetic characters."""
    result = []
    alpha_parts = []

    for comp in line:
        comp_clean = re.sub(r'[()]', '', comp)
        if re.search(r'[A-Za-z/&-]', comp_clean):
            alpha_parts.append(comp_clean)
        else:
            if alpha_parts:
                result.append(' '.join(alpha_parts))
                alpha_parts = []
            result.append(comp)

    if alpha_parts:
        result.append(' '.join(alpha_parts))

    return result


def number_grabbing():
    """Grabs all numbers from the PDF, merging components that contain alphabetic characters."""
    merged_lines = []
    for page in pdf.pages:
        text = page.extract_text().split('\n')
        for i, line in enumerate(text):
            line_parts = line.split()
            merged = merge_alpha_components(line_parts)
            if len(merged) < 10:
                continue
            full_line = [comp for comp in merged if not re.search(r'[A-Za-z/&-]', comp)]
            merged_lines.append(full_line)

    return merged_lines


with pdfplumber.open(r"C:\Users\ciaranqu\Documents\Projects\Proj1\SA727 Christchurch - I and E Cost Centre - APR 2025.pdf") as pdf:
    descs = description_grabbing()
    lines = number_grabbing()
    del cols[3]
    
    desc_line_pairs = []
    for i in range(len(lines)):
        desc_line_pairs.append((descs[i], lines[i]))
    
    rows = []
    for (desc, is_bold), values in desc_line_pairs:
        row = [desc] + values + [is_bold]
        rows.append(row)

    final_cols = ["description"] + cols + ["is_bold"]

    # Create DataFrame
    df = pd.DataFrame(rows, columns=final_cols)
    
    print(df)
    print(desc_line_pairs[0])
    print(descs[0])
    print(lines[0])


