import pdfplumber
import re
import pandas as pd
from collections import defaultdict
import report_generator as rg
import report_append as ra

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
    

def description_grabbing(pdf):
    """Grabs all descriptions from the PDF, character by character."""
    descs = []
    page_count = 1
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
                    descs.append((cleaned, is_bold, page_count))
                else:
                    descs.append((cleaned, is_bold, page_count))
        page_count += 1
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


def number_grabbing(pdf):
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


def grab_branch(pdf):
    """Grabs name of each branch."""
    branches = []
    for page in pdf.pages:
        text = page.extract_text().split('\n')
        if not text[4].startswith("Month"):
            branches.append(text[4])
        else:
            branches.append(text[3])
        
    return branches

def date_grabbing(pdf):
    """Grabs the date from the PDF."""
    date = ""
    for i, page in enumerate(pdf.pages):
        if i == 0:
            text = page.extract_text().split('\n')
            date = text[2]
        else:
            break
    return date


def open_pdf(path_to_pdf, output_path, output_filename):
    with pdfplumber.open(path_to_pdf) as pdf:
        descs = description_grabbing(pdf)
        lines = number_grabbing(pdf)
        del cols[8], cols[6], cols[3], cols[2]
        branches = grab_branch(pdf)
        date = date_grabbing(pdf)

        branch_dict = {}
        for b in branches:
            branch_dict[b] = []

        for l in lines:
            del l[7], l[5], l[2]
            for i in range(len(l)):
                element = l[i]
                cleaned_element = re.sub(r'[(),]', '', element)
                l[i] = int(cleaned_element)

        desc_line_pairs = []
        for i in range(len(lines)):
            desc_line_pairs.append((descs[i], lines[i]))

        rows = []
        for (desc, is_bold, pg_count), values in desc_line_pairs:
            row = [desc] + values + [is_bold]
            rows.append(row)
            associated_branch = branches[pg_count - 1]
            branch_dict[associated_branch].append(row)
        
        
        final_cols = ["description"] + cols + ["is_bold"]

        df_dict = {}
        for branch, data in branch_dict.items():
            df = pd.DataFrame(data, columns=final_cols)
            df_dict[branch] = df

        # Create Excel file
        complete_path = f"{output_path}\\{output_filename}.xlsx"
        load = True  # Set to True to load existing data or False to create new data
        if not load:
            rg.xlsx_create(df_dict, date, complete_path)
        else:
            # load new sheet onto existing sheet
            ra.xlsx_append(df_dict, date, complete_path)
        



