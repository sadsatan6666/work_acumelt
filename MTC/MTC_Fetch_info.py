#this code is used to fetch some necessary value from PDF it just fetches and prints to the terminal or output window , which later can be used to automate (the work is still in progress :)
import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# ==========================================
# TOOL 1: NEIGHBOR FINDER (For Tensile Report)
# ==========================================
def find_value_neighbor(elements, label_text, required_keyword="Mpa"):
    label_bbox = None
    for element in elements:
        if label_text in element.get_text():
            label_bbox = element.bbox
            break     
    if not label_bbox: return "Label Not Found"

    lx0, ly0, lx1, ly1 = label_bbox
    best_candidate_text = "Not Found"
    closest_distance = 9999
    
    for element in elements:
        text = element.get_text().strip()
        ex0, ey0, ex1, ey1 = element.bbox
        if label_text in text: continue
        is_vertically_aligned = (ey0 < ly1 + 2) and (ey1 > ly0 - 2)
        is_to_the_right = ex0 >= (lx0 - 5)
        has_keyword = required_keyword in text
        if is_vertically_aligned and is_to_the_right and has_keyword:
            distance = ex0 - lx1
            if distance < closest_distance:
                closest_distance = distance
                best_candidate_text = text
    return best_candidate_text

# ==========================================
# TOOL 2: CLEANING (Universal)
# ==========================================
def extract_number_only(text):
    match = re.search(r"([\d\.]+)", text)
    if match: return match.group(1)
    return text

# ==========================================
# LOGIC A: PROCESS TENSILE REPORT
# ==========================================
def process_tensile_file(pdf_path):
    print(f"\n--- Processing Tensile Report: {pdf_path} ---")
    elements = []
    for page_layout in extract_pages(pdf_path, page_numbers=[0]):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                elements.append(element)

    # 1. Tensile
    raw_tensile = find_value_neighbor(elements, "Tensile Strength", "Mpa")
    val_tensile = extract_number_only(raw_tensile)
    
    # 2. Yield
    raw_yield = find_value_neighbor(elements, "Yield Strength", "Mpa")
    val_yield = extract_number_only(raw_yield)

    # 3. Elongation
    raw_elongation = find_value_neighbor(elements, "Elongation", "%")
    val_elongation = extract_number_only(raw_elongation)

    print(f"Tensile Strength: '{val_tensile}'")
    print(f"Yield Strength:   '{val_yield}'")
    print(f"Elongation:       '{val_elongation}'")
    return val_tensile, val_yield, val_elongation

# ==========================================
# LOGIC B: PROCESS HARDNESS REPORT (FIXED PRECISION)
# ==========================================
def process_hardness_file(pdf_path):
    print(f"\n--- Processing Hardness Report: {pdf_path} ---")
    
    elements = []
    for page_layout in extract_pages(pdf_path, page_numbers=[0]):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                elements.append(element)

    # Find "Hardness" labels
    hardness_labels = [e for e in elements if "Hardness" in e.get_text()]
    hardness_labels.sort(key=lambda x: x.bbox[3], reverse=True)
    
    extracted_values = []
    count = 1

    for label in hardness_labels:
        lx0, ly0, lx1, ly1 = label.bbox
        label_text = label.get_text().strip()
        
        found_val = None

        # --- CHECK 1: Same Box ---
        # STRICT REGEX: Must be followed by HBW
        match_inside = re.search(r"([\d\.]+)\s*HBW", label_text)
        if match_inside:
            found_val = match_inside.group(1)
        
        # --- CHECK 2: Neighbor Box ---
        if not found_val:
            closest_dist = 9999
            for element in elements:
                etext = element.get_text().strip()
                ex0, ey0, ex1, ey1 = element.bbox
                
                if "HBW" not in etext: continue
                
                if (ey0 < ly1 + 5) and (ey1 > ly0 - 5): # Vertical Alignment
                    if ex0 > lx0: # To the right
                        dist = ex0 - lx1
                        if dist < closest_dist:
                            # --- CRITICAL FIX IS HERE ---
                            # Before: re.search(r"([\d\.]+)", etext) -> Grabbed first number (2.285)
                            # Now:    re.search(r"([\d\.]+)\s*HBW", etext) -> Grabs number attached to HBW (172.9)
                            
                            n_match = re.search(r"([\d\.]+)\s*HBW", etext)
                            if n_match:
                                closest_dist = dist
                                found_val = n_match.group(1)

        if found_val:
            print(f"({count}) Found Hardness: '{found_val}'")
            extracted_values.append(found_val)
            count += 1
            
    return extracted_values

# ==========================================
# MAIN EXECUTION BLOCK
# ==========================================
if __name__ == "__main__":
    # FILE 1: Tensile Report
    path_tensile = "/home/johnny/MTCAUTO/MTCAUTO/TENCILE_OG/F326-029(AF427)-4(9).pdf"
    # Note: If you don't want to run tensile, just comment out the next line
    t_data = process_tensile_file(path_tensile)

    # FILE 2: Hardness Report
    path_hardness = "/home/johnny/MTCAUTO/MTCAUTO/HARDNESS_OG/F335-023(AF577)-2_ON_CASTING.pdf"
    h_data = process_hardness_file(path_hardness)
    
    print("\n=== ALL JOBS DONE ===")
