#this file contains all in PDF and DOCX fetch code for MTC
# ==============================================================================
# UNIFIED MTC DATA EXTRACTION SCRIPT
# This script combines DOCX Microstructure analysis and PDF Mechanical analysis.
# ==============================================================================

import zipfile
import xml.etree.ElementTree as ET
import os
import re
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# ==============================================================================
# PART 1: DOCX MICROSTRUCTURE EXTRACTION (Based on code_1)
# ==============================================================================

def extract_micro_data_from_docx(docx_path):
    """
    Extracts microstructure values (Nodularity, Size, Ratio, etc.) from a DOCX file.
    Uses a hybrid approach (first/last occurrence) to reliably locate data in tables.
    """
    results = {}
    print(f"\n--- X-RAY SCANNING: {docx_path} ---")

    try:
        if not os.path.exists(docx_path):
            print("Error: DOCX File not found.")
            return results
            
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
    except Exception as e:
        print(f"Error reading DOCX content: {e}")
        return results

    tree = ET.fromstring(xml_content)
    all_text_chunks = []
    for elem in tree.iter():
        if elem.tag.endswith('}t'):
            text_content = elem.text
            if text_content is not None and text_content.strip():
                all_text_chunks.append(text_content.strip())

    # Define preference: 'last' for primary results (Graphite Size), 'first' for early data (Ratio)
    target_labels = {
        "Graphite Nodularity": "last",
        "Nodular Particles per mm²": "last",
        "Graphite Size": "last",
        "Graphite Form": "last",
        "Graphite Fraction": "last",
        "Ferrite / Pearlite Ratio": "first"
    }

    print("-" * 40)
    
    for label, occurrence_preference in target_labels.items():
        found_value = "Not Found"
        target_index = -1
        
        # Pass 1: Determine the target index (first or last occurrence)
        if occurrence_preference == "last":
            for i, chunk in enumerate(all_text_chunks):
                if label.lower() in chunk.lower():
                    target_index = i
        elif occurrence_preference == "first":
            for i, chunk in enumerate(all_text_chunks):
                if label.lower() in chunk.lower():
                    target_index = i
                    break

        # Pass 2: Search for the value starting from the target label index
        if target_index != -1:
            neighbors = all_text_chunks[target_index+1:target_index+6]
            
            # --- Advanced Value Search Strategy ---
            
            # 1. Graphite Fraction
            if label == "Graphite Fraction":
                for j, n in enumerate(neighbors):
                    if "%" in n and any(c.isdigit() for c in n):
                        found_value = n
                        break
                    if any(c.isdigit() for c in n) and j + 1 < len(neighbors) and neighbors[j+1] == "%":
                        found_value = f"{n}{neighbors[j+1]}"
                        break
            
            # 2. Graphite Form
            elif label == "Graphite Form":
                for n in neighbors:
                     if ("(" in n and ")" in n):
                        found_value = n
                        break

            # 3. Ferrite / Pearlite Ratio
            elif label == "Ferrite / Pearlite Ratio":
                combined = "".join(neighbors[0:3])
                match = re.search(r"(\d+\.?\d*%\s*/\s*\d+\.?\d*%)", combined)
                if match:
                    found_value = match.group(1)
            
            # 4. Graphite Nodularity
            elif label == "Graphite Nodularity":
                for n in neighbors:
                    if "%" in n and len(n) > 1:
                        found_value = n
                        break
            
            # 5. Nodular Particles per mm² and Graphite Size (General Number Logic)
            elif label in ["Nodular Particles per mm²", "Graphite Size"]:
                for n in neighbors:
                    if any(c.isdigit() for c in n) and not n.endswith('%'):
                        found_value = re.sub(r'[\s\.\,]+$', '', n) # Clean up punctuation
                        break
            
            results[label] = found_value
            print(f"{label:30} : '{found_value}'")
            
    print("-" * 40)
    return results

# ==============================================================================
# PART 2: PDF MECHANICAL ANALYSIS EXTRACTION (Based on code_2)
# ==============================================================================

# --- TOOL 1: NEIGHBOR FINDER (For Tensile Report) ---
def find_value_neighbor(elements, label_text, required_keyword="Mpa"):
    """Finds a value element next to a label based on position and keyword."""
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

# --- TOOL 2: CLEANING (Universal) ---
def extract_number_only(text):
    """Extracts the first number/decimal from a string."""
    match = re.search(r"([\d\.]+)", text)
    if match: return match.group(1)
    return text

# --- LOGIC A: PROCESS TENSILE REPORT ---
def process_tensile_file(pdf_path):
    """Extracts Tensile, Yield, and Elongation from a PDF report."""
    
    if not os.path.exists(pdf_path):
        print(f"Error: PDF File not found at {pdf_path}")
        return None, None, None
        
    print(f"\n--- Processing Tensile Report: {pdf_path} ---")
    elements = []
    # Read only the first page
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

# --- LOGIC B: PROCESS HARDNESS REPORT ---
def process_hardness_file(pdf_path):
    """Extracts Hardness values (HBW) from a PDF report."""
    
    if not os.path.exists(pdf_path):
        print(f"Error: PDF File not found at {pdf_path}")
        return []
        
    print(f"\n--- Processing Hardness Report: {pdf_path} ---")
    
    elements = []
    # Read only the first page
    for page_layout in extract_pages(pdf_path, page_numbers=[0]):
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                elements.append(element)

    # Find "Hardness" labels and sort them
    hardness_labels = [e for e in elements if "Hardness" in e.get_text()]
    hardness_labels.sort(key=lambda x: x.bbox[3], reverse=True) # Sort top-to-bottom
    
    extracted_values = []
    count = 1

    for label in hardness_labels:
        lx0, ly0, lx1, ly1 = label.bbox
        label_text = label.get_text().strip()
        
        found_val = None

        # --- CHECK 1: Same Box (Strict Regex for HBW) ---
        match_inside = re.search(r"([\d\.]+)\s*HBW", label_text)
        if match_inside:
            found_val = match_inside.group(1)
        
        # --- CHECK 2: Neighbor Box (Vertical Alignment Check) ---
        if not found_val:
            closest_dist = 9999
            for element in elements:
                etext = element.get_text().strip()
                ex0, ey0, ex1, ey1 = element.bbox
                
                # Must contain the unit
                if "HBW" not in etext: continue
                
                if (ey0 < ly1 + 5) and (ey1 > ly0 - 5): # Vertical Alignment
                    if ex0 > lx0: # To the right
                        dist = ex0 - lx1
                        if dist < closest_dist:
                            # Search for the number immediately preceding HBW
                            n_match = re.search(r"([\d\.]+)\s*HBW", etext)
                            if n_match:
                                closest_dist = dist
                                found_val = n_match.group(1)

        if found_val:
            print(f"({count}) Found Hardness: '{found_val}'")
            extracted_values.append(found_val)
            count += 1
            
    return extracted_values

# ==============================================================================
# MAIN EXECUTION BLOCK (Unified Entry Point)
# ==============================================================================

if __name__ == "__main__":
    # --- DEFINE FILE PATHS HERE ---
    # Replace these placeholder paths with the actual file paths on your system
    path_micro_report = "/home/johnny/MTCAUTO/MTCAUTO/MICRO_REPORT/F305-013-(BE406)-6.docx"
    path_tensile_report = "/home/johnny/MTCAUTO/MTCAUTO/TENCILE_OG/F326-029(AF427)-4(9).pdf"
    path_hardness_report = "/home/johnny/MTCAUTO/MTCAUTO/HARDNESS_OG/F335-023(AF577)-2_ON_CASTING.pdf"

    # 1. PROCESS MICROSTRUCTURE REPORT (DOCX)
    micro_data = extract_micro_data_from_docx(path_micro_report)

    # 2. PROCESS TENSILE REPORT (PDF)
    tensile_data = process_tensile_file(path_tensile_report)

    # 3. PROCESS HARDNESS REPORT (PDF)
    hardness_data = process_hardness_file(path_hardness_report)
    
    print("\n\n=== ALL JOBS DONE ===")
    print("\n--- SUMMARY ---")
    print("Microstructure Data (DOCX):", micro_data)
    print("Tensile Data (PDF):", f"Tensile={tensile_data[0]}, Yield={tensile_data[1]}, Elongation={tensile_data[2]}" if tensile_data else "N/A")
    print("Hardness Values (PDF):", hardness_data)
