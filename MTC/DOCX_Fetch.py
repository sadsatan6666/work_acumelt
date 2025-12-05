#this script will fetch all the necessary info from micro report which later can be used to make an MTC
import zipfile
import xml.etree.ElementTree as ET
import os
import re

def extract_hidden_values_final_hybrid(docx_path):
    print(f"--- X-RAY SCANNING: {docx_path} ---")

    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
    except Exception as e:
        print(f"Error reading file content: {e}")
        return

    tree = ET.fromstring(xml_content)
    
    all_text_chunks = []
    for elem in tree.iter():
        if elem.tag.endswith('}t'):
            text_content = elem.text
            if text_content is not None and text_content.strip():
                all_text_chunks.append(text_content.strip())

    # Define preference: 'last' for fields that might be empty early (like Graphite Size)
    # 'first' for fields consistently populated in the initial summary (like Ratio)
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
                    target_index = i # Store the last index found
        elif occurrence_preference == "first":
            for i, chunk in enumerate(all_text_chunks):
                if label.lower() in chunk.lower():
                    target_index = i
                    break # Stop at the first match

        # Pass 2: Search for the value starting from the target label index
        if target_index != -1:
            neighbors = all_text_chunks[target_index+1:target_index+6]
            
            # --- Advanced Value Search Strategy ---
            
            # 1. Graphite Fraction
            if label == "Graphite Fraction":
                for j, n in enumerate(neighbors):
                    # Look for combined or split percentage formats
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
                # Combine chunks and use regex to find the ratio pattern (X%/Y%)
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
                        # Clean up trailing non-numeric characters (like '.')
                        found_value = re.sub(r'[\s\.\,]+$', '', n)
                        break
                
        print(f"{label:30} : '{found_value}'")
    print("-" * 40)

# The function can now be used with your file path:
extract_hidden_values_final_hybrid("/home/johnny/MTCAUTO/MTCAUTO/MICRO_REPORT/F305-013-(BE406)-6.docx")
