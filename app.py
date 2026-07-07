import streamlit as st
import json
import re
from docx import Document
import io

# --- 1. CHEMISTRY REGEX FUNCTION ---
def append_science_text(paragraph, raw_text):
    parts = re.split(r'(<sub>.*?</sub>)', raw_text)
    for part in parts:
        if part.startswith('<sub>') and part.endswith('</sub>'):
            subscript_content = part.replace('<sub>', '').replace('</sub>', '')
            run = paragraph.add_run(subscript_content)
            run.font.subscript = True
        else:
            paragraph.add_run(part)

# --- 2. V3 COMPONENT DOCUMENT COMPILER ---
def generate_worksheet(json_data, template_path):
    doc = Document(template_path)
    title = json_data.get("worksheet_title", "Science Worksheet")
    blocks = json_data.get("document_blocks", [])

    for p in doc.paragraphs:
        # --- A. TITLE ---
        if "[[WORKSHEET_TITLE]]" in p.text:
            p.text = "" 
            append_science_text(p, title)
            
        # --- B. THE DYNAMIC CONTENT BLOCK ---
        elif "[[CONTENT_HERE]]" in p.text:
            
            for block in blocks:
                b_type = block.get("type", "unknown")
                
                if b_type == "subheading":
                    new_p = p.insert_paragraph_before('')
                    new_p.style = 'Heading 2'
                    append_science_text(new_p, block.get("content", ""))
                    
                elif b_type == "paragraph":
                    new_p = p.insert_paragraph_before('')
                    new_p.style = 'Normal'
                    append_science_text(new_p, block.get("content", ""))
                    
                elif b_type == "question":
                    new_p = p.insert_paragraph_before('')
                    new_p.style = 'List Number'
                    append_science_text(new_p, block.get("content", ""))
                    
                elif b_type == "image":
                    new_p = p.insert_paragraph_before('')
                    new_p.style = 'Normal'
                    # Formats the image placeholder clearly for the teacher
                    image_desc = block.get("content", "Insert image here")
                    run = new_p.add_run(f"[ 📷 PLACEHOLDER: {image_desc} ]")
                    run.font.bold = True
                    
                elif b_type == "table":
                    table_data = block.get("table_data", {})
                    rows_data = table_data.get("rows", [])
                    
                    if rows_data:
                        col_count = table_data.get("column_count", len(rows_data[0]))
                        row_count = len(rows_data)
                        table = doc.add_table(rows=row_count, cols=col_count)
                        try:
                            table.style = 'Science_Table_Style'
                        except KeyError:
                            pass
                        
                        # XML magic: Inserts table precisely before the current anchor
                        p._p.addprevious(table._tbl)
                        
                        # Populate cells
                        for r_idx, row_data in enumerate(rows_data):
                            for c_idx, cell_text in enumerate(row_data):
                                cell_paragraph = table.cell(r_idx, c_idx).paragraphs[0]
                                cell_paragraph.text = "" 
                                append_science_text(cell_paragraph, str(cell_text))
                                
                else:
                    # THE SAFETY NET: Catches any AI hallucinations
                    new_p = p.insert_paragraph_before('')
                    new_p.style = 'Normal'
                    error_text = str(block.get("content", f"[AI Generated Unrecognized Block: '{b_type}']"))
                    append_science_text(new_p, error_text)
            
            # Delete the master anchor paragraph
            p._element.getparent().remove(p._element)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer, title

# --- 3. STREAMLIT UI ---
st.set_page_config(page_title="Science Worksheet Generator", page_icon="🔬")

st.title("🔬 Science Worksheet Generator (V3)")
st.write("Paste the JSON output from Gemini below.")

json_input = st.text_area("Paste Gemini JSON here:", height=300)

if st.button("Generate Worksheet", type="primary"):
    if not json_input.strip():
        st.error("Please paste the JSON payload first.")
    else:
        try:
            json_data = json.loads(json_input)
            word_buffer, doc_title = generate_worksheet(json_data, "Generic Template.docx")
            st.success("Worksheet generated successfully!")
            
            st.download_button(
                label="📥 Download Word Document",
                data=word_buffer,
                file_name=f"{doc_title}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except json.JSONDecodeError:
            st.error("❌ Invalid JSON. Please check that Gemini output exactly matches the required structure.")
        except Exception as e:
            st.error(f"❌ An error occurred: {e}")
