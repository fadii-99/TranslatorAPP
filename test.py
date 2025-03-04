import os
from docx import Document

def docx_to_rtf(docx_path, rtf_path):
    # Check if the input file exists
    if not os.path.exists(docx_path):
        raise FileNotFoundError(f"The file {docx_path} does not exist.")
    
    # Load the .docx file
    doc = Document(docx_path)
    
    # Basic RTF header
    rtf_content = r"{\rtf1\ansi\deff0 {\fonttbl {\f0 Times New Roman;}}\f0\fs24\n"
    
    # Extract paragraphs and add to RTF content
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:  # Only add non-empty paragraphs
            # Escape special RTF characters
            text = text.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")
            rtf_content += text + r"\par\n"
    
    # Close the RTF document
    rtf_content += "}"
    
    # Write to the output file
    with open(rtf_path, "w", encoding="utf-8") as rtf_file:
        rtf_file.write(rtf_content)

# Example usage
if __name__ == "__main__":
    input_file = "Lawyers Docs.docx"  # Replace with your .docx file path
    output_file = "Lawyers Docs.rtf"  # Desired output .rtf file path
    
    try:
        docx_to_rtf(input_file, output_file)
        print(f"Successfully converted {input_file} to {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")