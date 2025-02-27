import os
import zipfile
from xml.etree import ElementTree as ET
import shutil
from openai import OpenAI


RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish", 
    "Pashto", "Sindhi", "Dhivehi", "Kurdish"
}

# Existing DOCX functions remain unchanged
def extract_docx(input_docx, extract_folder):
    """Extract DOCX contents to temporary folder"""
    with zipfile.ZipFile(input_docx, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)
    return os.path.join(extract_folder, "word", "document.xml")

def create_translated_docx(extract_folder, output_docx):
    """Create new DOCX file from modified extracted contents"""
    shutil.make_archive(output_docx.replace('.docx', ''), 'zip', extract_folder)
    if os.path.exists(output_docx):
        os.remove(output_docx)
    os.rename(output_docx.replace('.docx', '') + '.zip', output_docx)

def translate_xml_to_language(xml_path, target_language):
    """Translate text content in XML to the specified target language and set text direction"""
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    
    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    word_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    
    is_rtl = target_language in RTL_LANGUAGES
    
    for paragraph in root.findall(f'.//{word_ns}p'):
        text_elements = paragraph.findall(f'.//{word_ns}t')
        if not text_elements:  # Skip paragraphs with no text elements
            continue
        
        full_text = ""
        for text_elem in text_elements:
            if text_elem.text and text_elem.text.strip():
                full_text += text_elem.text
        
        if full_text.strip():  # Only translate if there’s meaningful text
            try:
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": f"You are a translator. Translate the following text to {target_language}."},
                        {"role": "user", "content": full_text}
                    ],
                    temperature=0.3
                )
                translated_text = response.choices[0].message.content
                
                # Distribute translated text back to elements (simplified: replace first, clear others)
                text_elements[0].text = translated_text
                for elem in text_elements[1:]:
                    elem.text = ""
            except Exception as e:
                print(f"Translation error for paragraph: {e}")
                continue
        
        # Set RTL direction for RTL languages
        if is_rtl:
            pPr = paragraph.find(f'{word_ns}pPr')
            if pPr is None:
                pPr = ET.SubElement(paragraph, f'{word_ns}pPr')
            
            bidi = pPr.find(f'{word_ns}bidi')
            if bidi is None:
                bidi = ET.SubElement(pPr, f'{word_ns}bidi')
            bidi.set(f'{word_ns}val', 'on')

    tree.write(xml_path, encoding='utf-8', xml_declaration=True)




def translate_text_file(input_txt, output_txt, target_language):
    """Translate a .txt file to the specified target language in segments"""
    client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    
    try:
        # Read the input text file
        with open(input_txt, 'r', encoding='utf-8') as file:
            text_content = file.read()
        
        if not text_content.strip():
            print("Input text file is empty")
            return
        
        # Split text into paragraphs (or customize the splitting logic)
        segments = text_content.split('\n\n')  # Split by double newline (paragraphs)
        translated_segments = []
        
        for segment in segments:
            if segment.strip():
                response = client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": f"You are a translator. Translate the following text to {target_language}."},
                        {"role": "user", "content": segment}
                    ],
                    temperature=0.3
                )
                translated_segments.append(response.choices[0].message.content)
        
        # Write translated content to output file
        with open(output_txt, 'w', encoding='utf-8') as file:
            file.write('\n\n'.join(translated_segments))
        
        print(f"Translation complete! Saved as: {output_txt}")
        
    except Exception as e:
        raise Exception(f"Error during TXT translation: {e}")



def translate_file(input_file, output_file, target_language):
    """Main function to translate either DOCX or TXT files"""
    file_extension = os.path.splitext(input_file)[1].lower()
    
    if file_extension == '.docx':
        extract_folder = 'temp_extract'
        
        if os.path.exists(extract_folder):
            try:
                shutil.rmtree(extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing temp_extract folder: {e}")
        
        try:
            os.makedirs(extract_folder, exist_ok=True)
        except Exception as e:
            raise Exception(f"Failed to create temp_extract folder: {e}")
        
        try:
            xml_path = extract_docx(input_file, extract_folder)
            translate_xml_to_language(xml_path, target_language)
            create_translated_docx(extract_folder, output_file)
            print(f"Translation complete! Saved as: {output_file}")
        except Exception as e:
            raise Exception(f"Error during DOCX translation: {e}")
        finally:
            if os.path.exists(extract_folder):
                try:
                    shutil.rmtree(extract_folder)
                except Exception as e:
                    print(f"Warning: Could not clean up temp_extract folder: {e}")
    
    elif file_extension == '.txt':
        translate_text_file(input_file, output_file, target_language)
    
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx or .txt files")

# Example usage:
if __name__ == "__main__":
    # For DOCX
    # translate_file("input.docx", "output_french.docx", "French")
    
    # For TXT
    # translate_file("input.txt", "output_spanish.txt", "Spanish")
    pass