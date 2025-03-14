import os
import json
import requests
import xml.etree.ElementTree as ET
from io import BytesIO
from dotenv import load_dotenv

# Updated imports according to new LangChain structure:
from langchain.chains import LLMChain
from langchain_core.prompts import PromptTemplate
from langchain_openai import ChatOpenAI

# Load environment variables (ensure OPENAI_API_KEY is set)
load_dotenv()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    raise ValueError("OPENAI_API_KEY environment variable not set.")

# API configuration for Matecat remains unchanged.
API_HOST = "translated-matecat-filters-v1.p.rapidapi.com"
API_KEY = "640e62c465mshaf4612908fb7e62p1845ebjsnfca96883ad61"
HEADERS = {
    "x-rapidapi-host": API_HOST,
    "x-rapidapi-key": API_KEY
}

# Configure ChatOpenAI via LangChain (using a supported model, e.g., gpt-3.5-turbo)
llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0)

# Build a prompt template for translation
translation_prompt = PromptTemplate(
    input_variables=["text", "target_language"],
    template=(
        "Translate the following text to {target_language} while preserving "
        "original formatting, colors, styling, and tables if present:\n\n{text}"
    )
)

# Build the translation chain
translation_chain = LLMChain(prompt=translation_prompt, llm=llm)

def convert_docx_to_xliff(input_docx_path, source_locale="en-GB", target_locale="ar-SA", extraction_params=None):
    url = "https://translated-matecat-filters-v1.p.rapidapi.com/api/v2/original2xliff"
    
    data = {
        "sourceLocale": source_locale,
        "targetLocale": target_locale,
    }
    if extraction_params:
        data["extractionParams"] = json.dumps(extraction_params)
    
    if not os.path.exists(input_docx_path):
        print(f"Error: File {input_docx_path} not found!")
        return None

    with open(input_docx_path, "rb") as f:
        files = {
            "document": (os.path.basename(input_docx_path), f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        }
        print("Converting DOCX to XLIFF...")
        response = requests.post(url, headers=HEADERS, data=data, files=files)

    if response.status_code == 200:
        try:
            response_json = response.json()
            if response_json.get("successful") and "xliff" in response_json:
                print("Conversion to XLIFF successful!")
                xliff_str = response_json
                return xliff_str.encode("utf-8")
            else:
                print("Error: API response did not include valid XLIFF content.")
                return None
        except Exception as e:
            print("Error parsing JSON response:", e)
            return None
    else:
        print("Error converting DOCX to XLIFF:", response.status_code, response.text)
        return None

def convert_xliff_to_docx(xliff_data, source_locale="en-GB", target_locale="ar-SA"):
    url = "https://translated-matecat-filters-v1.p.rapidapi.com/api/v2/xliff2original"
    data = {
        "sourceLocale": source_locale,
        "targetLocale": target_locale,
    }
    files = {
        "xliff": ("updated_file.xliff", xliff_data, "application/xml")
    }
    print("Converting XLIFF back to DOCX...")
    response = requests.post(url, headers=HEADERS, data=data, files=files)
    
    if response.status_code == 200:
        print("Conversion to DOCX successful!")
        return response.content['document']
    else:
        print("Error converting XLIFF to DOCX:", response.status_code, response.text)
        return None

def translate_text(text, target_language="Arabic"):
    """
    Uses LangChain's LLMChain (with ChatOpenAI) to translate the provided text.
    """
    try:
        print(f"Translating text: {text}")
        result = translation_chain.run(text=text, target_language=target_language)
        translated = result.strip()
        print(f"Translated text: {translated}")
        return translated
    except Exception as e:
        print("Translation error:", e)
        return text

def process_xliff(xliff_data, target_language):
    """
    Processes and translates the XLIFF data.
    It parses the XLIFF (XML), locates each translatable unit within the 
    file element for 'word/document.xml', translates the text in the <source> element,
    and stores the result in a corresponding <target> element.
    The original structure and formatting remain intact.
    """
    try:
        xliff_string = xliff_data.decode("utf-8-sig", errors="replace")
        start_index = xliff_string.find('<?xml')
        if start_index != -1:
            xliff_string = xliff_string[start_index:]
        else:
            print("Warning: XML declaration not found; proceeding with full string.")
        tree = ET.ElementTree(ET.fromstring(xliff_string))
    except Exception as e:
        print("Error parsing XLIFF:", e)
        return xliff_data

    ns = {'xliff': 'urn:oasis:names:tc:xliff:document:1.2'}
    root = tree.getroot()
    print(f"Root element: {root.tag}")
    
    # Locate the file element that contains the DOCX content.
    file_elem = None
    for fe in root.findall("xliff:file", ns):
        if fe.get("original") == "word/document.xml":
            file_elem = fe
            break

    if file_elem is None:
        print("No file element with original='word/document.xml' found. Exiting process_xliff without changes.")
        return xliff_data

    # Translate each trans-unit within this file element.
    for trans_unit in file_elem.findall(".//xliff:trans-unit", ns):
        print(f"Processing trans_unit: {trans_unit.attrib}")
        source_elem = trans_unit.find("xliff:source", ns)
        if source_elem is None or not source_elem.text:
            continue
        original_text = source_elem.text
        translated_text = translate_text(original_text, target_language)
        target_elem = trans_unit.find("xliff:target", ns)
        if target_elem is None:
            target_elem = ET.SubElement(trans_unit, f"{{{ns['xliff']}}}target")
        target_elem.text = translated_text

    ET.register_namespace('', ns['xliff'])
    bio = BytesIO()
    tree.write(bio, encoding="utf-8", xml_declaration=True)
    updated_xliff = bio.getvalue()
    print("XLIFF translation complete!")
    return updated_xliff

def main():
    input_docx = "input.docx"
    output_docx = "output.docx"

    # Convert DOCX to XLIFF using Matecat API with Arabic as target locale.
    xliff_data = convert_docx_to_xliff(input_docx, target_locale="ar-SA")
    print(xliff_data)
    if not xliff_data:
        return
    
    with open("intermediate.xliff", "wb") as f:
        f.write(xliff_data)
    print("Intermediate XLIFF saved as 'intermediate.xliff'.")

    # Process and translate the XLIFF (set target language to "Arabic").
    processed_xliff = process_xliff(xliff_data, target_language="Arabic")
    with open("processed.xliff", "wb") as f:
        f.write(processed_xliff)
    print("processed XLIFF saved as 'processed.xliff'.")

    # Convert the translated XLIFF back to DOCX using the same Arabic target locale.
    docx_data = convert_xliff_to_docx(processed_xliff, target_locale="ar-SA")
    if not docx_data:
        return

    with open(output_docx, "wb") as f:
        f.write(docx_data)
    print("Final DOCX saved as:", output_docx)

if __name__ == "__main__":
    main()
