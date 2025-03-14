import requests
import xml.etree.ElementTree as ET
from langchain.chat_models import ChatOpenAI
from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import os

# Initialize OpenAI LLM
llm = ChatOpenAI(model="gpt-3.5-turbo", temperature=0)

# Build translation prompt template
translation_prompt = PromptTemplate(
    input_variables=["text", "target_language"],
    template="Translate the following text to {target_language} while preserving original meaning: {text}"
)

# Build translation chain
translation_chain = LLMChain(prompt=translation_prompt, llm=llm)

def translate_text(text, target_language="Arabic"):
    """
    Uses LangChain's LLMChain to translate the provided text.
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


def docx_to_xliff(input_file_path):
    """
    Convert .docx to XLIFF using RapidAPI with multipart/form-data
    """
    url = "https://translated-matecat-filters-v1.p.rapidapi.com/api/v2/original2xliff"
    headers = {
        'x-rapidapi-host': 'translated-matecat-filters-v1.p.rapidapi.com',
        'x-rapidapi-key': '385d0a1310msh4e957f9d0387159p1c24e7jsn4e1c90b423c5'
    }
    
    with open(input_file_path, 'rb') as file:
        files = {
            'document': (os.path.basename(input_file_path), file, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        }
        data = {
            'sourceLocale': 'en-GB',
            'targetLocale': 'it-IT'  # We'll change this later in the XLIFF
        }
        
        # Send request with multipart/form-data
        response = requests.post(url, headers=headers, data=data, files=files)
        
        if response.status_code == 200:
            try:
                # Try to parse as JSON first
                response_json = response.json()
                xliff_content = response_json.get("xliff")
                if not xliff_content:
                    raise ValueError("No 'xliff' key found in JSON response")
                # Convert string to bytes
                xliff_bytes = xliff_content.encode('utf-8')
            except ValueError:
                # If not JSON, assume it's raw XLIFF content as bytes
                xliff_bytes = response.content
            
            # Write bytes to file
            with open('original.xliff', 'wb') as f:
                f.write(xliff_bytes)
            return 'original.xliff'
        else:
            raise Exception(f"Failed to convert docx to xliff: {response.status_code} - {response.text}")
# Rest of your code remains the same, just replace the docx_to_xliff function with this version


def translate_xliff(input_xliff_path, output_xliff_path):
    """
    Read XLIFF, translate content to Arabic, and create new XLIFF
    """
    # Parse XLIFF
    tree = ET.parse(input_xliff_path)
    root = tree.getroot()
    
    # Update target language to Arabic
    for file_elem in root.findall('.//{urn:oasis:names:tc:xliff:document:1.2}file'):
        file_elem.set('target-language', 'ar-SA')
    
    # Translate each segment
    for trans_unit in root.findall('.//{urn:oasis:names:tc:xliff:document:1.2}trans-unit'):
        source = trans_unit.find('{urn:oasis:names:tc:xliff:document:1.2}source')
        target = trans_unit.find('{urn:oasis:names:tc:xliff:document:1.2}target')
        
        if source is not None:
            source_text = source.text
            if source_text and source_text.strip():
                translated_text = translate_text(source_text, "Arabic")
                
                # If target doesn't exist, create it
                if target is None:
                    target = ET.SubElement(trans_unit, '{urn:oasis:names:tc:xliff:document:1.2}target')
                target.text = translated_text
    
    # Save translated XLIFF
    tree.write(output_xliff_path, encoding='utf-8', xml_declaration=True)
    return output_xliff_path



def xliff_to_docx(input_xliff_path, output_docx_path):
    """
    Convert XLIFF back to .docx using RapidAPI with multipart/form-data
    """
    url = "https://translated-matecat-filters-v1.p.rapidapi.com/api/v2/xliff2original"
    headers = {
        'x-rapidapi-host': 'translated-matecat-filters-v1.p.rapidapi.com',
        'x-rapidapi-key': '640e62c465mshaf4612908fb7e62p1845ebjsnfca96883ad61'
    }
    
    with open(input_xliff_path, 'rb') as file:
        files = {
            'xliff': (os.path.basename(input_xliff_path), file, 'application/xml')
        }
        
        # Send request with multipart/form-data (handled automatically by requests)
        response = requests.post(url, headers=headers, files=files)
        
        if response.status_code == 200:
            with open(output_docx_path, 'wb') as f:
                f.write(response.content)
            return output_docx_path
        else:
            raise Exception(f"Failed to convert xliff to docx: {response.status_code} - {response.text}")
        




def main():
    input_docx = "input.docx"  # Replace with your input file path
    output_docx = "translated_output.docx"
    
    try:
        # Step 1: Convert docx to xliff
        print("Converting .docx to XLIFF...")
        xliff_path = docx_to_xliff(input_docx)
        
        # Step 2: Translate XLIFF to Arabic
        # print("Translating XLIFF content to Arabic...")
        # translated_xliff_path = translate_xliff(xliff_path, 'translated.xliff')
        
        # Step 3: Convert translated XLIFF back to docx
        print("Converting translated XLIFF back to .docx...")
        xliff_to_docx(xliff_path, output_docx)
        
        print(f"Translation completed successfully. Output saved to {output_docx}")
        
        # Clean up temporary files
        os.remove(xliff_path)
        # os.remove(translated_xliff_path)
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()



# You are provided with an XML snippet containing a <trans-unit> element. Your task is to translate the English text within the XML to Arabic, but you must preserve the XML tags, attributes, and overall structure exactly as they are. 

# Specifically:
# - Translate the text content inside the <source>, <seg-source>, and <target> tags, including the text inside nested tags such as <mrk> and <g>.
# - Do NOT modify any XML tags or attributes.
# - Ensure that only the text is translated and that the output retains the original XML structure and tag hierarchy.

# For example, given the following input:
# <trans-unit id="NFDBB2FA9-tu1" xml:space="preserve">
#     <source xml:lang="en-GB">ContractsCounsel has assisted 419 clients with commercial leases and maintains a network of 214 <g id="1">real estate</g> lawyers available daily. These lawyers collectively have <g id="2">71 reviews</g> to help you choose the best lawyer for your needs. Customers rate lawyers for commercial lease matters 4.97.</source>
#     <seg-source>
#         <mrk mid="0" mtype="seg">ContractsCounsel has assisted 419 clients with commercial leases and maintains a network of 214 <g id="1">real estate</g> lawyers available daily.</mrk>
#         <mrk mid="1" mtype="seg">These lawyers collectively have <g id="2">71 reviews</g> to help you choose the best lawyer for your needs.</mrk>
#         <mrk mid="2" mtype="seg">Customers rate lawyers for commercial lease matters 4.97.</mrk>
#     </seg-source>
#     <target xml:lang="it-IT">
#         <mrk mid="0" mtype="seg">ContractsCounsel has assisted 419 clients with commercial leases and maintains a network of 214 <g id="1">real estate</g> lawyers available daily.</mrk>
#         <mrk mid="1" mtype="seg">These lawyers collectively have <g id="2">71 reviews</g> to help you choose the best lawyer for your needs.</mrk>
#         <mrk mid="2" mtype="seg">Customers rate lawyers for commercial lease matters 4.97.</mrk>
#     </target>
# </trans-unit>

# Your output should be the same XML snippet with all English text translated to Arabic while keeping every XML tag and attribute unchanged.

























