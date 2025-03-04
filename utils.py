import os
import zipfile
import shutil
from xml.etree import ElementTree as ET
from openai import OpenAI
from lxml import etree
import re
import subprocess




RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish", 
    "Pashto", "Sindhi", "Dhivehi", "Kurdish"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.extract_folder = 'temp_extract'
        self.word_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    def extract_docx(self):
        """Extract DOCX contents to a temporary folder and return the document.xml path"""
        with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
            zip_ref.extractall(self.extract_folder)
        return os.path.join(self.extract_folder, "word", "document.xml")

    def create_translated_docx(self):
        """Create a new DOCX file from the modified extracted contents"""
        base_name = self.output_file.replace('.docx', '')
        shutil.make_archive(base_name, 'zip', self.extract_folder)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(base_name + '.zip', self.output_file)

    def translate_xml_to_language(self, xml_path):
        """Translate text content in XML and adjust text direction for RTL languages"""
        tree = ET.parse(xml_path)
        root = tree.getroot()
        ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        is_rtl = self.target_language in RTL_LANGUAGES

        for paragraph in root.findall(f'.//{self.word_ns}p'):
            text_elements = paragraph.findall(f'.//{self.word_ns}t')
            if not text_elements:
                continue

            # Concatenate all text elements in the paragraph
            full_text = "".join(text_elem.text for text_elem in text_elements if text_elem.text and text_elem.text.strip())
            if full_text.strip():
                try:
                    response = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "system", "content": f"You are a translator. Translate the following text to {self.target_language}."},
                            {"role": "user", "content": full_text}
                        ],
                        temperature=0.3
                    )
                    translated_text = response.choices[0].message.content

                    # Replace first text element with translated text; clear the others.
                    text_elements[0].text = translated_text
                    for elem in text_elements[1:]:
                        elem.text = ""
                except Exception as e:
                    print(f"Translation error for paragraph: {e}")
                    continue

            # Set RTL direction if required
            if is_rtl:
                pPr = paragraph.find(f'{self.word_ns}pPr')
                if pPr is None:
                    pPr = ET.SubElement(paragraph, f'{self.word_ns}pPr')
                bidi = pPr.find(f'{self.word_ns}bidi')
                if bidi is None:
                    bidi = ET.SubElement(pPr, f'{self.word_ns}bidi')
                bidi.set(f'{self.word_ns}val', 'on')

        tree.write(xml_path, encoding='utf-8', xml_declaration=True)

    def run(self):
        # Clean up any existing temporary extraction folder
        if os.path.exists(self.extract_folder):
            try:
                shutil.rmtree(self.extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing {self.extract_folder} folder: {e}")
        os.makedirs(self.extract_folder, exist_ok=True)

        try:
            xml_path = self.extract_docx()
            self.translate_xml_to_language(xml_path)
            self.create_translated_docx()
            print(f"Translation complete! Saved as: {self.output_file}")
        except Exception as e:
            print(f"Error during DOCX translation: {e}")
        finally:
            if os.path.exists(self.extract_folder):
                try:
                    shutil.rmtree(self.extract_folder)
                except Exception as e:
                    print(f"Warning: Could not clean up {self.extract_folder} folder: {e}")



class TxtTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    def translate_text_file(self):
        """Read a TXT file, translate its content in segments, and write the translated text to an output file."""
        try:
            with open(self.input_file, 'r', encoding='utf-8') as file:
                text_content = file.read()

            if not text_content.strip():
                print("Input text file is empty")
                return

            # Split text into segments (e.g., paragraphs)
            segments = text_content.split('\n\n')
            translated_segments = []

            for segment in segments:
                if segment.strip():
                    response = self.client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "system", "content": f"You are a translator. Translate the following text to {self.target_language}."},
                            {"role": "user", "content": segment}
                        ],
                        temperature=0.3
                    )
                    translated_segments.append(response.choices[0].message.content)

            with open(self.output_file, 'w', encoding='utf-8') as file:
                file.write('\n\n'.join(translated_segments))

            print(f"Translation complete! Saved as: {self.output_file}")
        except Exception as e:
            print(f"Error during TXT translation: {e}")

    def run(self):
        self.translate_text_file()



class OdtTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.extract_dir = "extracted_odt"
        self.content_xml_path = os.path.join(self.extract_dir, "content.xml")
        self.namespaces = {
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
        }
        # Initialize the LLM translation client
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
    
    def extract_odt(self):
        if not os.path.exists(self.input_file):
            raise FileNotFoundError(f"Input file '{self.input_file}' not found!")
        with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
            zip_ref.extractall(self.extract_dir)
    
    def parse_content(self):
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(self.content_xml_path, parser)
        root = tree.getroot()
        # Register namespaces for correct XML handling
        for prefix, uri in self.namespaces.items():
            etree.register_namespace(prefix, uri)
        return tree, root
    
    def extract_text_nodes(self, root):
        # Get nodes that may contain translatable text
        text_nodes = (
            root.findall('.//text:h', self.namespaces) +
            root.findall('.//text:p', self.namespaces) +
            root.findall('.//text:span', self.namespaces)
        )
        texts_to_translate = []
        node_mapping = []  # list of tuples (node, 'text' or 'tail')
        for node in text_nodes:
            if node.text and node.text.strip():
                texts_to_translate.append(node.text)
                node_mapping.append((node, 'text'))
            for child in node:
                if child.text and child.text.strip():
                    texts_to_translate.append(child.text)
                    node_mapping.append((child, 'text'))
                if child.tail and child.tail.strip():
                    texts_to_translate.append(child.tail)
                    node_mapping.append((child, 'tail'))
        return texts_to_translate, node_mapping, text_nodes
    
    def translate_texts(self, texts):
        translated_texts = []
        for text in texts:
            try:
                response = self.client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": f"You are a translator. Translate the following text to {self.target_language}."},
                        {"role": "user", "content": text}
                    ],
                    temperature=0.3
                )
                translated = response.choices[0].message.content
                translated_texts.append(translated)
            except Exception as e:
                print(f"Error translating '{text}': {e}")
                translated_texts.append(text)
        return translated_texts
    
    def update_content(self, node_mapping, translated_texts):
        for (node, attr), translated_text in zip(node_mapping, translated_texts):
            if attr == 'text':
                node.text = translated_text
            elif attr == 'tail':
                node.tail = translated_text
    
    def add_rtl_formatting(self, root, text_nodes):
        # Only add RTL formatting if the target language is RTL
        if self.target_language not in RTL_LANGUAGES:
            return
        
        # Locate or create the <office:automatic-styles> element
        auto_styles = root.find('.//office:automatic-styles', self.namespaces)
        if auto_styles is None:
            auto_styles = etree.SubElement(root, '{urn:oasis:names:tc:opendocument:xmlns:office:1.0}automatic-styles')
        
        # Find or create the RTL style
        rtl_style = auto_styles.find(".//style:style[@style:name='RTL_Style']", self.namespaces)
        if rtl_style is None:
            rtl_style = etree.SubElement(auto_styles, '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style')
            rtl_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name', 'RTL_Style')
            rtl_style.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family', 'paragraph')
            para_props = etree.SubElement(rtl_style, '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}paragraph-properties')
            para_props.set('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}writing-mode', 'rl-tb')
            para_props.set('{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}text-align', 'end')
        
        # Apply RTL style to paragraph and heading elements, and optionally to spans
        for node in text_nodes:
            if node.tag in (
                '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}h',
                '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p'
            ):
                node.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name', 'RTL_Style')
            elif node.tag == '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span':
                node.set('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name', 'RTL_Style')
    
    def repackage_odt(self):
        # Package the modified files back into an .odt archive.
        base_name = "temp_output"
        shutil.make_archive(base_name, 'zip', self.extract_dir)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(base_name + '.zip', self.output_file)
    
    def cleanup(self):
        if os.path.exists(self.extract_dir):
            shutil.rmtree(self.extract_dir)
    
    def run(self):
        # Step 1: Extract the .odt file
        self.extract_odt()
        
        # Step 2: Parse content.xml
        tree, root = self.parse_content()
        
        # Step 3: Extract text nodes to translate
        texts_to_translate, node_mapping, text_nodes = self.extract_text_nodes(root)
        
        # Step 4: Translate texts using the LLM
        translated_texts = self.translate_texts(texts_to_translate)
        
        # Step 5: Update XML with the translated texts
        self.update_content(node_mapping, translated_texts)
        
        # Step 6: Add RTL formatting if the target language is RTL
        self.add_rtl_formatting(root, text_nodes)
        
        # Step 7: Save the modified XML back to content.xml
        tree.write(self.content_xml_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        
        # (Optional) Output the first few lines for inspection
        print("\nFirst 5 lines of modified content.xml:")
        with open(self.content_xml_path, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f, 1):
                print(f"Line {i}: {line.strip()}")
                if i >= 5:
                    break
        
        # Validate the modified XML
        try:
            etree.parse(self.content_xml_path, etree.XMLParser(remove_blank_text=True))
            print("XML validation successful!")
        except etree.ParseError as e:
            print(f"XML validation failed: {e}")
            raise
        
        # Step 8: Repackage the folder contents into a new .odt file
        self.repackage_odt()
        print(f"Translation complete! Saved as: {self.output_file}")
        
        # Cleanup the temporary extraction folder
        self.cleanup()


class DocTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        # OpenAI library should be preconfigured with your API key
        # For example: openai.api_key = os.environ.get("OPENAI_API_KEY")
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

    def extract_text_from_doc(self):
        """
        Extracts text from a .doc file using antiword.
        Returns the extracted text as a UTF-8 string.
        """
        try:
            # Run antiword to extract text from the document.
            result = subprocess.run(
                ['antiword', self.input_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            if result.returncode != 0:
                print(f"Error extracting text using antiword: {result.stderr}")
                return ""
            return result.stdout
        except Exception as e:
            print(f"Error extracting text: {e}")
            return ""

    def translate_text(self, text):
        """
        Translates text content paragraph by paragraph.
        For right-to-left (RTL) languages, a right-to-left mark is added.
        """
        paragraphs = text.split('\n')
        translated_paragraphs = []
        is_rtl = self.target_language in RTL_LANGUAGES

        for paragraph in paragraphs:
            if not paragraph.strip():
                translated_paragraphs.append("")
                continue

            try:
                response = self.client.ChatCompletion.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": f"You are a translator. Translate the following text to {self.target_language}."},
                        {"role": "user", "content": paragraph}
                    ],
                    temperature=0.3
                )
                translated = response.choices[0].message.content.strip()
                
                # If the target language is RTL, prepend the RTL mark
                if is_rtl:
                    translated = "\u200F" + translated
                
                translated_paragraphs.append(translated)
            except Exception as e:
                print(f"Translation error for paragraph '{paragraph}': {e}")
                translated_paragraphs.append(paragraph)

        return "\n".join(translated_paragraphs)

    def create_translated_doc(self, translated_text):
        """
        Creates a new .doc file (as a plain text file with a .doc extension)
        containing the translated text.
        """
        try:
            with open(self.output_file, 'w', encoding='utf-8') as f:
                f.write(translated_text)
            print(f"Translation complete! Saved as: {self.output_file}")
        except Exception as e:
            print(f"Error creating translated doc file: {e}")

    def run(self):
        text = self.extract_text_from_doc()
        print(f"Extracted text: {text[:100]}...")  # Debug: Check if text is extracted
        if not text:
            print("No text extracted from the document.")
            return
        translated_text = self.translate_text(text)
        print(f"Translated text: {translated_text[:100]}...")  # Debug: Check translation
        self.create_translated_doc(translated_text)



def translate_file(input_file, output_file, target_language):
    """
    Dispatch function that creates an instance of the appropriate translator class based on file extension.
    """
    file_extension = os.path.splitext(input_file)[1].lower()

    if file_extension == '.docx':
        translator = DocxTranslator(input_file, output_file, target_language)
    elif file_extension == '.txt':
        translator = TxtTranslator(input_file, output_file, target_language)
    elif file_extension == '.odt':
        translator = OdtTranslator(input_file, output_file, target_language)
    elif file_extension == '.doc':
        translator = DocTranslator(input_file, output_file, target_language)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx, .txt, .odt, .doc files")

    translator.run()

