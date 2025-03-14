import os
import zipfile
import shutil
import tempfile
from xml.etree import ElementTree as ET
from lxml import etree
import re
import subprocess
from openai import OpenAI
import openai

RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish", 
    "Pashto", "Sindhi", "Dhivehi", "Kurdish"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        # Use a unique temporary folder for multi-user handling.
        self.extract_folder = tempfile.mkdtemp(prefix="docx_extract_")
        self.word_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        # Configure the OpenAI client (make sure OPENAI_API_KEY is set in your environment)
        openai.api_key = os.environ.get("OPENAI_API_KEY")
        self.client = openai

    def extract_docx(self):
        """Extract DOCX contents to a temporary folder and return the document.xml path."""
        with zipfile.ZipFile(self.input_file, 'r') as zip_ref:
            zip_ref.extractall(self.extract_folder)
        return os.path.join(self.extract_folder, "word", "document.xml")

    def create_translated_docx(self):
        """Create a new DOCX file from the modified extracted contents."""
        base_name = self.output_file.replace('.docx', '')
        shutil.make_archive(base_name, 'zip', self.extract_folder)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(base_name + '.zip', self.output_file)

    def translate_text(self, text):
        """Translate a piece of text using the OpenAI API."""
        try:
            response = self.client.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": f"You are a translator. Translate the following text to {self.target_language}."},
                    {"role": "user", "content": text}
                ],
                temperature=0.3
            )
            translated_text = response.choices[0].message.content.strip()
            return translated_text
        except Exception as e:
            print(f"Translation error: {e}")
            # In case of an error, return the original text to avoid data loss.
            return text

    def translate_xml_to_language(self, xml_path):
        """Translate all text nodes in the document.xml while preserving XML structure and styles."""
        tree = ET.parse(xml_path)
        root = tree.getroot()
        ET.register_namespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
        is_rtl = self.target_language in RTL_LANGUAGES

        # Process each text element (<w:t>) individually
        for text_elem in root.findall(f'.//{self.word_ns}t'):
            if text_elem.text and text_elem.text.strip():
                original_text = text_elem.text.strip()
                translated_text = self.translate_text(original_text)
                text_elem.text = translated_text

        # If the target language is RTL, adjust paragraph properties accordingly.
        if is_rtl:
            for paragraph in root.findall(f'.//{self.word_ns}p'):
                pPr = paragraph.find(f'{self.word_ns}pPr')
                if pPr is None:
                    pPr = ET.SubElement(paragraph, f'{self.word_ns}pPr')
                bidi = pPr.find(f'{self.word_ns}bidi')
                if bidi is None:
                    bidi = ET.SubElement(pPr, f'{self.word_ns}bidi')
                bidi.set(f'{self.word_ns}val', 'on')

        # Write back the modified XML
        tree.write(xml_path, encoding='utf-8', xml_declaration=True)

    def run(self):
        # Ensure any pre-existing temporary folder is not used.
        if os.path.exists(self.extract_folder):
            try:
                shutil.rmtree(self.extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing folder: {e}")
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
                    print(f"Warning: Could not clean up temporary folder: {e}")




class OdtTranslator:
    def __init__(self, input_file, output_file, target_language):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        # Use a unique temporary directory for multi-user handling.
        self.extract_dir = tempfile.mkdtemp(prefix="odt_extract_")
        self.content_xml_path = os.path.join(self.extract_dir, "content.xml")
        self.namespaces = {
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'
        }
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
        # Get nodes that may contain translatable text.
        text_nodes = (
            root.findall('.//text:h', self.namespaces) +
            root.findall('.//text:p', self.namespaces) +
            root.findall('.//text:span', self.namespaces)
        )
        texts_to_translate = []
        node_mapping = []  # list of tuples (node, attribute 'text' or 'tail')
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
        base_name = self.output_file.replace('.odt', '')
        shutil.make_archive(base_name, 'zip', self.extract_dir)
        if os.path.exists(self.output_file):
            os.remove(self.output_file)
        os.rename(base_name + '.zip', self.output_file)
    
    def cleanup(self):
        if os.path.exists(self.extract_dir):
            shutil.rmtree(self.extract_dir)
    
    def run(self):
        # Step 1: Extract the .odt file into a unique temporary directory.
        self.extract_odt()
        
        # Step 2: Parse content.xml from the extracted folder.
        tree, root = self.parse_content()
        
        # Step 3: Extract text nodes to translate.
        texts_to_translate, node_mapping, text_nodes = self.extract_text_nodes(root)
        
        # Step 4: Translate texts using the LLM.
        translated_texts = self.translate_texts(texts_to_translate)
        
        # Step 5: Update XML with the translated texts.
        self.update_content(node_mapping, translated_texts)
        
        # Step 6: Add RTL formatting if the target language is RTL.
        self.add_rtl_formatting(root, text_nodes)
        
        # Step 7: Save the updated XML.
        # This replaces the original content.xml file in the extracted folder with the new translated version.
        tree.write(self.content_xml_path, encoding='utf-8', xml_declaration=True, pretty_print=True)
        
        # (Optional) Output the first few lines for inspection.
        print("\nFirst 5 lines of modified content.xml:")
        with open(self.content_xml_path, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f, 1):
                print(f"Line {i}: {line.strip()}")
                if i >= 5:
                    break
        
        # Validate the modified XML.
        try:
            etree.parse(self.content_xml_path, etree.XMLParser(remove_blank_text=True))
            print("XML validation successful!")
        except etree.ParseError as e:
            print(f"XML validation failed: {e}")
            raise
        
        # Step 8: Repackage the folder contents into a new .odt file.
        self.repackage_odt()
        print(f"Translation complete! Saved as: {self.output_file}")
        
        # Cleanup the temporary extraction folder.
        self.cleanup()
