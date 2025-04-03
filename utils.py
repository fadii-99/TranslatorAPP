import os
import zipfile
import shutil
import tempfile
from langchain_openai import ChatOpenAI 
from lxml import etree
from dotenv import load_dotenv
from lxml import etree
from modernmt import ModernMT
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate, PromptTemplate


load_dotenv()

# ModernMT_key = os.environ.get("ModernMT_key")

from modernmt import ModernMT


RTL_LANGUAGES = {
    "Arabic", "Hebrew", "Persian", "Urdu", "Yiddish", 
    "Pashto", "Sindhi", "Dhivehi", "Kurdish"
}

class DocxTranslator:
    def __init__(self, input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY):
        self.input_file = input_file
        self.output_file = output_file
        self.target_language = target_language
        self.source_lang = 'English'
        self.extract_folder = tempfile.mkdtemp(prefix="docx_extract_")
        self.word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        if ModernMT_key:
            self.mmt = ModernMT(ModernMT_key)
            self.OPENAI_API_KEY = None  # Do not initialize OpenAI client
        elif OPENAI_API_KEY:
            self.mmt = None  # Do not initialize ModernMT
            self.OPENAI_API_KEY = OPENAI_API_KEY
        else:
            raise ValueError("Either ModernMT_key or OPENAI_API_KEY must be provided")

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
        if self.OPENAI_API_KEY:  # Using OpenAI API for translation
            prompt = PromptTemplate.from_template(
                    """
                    You are tasked with translating the given sentence from {source_language} into {output_language}. 
                    Your translation must be precise, accurate, and natural-sounding in the target language, preserving the original meaning without adding or omitting information. 
                    Avoid overly literal translations and instead focus on conveying the correct context clearly and concisely.

                    Original sentence ({source_language}):
                    "{input}"

                    Translated sentence ({output_language}):
                    """
                )

            llm = ChatOpenAI(model_name="gpt-4o", api_key=self.OPENAI_API_KEY)
            chain = prompt | llm

            try:
                target_lang = ''
                if self.target_language == "ar":
                    target_lang = "Arabic"
                response = chain.invoke(
                    {
                        "source_language": self.source_lang,
                        "output_language": target_lang,
                        "input": text,
                    }
                )
                print(f"response:  {response}")
                translated_text = response.content.strip()

                return translated_text
            except Exception as e:
                print(f"Translation error: {e}")
                return text  # Return original text in case of error

        elif self.mmt:  # Using ModernMT API for translation
            try:
                translation = self.mmt.translate("en", self.target_language, text)
                translated_text = translation.translation
                return translated_text
            except Exception as e:
                print(f"Error translating text with ModernMT: {e}")
                return text  # Return original text in case of error


    def translate_xml_to_language(self, xml_path, source_lang="en", target_lang="ar", output_path=None):
        # Parse the XML file
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(xml_path, parser)
        xml_string = etree.tostring(tree, encoding='unicode', pretty_print=True)

        # print("tags",xml_string[:50])
        
        # Parse the xml_string into an element tree
        root = etree.fromstring(xml_string)

        # print(f"Translated to: {root}") # Debug:
        
        # Iterate through the XML tree tag by tag
        for element in root.iter():
            if element.text and element.text.strip():
                original_text = element.text.strip()
                try:
                    original_text = element.text.strip()
                    translated_text = self.translate_text(original_text)
                    print(f"Translated: {original_text} -> \n\n from langchain  \n{translated_text}")  # Debug: show translated text.
                    element.text = translated_text
                    
                except Exception as e:
                    print(f"Error translating text: {e}")
            
            if element.tail and element.tail.strip():
                original_tail = element.tail.strip()
                
                try:
                    translated_tail = self.translate_text(original_tail)
                    
                    element.tail = translated_tail
                    
                except Exception as e:
                    print(f"Error translating tail text : {e}")

        if target_lang.lower() == 'ar':
        # Iterate over all paragraph elements (w:p)
            for p in root.findall('.//w:p', namespaces={'w': self.word_ns}):
                # Get or create paragraph properties (w:pPr)
                pPr = p.find('w:pPr', namespaces={'w': self.word_ns})
                if pPr is None:
                    pPr = etree.SubElement(p, '{%s}pPr' % self.word_ns)
                # Add the bidi element if it's not already present
                if pPr.find('{%s}bidi' % self.word_ns) is None:
                    etree.SubElement(pPr, '{%s}bidi' % self.word_ns)

        # Save the translated XML if output path is provided
        if output_path:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(etree.tostring(root, encoding='unicode', pretty_print=True))
        else:
            print(etree.tostring(root, encoding='unicode', pretty_print=True))

    def run(self):
        # Clean up any existing temporary folder.
        if os.path.exists(self.extract_folder):
            try:
                shutil.rmtree(self.extract_folder)
            except Exception as e:
                print(f"Warning: Could not remove existing folder: {e}")
        os.makedirs(self.extract_folder, exist_ok=True)

        try:
            xml_path = self.extract_docx()
            self.translate_xml_to_language(xml_path, output_path=xml_path)

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




def translate_file(input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY):
    """
    Dispatch function that creates an instance of the appropriate translator class based on file extension.
    """
    file_extension = os.path.splitext(input_file)[1].lower()

    if file_extension == '.docx':
        translator = DocxTranslator(input_file, output_file, target_language, ModernMT_key, OPENAI_API_KEY)
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Please use .docx")

    translator.run()