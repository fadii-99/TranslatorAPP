import asyncio
import subprocess
import os
import pypandoc
from googletrans import Translator
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def rtf_to_docx(input_rtf, output_docx):
    """
    Convert an RTF file to DOCX using Pandoc.
    Ensure that Pandoc is installed on your system.
    """
    pypandoc.convert_file(input_rtf, 'docx', outputfile=output_docx)
    print(f"Converted {input_rtf} to {output_docx}")

def docx_to_rtf_libreoffice(input_docx, output_rtf):
    """
    Convert a DOCX file to RTF using LibreOffice in headless mode.
    LibreOffice usually handles RTL formatting better than Pandoc.
    """
    subprocess.run(["libreoffice", "--headless", "--convert-to", "rtf", input_docx, "--outdir", "."])
    # LibreOffice names the output file based on the DOCX file name.
    base = os.path.splitext(os.path.basename(input_docx))[0] + ".rtf"
    if os.path.exists(base):
        os.rename(base, output_rtf)
        print(f"Converted {input_docx} to {output_rtf} using LibreOffice")
    else:
        print("Conversion failed: Output RTF file not found.")

async def translate_paragraph(translator, text, src='en', dest='ar'):
    """
    Asynchronously translate a single paragraph's text.
    """
    try:
        # Await the translation coroutine.
        translation = await translator.translate(text, src=src, dest=dest)
        return translation.text
    except Exception as e:
        print("Translation error:", e)
        return text

def set_paragraph_rtl(paragraph):
    """
    Sets right-to-left (RTL) formatting for a paragraph by modifying its XML.
    This adds the <w:bidi w:val="1"/> element to the paragraph properties.
    """
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

async def translate_docx_async(input_docx, output_docx, src='en', dest='ar'):
    """
    Asynchronously translates each paragraph in a DOCX file.
    If the destination language is RTL (like Arabic), the code applies RTL formatting.
    """
    translator = Translator()
    doc = Document(input_docx)
    tasks = []

    # Create translation tasks for each non-empty paragraph.
    for para in doc.paragraphs:
        original_text = para.text.strip()
        if original_text:
            tasks.append(translate_paragraph(translator, original_text, src, dest))
        else:
            tasks.append(asyncio.sleep(0, result=""))

    # Run all translation tasks concurrently.
    translated_texts = await asyncio.gather(*tasks)

    # List of RTL language codes; adjust as needed.
    rtl_languages = ['ar', 'he', 'fa', 'ur']
    is_rtl = dest.lower() in rtl_languages

    # Update paragraphs with translated text and apply RTL formatting if necessary.
    for para, translated_text in zip(doc.paragraphs, translated_texts):
        if translated_text:
            para.text = translated_text
            if is_rtl:
                set_paragraph_rtl(para)

    doc.save(output_docx)
    print(f"Translated content saved to {output_docx}")

def main():
    input_rtf = 'input.rtf'                    # Your original RTF file
    intermediate_docx = 'intermediate.docx'    # DOCX after converting from RTF
    translated_docx = 'translated.docx'        # DOCX after translation with RTL formatting
    output_rtf = 'translated.rtf'              # Final output RTF file

    # Step 1: Convert the original RTF to DOCX.
    rtf_to_docx(input_rtf, intermediate_docx)

    # Step 2: Translate the DOCX content asynchronously (English → Arabic)
    #         and apply RTL formatting.
    asyncio.run(translate_docx_async(intermediate_docx, translated_docx, src='en', dest='ar'))

    # Step 3: Convert the translated DOCX back to RTF using LibreOffice.
    docx_to_rtf_libreoffice(translated_docx, output_rtf)

if __name__ == "__main__":
    main()
