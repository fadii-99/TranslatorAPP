import os
import streamlit as st
import pandas as pd
import io
from langchain_community.document_loaders import PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from dotenv import load_dotenv
from utils import translate_file  # Import DOCX translation function from your provided utils.py

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")

# Function to process and translate PDF
def process_pdf(file_path, selected_model, target_langs):
    loader = PyPDFLoader(file_path)
    docs = loader.load()

    text_splitter = RecursiveCharacterTextSplitter(chunk_size=512, chunk_overlap=0)
    splits = text_splitter.split_documents(docs)

    prompt_translation = ChatPromptTemplate.from_messages(
        [
            (
                "system",
                """
                Your task is to translate the provided text from English 
                to the target language while maintaining its original meaning and context.
                Ensure the translation is accurate, fluent, and culturally appropriate. 
                Return the translated text only.
                target language: {target}
                source text: {text}
                """,
            ),
        ]
    )

    llm = ChatOpenAI(model=selected_model, temperature=0.01)
    runnable_translation = prompt_translation | llm

    def get_translation(text, target):
        return runnable_translation.invoke({"target": target, "text": text}).content

    results = {"ENGLISH": []}
    for lang in target_langs:
        results[lang.upper()] = []

    for split in splits:
        results["ENGLISH"].append(split.page_content)
        for lang in target_langs:
            results[lang.upper()].append(get_translation(split.page_content, lang))

    return pd.DataFrame(results)

# Streamlit UI
st.set_page_config(page_title="Document Translator", layout="centered")
st.title("Document Translator with GPT-4o")

# formats = ["TXT"]

# Step 1: User selects file format
# file_format = st.selectbox("Choose File Format to Upload:", formats,   index=0 )
file_format = st.selectbox("Choose File Format to Upload:", ["PDF", "DOCX", "TXT"])

# Step 2: File upload based on selected format
if file_format == "PDF":
    uploaded_file = st.file_uploader(f"Upload your {file_format} file", type=["pdf"])
    file_extension = "pdf" if uploaded_file else None
elif file_format == "DOCX":
    uploaded_file = st.file_uploader(f"Upload your {file_format} file", type=["docx"])
    file_extension = "docx" if uploaded_file else None
if file_format == "TXT":
    uploaded_file = st.file_uploader(f"Upload your {file_format} file", type=["txt"])
    file_extension = "txt" if uploaded_file else None

# Model selection
selected_model = st.selectbox("Select Translation Model:", ["gpt-4o", "gpt-4o-mini"])

# Fixed source language
st.info("Source language: English (Fixed)")

# Target language selection
languages = ["French", "Arabic", "Spanish", "German"]
target_langs = []

if uploaded_file:
    if file_extension == "pdf":
        target_langs = st.multiselect("Select Target Languages:", languages, default=["French", "Arabic"])
    elif file_extension == "docx":
        target_lang = st.selectbox("Select Target Language (for DOCX):", languages, index=1)  # Default to Arabic
        target_langs = [target_lang] if target_lang else []
    elif file_extension == "txt":
        target_lang = st.selectbox("Select Target Language (for TXT):", languages, index=1)  # Default to Arabic
        target_langs = [target_lang] if target_lang else []
else:
    st.write(f"Please upload a {file_format} file to select target languages.")

# Add a Translate button to initiate translation
if uploaded_file and selected_model and target_langs:
    if st.button("Translate"):
        st.success(f"File uploaded successfully! Processing {file_format} file...")
        
        # Save uploaded file temporarily
        input_file_path = f"uploaded_file.{file_extension}"
        with open(input_file_path, "wb") as f:
            f.write(uploaded_file.read())

        # Start translation process with loading spinner
        with st.spinner(f"Translating {file_format} to {', '.join(target_langs)}..."):
            if file_extension == "pdf":
                # Process PDF
                df_results = process_pdf(input_file_path, selected_model, [lang.lower() for lang in target_langs])

                # Show preview of translated data
                st.write("Translation Preview:")
                st.dataframe(df_results.head(20))

                # Provide download button for Excel file
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                    df_results.to_excel(writer, index=False)
                excel_buffer.seek(0)

                st.download_button(
                    label="Download Translations as Excel",
                    data=excel_buffer,
                    file_name="translated_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            elif file_extension == "docx":
                # Process DOCX
                output_file_path = f"translated_document_{target_langs[0].lower()}.docx"
                translate_file(input_file_path, output_file_path, target_langs[0])  # Use function from utils.py

                # Provide download button for translated DOCX
                with open(output_file_path, "rb") as f:
                    st.download_button(
                        label=f"Download Translated DOCX ({target_langs[0]})",
                        data=f,
                        file_name=f"translated_document_{target_langs[0].lower()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

            elif file_extension == "txt":
                output_file_path = f"translated_document_{target_langs[0].lower()}.txt"
                translate_file(input_file_path, output_file_path, target_langs[0])  # Use function from utils.py

                # Provide download button for translated DOCX
                with open(output_file_path, "rb") as f:
                    st.download_button(
                        label=f"Download Translated TXT ({target_langs[0]})",
                        data=f,
                        file_name=f"translated_document_{target_langs[0].lower()}.txt",
                        mime="text/plain",
                    )

        # Spinner stops here, and success message appears
        st.success("Translation completed and file is ready for download.")

        # Clean up temporary files
        if os.path.exists(input_file_path):
            os.remove(input_file_path)
        if file_extension == "docx" and os.path.exists(output_file_path):
            os.remove(output_file_path)
else:
    if uploaded_file and not target_langs:
        st.warning("Please select a target language and click 'Translate' to proceed.")
    elif uploaded_file:
        st.info("Select a target language and click 'Translate' to start.")