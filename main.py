import streamlit as st
from dotenv import load_dotenv
from htmlTemplates import css, bot_template, user_template
import pdf_rag
import excel_rag
from pathlib import Path
import os

# For folder selection
from tkinter import Tk, filedialog

def handle_userinput(user_question):
    response = st.session_state.conversation({'question': user_question})
    st.session_state.chat_history = response['chat_history']

    for i, message in enumerate(st.session_state.chat_history):
        if i % 2 == 0:
            st.write(user_template.replace(
                "{{MSG}}", message.content), unsafe_allow_html=True)
        else:
            st.write(bot_template.replace(
                "{{MSG}}", message.content), unsafe_allow_html=True)

def main():
    load_dotenv()
    st.set_page_config(page_title="Chat with your Documents", page_icon=":books:")
    st.write(css, unsafe_allow_html=True)

    if "conversation" not in st.session_state:
        st.session_state.conversation = None
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = None

    st.header("Chat with your Documents :books:")
    user_question = st.text_input("Ask a question about your documents:")
    if user_question:
        handle_userinput(user_question)

    with st.sidebar:
        st.subheader("Your documents")

        if st.button("Select Folder"):
            folder_path = select_folder()
            if folder_path:
                st.write(f"Selected folder: {folder_path}")
                process_folder(folder_path)

def select_folder():
    # Hide Tkinter root window
    root = Tk()
    root.withdraw()
    root.call('wm', 'attributes', '.', '-topmost', True)
    folder_path = filedialog.askdirectory()
    root.destroy()
    return folder_path

def process_folder(folder_path):
    pdf_files = []
    excel_files = []

    for file in Path(folder_path).rglob('*'):
        if file.suffix.lower() == '.pdf':
            pdf_files.append(file)
        elif file.suffix.lower() in ['.xls', '.xlsx']:
            excel_files.append(file)

    if pdf_files:
        with st.spinner("Processing PDF documents..."):
            pdf_rag.process_pdfs([open(file, 'rb') for file in pdf_files])
    if excel_files:
        with st.spinner("Processing Excel documents..."):
            excel_rag.process_excels([open(file, 'rb') for file in excel_files])

if __name__ == '__main__':
    main()
