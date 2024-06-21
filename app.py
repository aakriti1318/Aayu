# app.py

import streamlit as st
from dotenv import load_dotenv
from htmlTemplates import css, bot_template, user_template
import pdf_rag
import excel_rag

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
        uploaded_files = st.file_uploader("Upload your documents here and click on 'Process'", accept_multiple_files=True)
        if st.button("Process"):
            if uploaded_files:
                pdf_files = [file for file in uploaded_files if file.name.endswith('.pdf')]
                excel_files = [file for file in uploaded_files if file.name.endswith(('.xls', '.xlsx'))]

                if pdf_files:
                    with st.spinner("Processing PDF documents..."):
                        pdf_rag.process_pdfs(pdf_files)
                if excel_files:
                    with st.spinner("Processing Excel documents..."):
                        excel_rag.process_excels(excel_files)

if __name__ == '__main__':
    main()
