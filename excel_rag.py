# excel_rag.py

import os
import pandas as pd
from dotenv import load_dotenv
import nest_asyncio
from llama_parse import LlamaParse
from llama_index.llms.openai import OpenAI
from llama_index.core import Settings, VectorStoreIndex
from langchain.memory import ConversationBufferMemory
from langchain.chains import ConversationalRetrievalChain
import streamlit as st

nest_asyncio.apply()

load_dotenv()  # Load environment variables from .env file

def process_excels(excel_files):
    api_key = os.getenv("LLAMA_CLOUD_API_KEY")
    openai_api_key = os.getenv("OPENAI_API_KEY")

    # Ensure the API key is available
    if not api_key or not openai_api_key:
        st.error("API keys for LLAMA and OpenAI must be set in the .env file.")
        return

    # Initialize LlamaParse with the API key
    parser = LlamaParse(api_key=api_key, result_type="markdown")

    documents = []
    for excel_file in excel_files:
        # Read the content of the Excel file
        excel_content = excel_file.read()

        # Convert the content to a DataFrame
        df = pd.read_excel(excel_file, engine='openpyxl')  # use the correct engine

        # Convert the DataFrame to a string format
        df_string = df.to_string()

        # Load data using the parser
        documents.extend(parser.load_data(df_string))

    # Initialize OpenAI with the OpenAI API key
    llm = OpenAI(api_key=openai_api_key, model="gpt-4")
    Settings.llm = llm

    # Create the index and query engine
    index = VectorStoreIndex.from_documents(documents)
    query_engine = index.as_query_engine()

    # Initialize ConversationBufferMemory
    memory = ConversationBufferMemory(memory_key='chat_history', return_messages=True)
    conversation_chain = ConversationalRetrievalChain(
        llm=llm,
        retriever=index.as_retriever(),
        memory=memory
    )

    # Store the conversation chain in Streamlit session state
    st.session_state.conversation = conversation_chain
