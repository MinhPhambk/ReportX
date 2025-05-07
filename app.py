import os
from dotenv import load_dotenv
import pandas as pd
import streamlit as st
from pandasai import SmartDatalake
from pandasai.llm.openai import OpenAI
from pandasai.middlewares import StreamlitMiddleware
from pandasai.responses.streamlit_response import StreamlitResponse

# Loading Environment Variables Using the `dotenv` Package
load_dotenv()


if __name__ == "__main__":
    st.set_page_config(
        layout="wide",
        page_icon="./image/logo.svg",
        page_title="Chat and Automate Report",
    )
    st.title("Chat and Automate Report")

    # Sidebar for API Key settings
    with st.sidebar:
        st.header(
            "Set your API Key",
            help="You can get it from [OpenAI](https://platform.openai.com/account/api-keys/), or buy it conveniently from [here](https://api.nextweb.fun/).",
        )

        openai_api_base = os.getenv("OPENAI_API_BASE", "")
        openai_api_key = os.getenv("OPENAI_API_KEY", "")

        # Allow override manually if needed
        openai_api_base = st.text_input(
            "Enter API Base", value=openai_api_base
        ) or openai_api_base

        openai_api_key = st.text_input(
            "Enter API Key", value=openai_api_key, type="password"
        ) or openai_api_key

        # Create llm instance
        llm = OpenAI(api_token=openai_api_key)
        llm.api_base = openai_api_base

    with st.container():
        input_files = st.file_uploader(
            "Upload files", type=["xlsx", "csv"], accept_multiple_files=True
        )

        if len(input_files) > 0:
            data_list = []
            for input_file in input_files:
                if input_file.name.lower().endswith(".csv"):
                    data = pd.read_csv(input_file)
                else:
                    # Load Excel file with all formatting preserved
                    data = pd.read_excel(input_file, engine='openpyxl')
                
                # Display the dataframe as it is
                st.write(f"Data from file: {input_file.name}")
                st.dataframe(data, use_container_width=True)  # This shows the content of the file

                # Allow the user to download the original file back
                st.download_button(
                    label=f"Download original {input_file.name}",
                    data=input_file,
                    file_name=input_file.name,
                    mime="application/octet-stream",
                )
                
                data_list.append(data)
        else:
            st.header("Example Data")
            data = pd.read_excel("./Sample.xlsx")
            st.dataframe(data, use_container_width=True)
            data_list = [data]

        # Create SmartDatalake instance
        df = SmartDatalake(
            dfs=data_list,
            config={
                "llm": llm,
                "verbose": True,
                "response_parser": StreamlitResponse,
                "middlewares": [StreamlitMiddleware()],
            },
        )

        # Input text
        st.header("Ask anything!")
        input_text = st.text_area(
            "Enter your question", value="What is the total profit for each country?"
        )

        if input_text is not None:
            if st.button("Start Execution"):
                result = df.chat(input_text)

                # Display the result and code in two columns
                col1, col2 = st.columns(2)

                with col1:
                    # Display the result
                    st.header("Answer")
                    st.write(result)
