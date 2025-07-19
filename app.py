import streamlit as st
import pandas as pd
import os
import requests
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set Streamlit layout
st.set_page_config(page_title="Excel Chatbot", layout="wide")
st.title("📊 Insight Sheet")

# Load API key
MISTRAL_API_KEY = os.getenv("MISTRAL_API_KEY")
if not MISTRAL_API_KEY:
    st.warning("Please set your Mistral API token as the 'MISTRAL_API_KEY' environment variable.")
    st.stop()

# Upload Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Data Preview:")
    st.dataframe(df)

    # Ask user question
    user_input = st.text_input("Ask a question about your Excel data:")

    if user_input:
        # Prompt to Mistral: ask for Python code only
        prompt = f"""
You are a data analyst. Given the following DataFrame (called `df`), generate only the Python code to answer the user's question.
Do not explain the answer. Only return code.
Do not use input() or print().
Only use Pandas/Matplotlib/Seaborn. Assume `df` is already loaded.
User Question: "{user_input}"
Data Preview:
{df.head(5).to_markdown(index=False)}

Python code:
"""

        # Send request to Mistral
        response = requests.post(
            "https://api.mistral.ai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": "mistral-small",  # better at reasoning
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0
            }
        )

        if response.status_code == 200:
            result_text = response.json()["choices"][0]["message"]["content"]
            st.code(result_text.strip(), language="python")

            try:
                # Redirect stdout (for capturing print if needed)
                output_buffer = io.StringIO()

                # Execute code safely in local scope
                local_vars = {'df': df, 'plt': plt, 'sns': sns, 'pd': pd, 'st': st}
                # Clean up escaped underscores
                result_text_clean = result_text.replace("\\_", "_").strip()

                exec(result_text, {}, local_vars)

                # If plot was created, show it
                if plt.get_fignums():
                    st.pyplot(plt.gcf())
                    plt.clf()

                # Show textual output if any variable was created
                for key, val in local_vars.items():
                    if isinstance(val, pd.DataFrame):
                        st.write(f"### Result DataFrame: `{key}`")
                        st.dataframe(val)
            except Exception as e:
                st.error(f"❌ Error while executing generated code:\n{str(e)}")
        else:
            st.error(f"API Error {response.status_code}: {response.text}")
