import streamlit as st
import pandas as pd
import os
import requests
import matplotlib.pyplot as plt
import seaborn as sns
import io
import textwrap
import sys

# Streamlit page setup
st.set_page_config(page_title="Excel Chatbot", layout="wide")
st.title("📊 Insight Sheet")

# Load Mistral API key
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
        prompt = f"""
You are a Python data analyst. Given a pandas DataFrame `df`, write Python code to answer the user's question below.

Requirements:
- Only return executable Python code (no explanations)
- DO NOT use input() or print()
- If the result is a number, assign it to a variable named `result`
- If it’s a DataFrame, assign it to `result`
- If it’s a chart, use matplotlib/seaborn to plot

User Question: "{user_input}"

Data Preview:
{df.head(5).to_markdown(index=False)}

Python code:
"""

        # Request to Mistral API
        response = requests.post(
            "https://api.mistral.ai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": "mistral-small",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0
            }
        )

        if response.status_code == 200:
            result_text = response.json()["choices"][0]["message"]["content"]

            # Clean and dedent code
            clean_code = (
                result_text.replace("\\_", "_")
                .replace("```python", "")
                .replace("```", "")
                .strip()
            )
            clean_code = textwrap.dedent(clean_code)

            # Prepare execution environment
            local_vars = {
                'df': df,
                'st': st,
                'pd': pd,
                'plt': plt,
                'sns': sns
            }

            try:
                exec(clean_code, {}, local_vars)

                # Show result
                if "result" in local_vars:
                    result = local_vars["result"]
                    if isinstance(result, pd.DataFrame):
                        st.write("### Result:")
                        st.dataframe(result)
                    else:
                        st.success(f"✅ Answer: {result}")
                elif plt.get_fignums():
                    st.pyplot(plt.gcf())
                    plt.clf()
                else:
                    st.warning("Code executed but no output was returned.")
            except Exception as e:
                st.error(f"❌ Error while executing generated code:\n{str(e)}")

        else:
            st.error(f"❌ API Error {response.status_code}: {response.text}")
