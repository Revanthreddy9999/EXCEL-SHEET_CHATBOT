import streamlit as st
import pandas as pd
import os
import requests
import matplotlib.pyplot as plt
import seaborn as sns
import io
import textwrap

# Streamlit page setup
st.set_page_config(page_title="Excel Chatbot", layout="wide")
st.title("📊 Insight Sheet")

# Load Mistral API key from environment
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
        # Prepare Mistral prompt
        prompt = f"""
You are a data analyst. Given the following DataFrame named `df`, generate Python code to answer the user's question.

Requirements:
- ONLY return valid Python code (NO explanations or text)
- Use pandas, matplotlib, seaborn, or Streamlit as needed
- Assume df is already defined
- DO NOT use print() or input()

User Question: "{user_input}"

Data Preview (first 5 rows):
{df.head(5).to_markdown(index=False)}

Python code:
"""

        # Call Mistral API
        response = requests.post(
            "https://api.mistral.ai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": "mistral-small",  # Use mistral-medium if available
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0
            }
        )

        if response.status_code == 200:
            result_text = response.json()["choices"][0]["message"]["content"]

            # Clean and sanitize code
            clean_code = (
                result_text
                .replace("\\_", "_")           # Remove markdown escaped underscores
                .replace("```python", "")
                .replace("```", "")
                .strip()
            )
            clean_code = textwrap.dedent(clean_code)  # Normalize indentation

            # Show code to user
            st.code(clean_code, language="python")

            try:
                local_vars = {
                    'df': df,
                    'st': st,
                    'pd': pd,
                    'plt': plt,
                    'sns': sns
                }

                # Simple safety check
                unsafe_keywords = ["import os", "open(", "eval(", "exec(", "subprocess", "system"]
                if any(keyword in clean_code for keyword in unsafe_keywords):
                    st.error("⚠️ Unsafe code detected. Execution aborted.")
                else:
                    # Run the code
                    exec(clean_code, {}, local_vars)

                    # Show plot if created
                    if plt.get_fignums():
                        st.pyplot(plt.gcf())
                        plt.clf()

                    # Show any result DataFrame created
                    for key, val in local_vars.items():
                        if isinstance(val, pd.DataFrame) and key != "df":
                            st.write(f"### Result DataFrame: `{key}`")
                            st.dataframe(val)

            except Exception as e:
                st.error(f"❌ Error while executing generated code:\n{str(e)}")

        else:
            st.error(f"❌ API Error {response.status_code}: {response.text}")
