import streamlit as st, sys, platform
st.set_page_config(page_title="Smoke Test", layout="wide")
st.title("✅ Streamlit no ar!")
st.write("Python:", sys.version)
st.write("SO:", platform.platform())
st.success("Se você está vendo esta página, o deploy está OK. Próximo passo: voltar o app completo.")
