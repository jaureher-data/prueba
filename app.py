import os
from pathlib import Path

# Crear secrets.toml dinámicamente
def crear_secrets_streamlit():

    Path(".streamlit").mkdir(exist_ok=True)

    secrets = f"""
[auth]
redirect_uri = "{os.environ["AUTH_REDIRECT_URI"]}"
cookie_secret = "{os.environ["AUTH_COOKIE_SECRET"]}"

[auth.google]
client_id = "{os.environ["GOOGLE_CLIENT_ID"]}"
client_secret = "{os.environ["GOOGLE_CLIENT_SECRET"]}"
server_metadata_url = "https://accounts.google.com/.well-known/openid-configuration"
"""

    Path(".streamlit/secrets.toml").write_text(
        secrets.strip(),
        encoding="utf-8"
    )

crear_secrets_streamlit()

# IMPORTANTE:
# Streamlit se importa DESPUÉS
import streamlit as st

st.title("App protegida")

if not st.user.is_logged_in:

    if st.button("Login Google"):
        st.login("google")

    st.stop()

st.success(f"Bienvenido {st.user.email}")
