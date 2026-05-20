import streamlit as st

st.set_page_config(page_title="App con login Gmail")

if not st.user.is_logged_in:
    st.title("Inicio de sesión")
    st.write("Debes iniciar sesión con tu cuenta de Google.")
    st.login("google")
    st.stop()

st.sidebar.write(f"Usuario: {st.user.email}")

if st.sidebar.button("Cerrar sesión"):
    st.logout()

st.title("Aplicación protegida")
st.success(f"Bienvenido, {st.user.name}")

st.write("Contenido privado de la app.")
