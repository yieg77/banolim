import streamlit as st

st.title("Hello, Streamlit!")
st.write("This is a simple Streamlit app to test GitHub and Streamlit deployment.")

name = st.text_input("What's your name")
if name :
    st.write(f"It's you, {name}! 👋")