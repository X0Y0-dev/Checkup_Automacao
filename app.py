import streamlit as st
from src.menu import render_header, render_sidebar

render_header("Tela inicial")
render_sidebar()

st.title("Automação Check-up")
st.write("Bem-vindos ao sistema de automação de processos administrativos do Check-up. Aqui, você pode acessar uma variedade de ferramentas e recursos para facilitar suas tarefas diárias. Explore as opções disponíveis no menu lateral para começar a otimizar seus processos administrativos de forma eficiente e eficaz.")
st.write("As ferramentas atualmente disponíveis são:")
st.markdown("- Solitiação de Elegibilidade")