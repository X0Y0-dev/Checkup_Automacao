import streamlit as st
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
logo_path = BASE_DIR / "assets" / "logo_hsl.png"

PAGINAS = {
    "Home": "app.py",
    "Elegibilidade": "pages/elegibilidade.py",
}

# HEADER PADRONIZADO
def render_header(text):

    # estilização do header
    with st.container():
        st.markdown('<style>div.block-container{padding-top:2rem;}</style>',unsafe_allow_html = True)
        st.markdown('<div class = "header">', unsafe_allow_html = True)
        col1, col2, col3 = st.columns([1,3,1])
        
        # Logo
        with col1:
            st.image(BASE_DIR / "pages" / "assets" / "logo_hsl.png", width=150)

        # Título da página
        with col2:
            st.markdown(f"""
                <div style = 'display: flex; align-items: center; height: 100%; padding-top: 20px;'>
                    <h1 style = 'text-align: center; width: 100%; margin: 0'>{text}</h1>
                </div>
            """, unsafe_allow_html=True)

# SIDEBAR PADRONIZADA
def render_sidebar():
    # inicia em home
    if "menu_ativo" not in st.session_state:
        st.session_state.menu_ativo = "Home"
    
    # logica pra muda o menu
    def mudar_menu(menu):
        st.session_state.menu_ativo = menu

        if menu == "Home":
            st.switch_page("app.py")

        st.switch_page(PAGINAS[menu])
    
    # estilização da sidebar
    st.sidebar.markdown("""
    <style>
        section[data-testid = "stSidebar"] > div {
            background-color: #8ddbe8;
            padding: 20px;
        }

        div.stButton > button {
            width: 100%;
            height: 80px;
            border-radius: 10px;
            margin-bottom: 10px;
            transition: all 0.2s ease;
        }

        div.stButton > button[kind = "primary"] {
            background-color: #05437c;
            color: white;
            border: 1px solid #000000;
        }

        div.stButton > button[kind = "secondary"] {
            background-color: rgba(5, 67, 124, 0.35);
            color: gray;
            border: 1px solid #05437c;
        }

        div.stButton > button[kind = "secondary"]:hover {
            background-color: rgba(5, 67, 124, 0.75);
        }

        [data-testid = "stSidebarNav"] {
        display: none;
        }
    </style>
    """, unsafe_allow_html=True)

    # botooes que piscam diferente :D
    for menu in PAGINAS.keys():
        if st.sidebar.button(menu,use_container_width=True,
            type = "primary" if st.session_state.menu_ativo == menu else "secondary",
        ):
            st.session_state.menu_ativo = menu
            st.switch_page(PAGINAS[menu])