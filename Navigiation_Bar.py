import streamlit as st 
from streamlit_option_menu import option_menu

def nav_bar():
    selected_page = option_menu(
        menu_title=None,
        options=["File Selection","Path"],
        orientation="horizontal"
    )

    return selected_page

