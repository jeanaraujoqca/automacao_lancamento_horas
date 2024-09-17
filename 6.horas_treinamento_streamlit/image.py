import streamlit as st 
import base64 

# Função para carregar imagem local e codificar como base64
def get_base64_of_image(image_file):
    with open(image_file, "rb") as image:
        return base64.b64encode(image.read()).decode()

# Função para adicionar CSS com a imagem de fundo
def add_css(background_image_base64):
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url(data:image/jpeg;base64,{background_image_base64});
            background-size: cover;
        }}
        #root > div:nth-child(1) > div.withScreencast > div > div > header {{
            background-image: url(data:image/jpeg;base64,{background_image_base64});
            background-size: cover;
            color: White;
        }}
        .stApp h1 {{
            color: #B22222;
            font-size: 50px;
        }}
        .stApp p {{
            color: #A52A2A;
            font-size: 18px;
        }}
        #root > div:nth-child(1) > div.withScreencast > div > div > div > section > div.block-container.st-emotion-cache-13ln4jf.ea3mdgi5 > div > div > div > div:nth-child(5) > div > section > button {{
            color: #A52A2A;
            font-size: 18px;
        }}
        #root > div:nth-child(1) > div.withScreencast > div > div > div > section > div.block-container.st-emotion-cache-13ln4jf.ea3mdgi5 > div > div > div > div:nth-child(5) > div > section {{
            background-color: #1E2345;
            color: White;
        }}   
        #root > div:nth-child(1) > div.withScreencast > div > div > div > section > div.block-container.st-emotion-cache-13ln4jf.ea3mdgi5 > div > div > div > div:nth-child(5) > div > section > div > div > small {{
            color: White;
        }}
        
        </style>
        """,
        unsafe_allow_html=True
    )

