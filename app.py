import streamlit as st
from presentation_generator import generate_presentation_from_prompt, slugify_filename, pptx_to_pdf
import os
from dotenv import load_dotenv
load_dotenv()

# Inject custom CSS for a modern look with a gradient background that works in Streamlit
st.markdown('''
    <style>
    .stApp {
        background: linear-gradient(135deg, #f7f9fa 0%, #e3e6f3 100%);
        min-height: 100vh;
        padding-bottom: 2rem;
    }
    .main {
        background: transparent !important;
    }
    .stTextArea textarea {font-size: 1.1rem;}
    .stButton button {background-color: #0057b8; color: white; font-weight: bold; border-radius: 6px; padding: 0.5em 1.5em;}
    .stButton button:hover {background-color: #003f7d;}
    .stSelectbox div {font-size: 1.05rem;}
    .stSlider > div {font-size: 1.05rem;}
    .stDownloadButton button {background-color: #00b894; color: white; font-weight: bold; border-radius: 6px;}
    .stDownloadButton button:hover {background-color: #008c6e;}
    .stSuccess {color: #00b894;}
    </style>
''', unsafe_allow_html=True)

st.set_page_config(page_title="AI Presentation Generator", layout="centered")
st.title("AI Presentation Generator")
st.write("Enter your topic or prompt below. The app will generate a PowerPoint presentation with AI-generated content and images.")

prompt = st.text_area("Presentation Topic or Prompt", "The Future of Artificial Intelligence")

# Slide count selection
slide_range = st.slider("Number of Slides (range)", min_value=3, max_value=15, value=(5, 10))
min_slides, max_slides = slide_range

# Font selection
font_options = [
    "Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma", "Georgia", "Comic Sans MS"
]
font_name = st.selectbox("Font for Slides", font_options, index=0)

if st.button("Generate Presentation"):
    with st.spinner("Generating presentation. This may take up to a minute..."):
        filename = slugify_filename(prompt)
        generate_presentation_from_prompt(
            prompt,
            min_slides=min_slides,
            max_slides=max_slides,
            font_name=font_name
        )
        pptx_ready = os.path.exists(filename)
        pdf_filename = filename.replace('.pptx', '.pdf')
        pdf_ready = False
        if pptx_ready:
            # Try to convert to PDF
            pdf_ready = pptx_to_pdf(os.path.abspath(filename), os.path.abspath(pdf_filename)) and os.path.exists(pdf_filename)
        if pptx_ready:
            st.success(f"Presentation generated: {filename}")
            col1, col2 = st.columns(2)
            with col1:
                with open(filename, "rb") as f:
                    st.download_button(
                        label="Download PPTX",
                        data=f,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            with col2:
                if pdf_ready:
                    with open(pdf_filename, "rb") as f:
                        st.download_button(
                            label="Download PDF",
                            data=f,
                            file_name=pdf_filename,
                            mime="application/pdf"
                        )
                else:
                    st.warning("PDF conversion failed or PowerPoint is not installed.")
        else:
            st.error("Failed to generate the presentation file.") 