import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

# Function to create a PowerPoint slide from the provided data
def create_ppt_slide(data):
    prs = Presentation()
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    title = slide.shapes.title
    title.text = "Crisis Scenario Overview"

    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    for index, row in data.iterrows():
        p = text_frame.add_paragraph()
        p.text = f"{row['Field']}: {row['Details']}"
        p.font.size = Pt(14)
    
    return prs

# Streamlit app to take CSV input and generate a PowerPoint slide
st.title("PowerPoint Slide Creator")

data_file = st.file_uploader("Upload CSV File", type=["csv"])

if data_file is not None:
    # Read the CSV file
    try:
        # Attempt to read the CSV file with more robust handling
        df = pd.read_csv(data_file, encoding='utf-8', skip_blank_lines=True, engine='python')
        st.write("### Uploaded Data:")
        st.write(df)

        # Check if the required columns are present
        if 'Field' in df.columns and 'Details' in df.columns:
            # Button to generate PowerPoint
            if st.button("Create PowerPoint Slide"):
                ppt = create_ppt_slide(df[['Field', 'Details']])
                
                # Save to BytesIO object
                ppt_io = BytesIO()
                ppt.save(ppt_io)
                ppt_io.seek(0)
                
                # Provide download link
                st.download_button(label="Download PowerPoint Slide", 
                                   data=ppt_io, 
                                   file_name="scenario_overview.pptx", 
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        else:
            st.error("CSV must contain 'Field' and 'Details' columns.")
    except pd.errors.ParserError:
        st.error("Error reading file: The CSV file appears to be badly formatted. Please check for missing quotes or incorrect delimiters.")
    except Exception as e:
        st.error(f"Error reading file: {e}")
