import streamlit as st
import pandas as pd
from docx import Document
from jinja2 import Template
import os
from io import BytesIO

# Utility to render Word doc with Jinja2
def render_docx_template(template_path, context):
    doc = Document(template_path)
    for para in doc.paragraphs:
        for run in para.runs:
            run.text = Template(run.text).render(**context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.set_page_config(page_title="Loan Report Generator", layout="centered")
st.title("📄 Loan Report Generator")

# Ensure the templates directory exists
if not os.path.exists("templates"):
    os.makedirs("templates")
    st.warning("Please add your Word templates to the 'templates' directory.")

uploaded_excel = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_excel:
    df = pd.read_excel(uploaded_excel)

    st.write("### Preview Customer Data")
    st.dataframe(df)

    selected_index = st.number_input("Select customer row (0 = first row)", min_value=0, max_value=len(df)-1, value=0)
    customer_data = df.iloc[selected_index].to_dict()

    st.write("Selected Customer:")
    st.json(customer_data)

    if st.button("Generate Reports"):
        template_dir = "templates"
        output_buffers = []

        for template_name in os.listdir(template_dir):
            if template_name.endswith(".docx"):
                template_path = os.path.join(template_dir, template_name)
                output_buffer = render_docx_template(template_path, customer_data)
                output_buffers.append((template_name, output_buffer))

        st.success(f"✅ Generated {len(output_buffers)} reports.")

        for name, buffer in output_buffers:
            st.download_button(
                label=f"Download {name}",
                data=buffer,
                file_name=f"{customer_data['Name']}_{name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    
