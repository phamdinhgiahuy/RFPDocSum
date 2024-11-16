import streamlit as st
import os
import glob

st.write("# RFP Event Configuration")
event_name = st.text_input("RFP Event Name")

col1, col2 = st.columns(2)
with col1:
    doc_type1 = st.text_input("💰Pricing documents keyword", value="Pricing")
with col2:
    doc_type2 = st.text_input(
        "❓Questionnaire documents keyword", value="Questionnaire"
    )

if "template_files" not in st.session_state:
    st.session_state.template_files = {doc_type1: None, doc_type2: None}

st.write("### Template Files")
for doc_type in [doc_type1, doc_type2]:
    uploaded_file = st.file_uploader(
        f"Please upload {doc_type} template file", key=f"{doc_type}_template"
    )
    if uploaded_file is not None:
        st.session_state.template_files[doc_type] = uploaded_file


if "suppliers" not in st.session_state:
    st.session_state.suppliers = []
st.write("### Suppliers Information")
num_suppliers = st.number_input(
    "Number of Suppliers", min_value=1, step=1, key="num_suppliers"
)


if len(st.session_state.suppliers) != num_suppliers:
    # append new suppliers
    for i in range(num_suppliers - len(st.session_state.suppliers)):
        st.session_state.suppliers.append(
            {"name": None, doc_type1: None, doc_type2: None}
        )


for i in range(num_suppliers):
    with st.expander(f"Supplier {i+1} Information"):
        supplier_name = st.text_input(f"Supplier {i+1} Name", key=f"name_{i}")

        if supplier_name:
            st.session_state.suppliers[i]["name"] = supplier_name

        for doc_type in [doc_type1, doc_type2]:
            uploaded_file = st.file_uploader(
                f"Please upload {doc_type} file for Supplier {i+1}",
                key=f"{doc_type}_upload_{i}",
            )
            if uploaded_file is not None:
                st.session_state.suppliers[i][doc_type] = uploaded_file


with st.expander("Debugging Info"):
    st.write("Event Name:", event_name)
    st.write("Number of Suppliers:", num_suppliers)
    st.write("Template Files:", st.session_state.template_files)
    if "suppliers" in st.session_state and len(st.session_state.suppliers) > 0:
        st.write(st.session_state.suppliers)
