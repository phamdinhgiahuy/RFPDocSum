import streamlit as st

st.write("# RFP Event Configuration")

# Initialize or retrieve the event name
if "event_name" not in st.session_state:
    st.session_state.event_name = ""
event_name = st.text_input("RFP Event Name", value=st.session_state.event_name)
st.session_state.event_name = event_name

# Initialize document types
if "doc_types" not in st.session_state:
    st.session_state.doc_types = ["Pricing", "Questionnaire"]
doc_type1, doc_type2 = st.session_state.doc_types

# Event Option
if "event_option" not in st.session_state:
    st.session_state.event_option = "Combined Pricing and Questionnaire Document"
event_option = st.radio(
    "Choose the response document configuration for this event",
    (
        "Combined Pricing and Questionnaire Document",
        "Separate Pricing and Questionnaire Documents",
    ),
    index=(
        "Combined Pricing and Questionnaire Document",
        "Separate Pricing and Questionnaire Documents",
    ).index(st.session_state.event_option),
)
st.session_state.event_option = event_option

# Template Files
if "template_files" not in st.session_state:
    st.session_state.template_files = {doc_type1: None, doc_type2: None}

st.write("### Event Template Files")
if event_option == "Combined Pricing and Questionnaire Document":
    combined_template = st.file_uploader(
        "Please upload combined template file", key="combined_template"
    )
    if combined_template is not None:
        st.session_state.template_files[doc_type1] = combined_template
        st.session_state.template_files[doc_type2] = combined_template
else:
    for doc_type in st.session_state.doc_types:
        uploaded_file = st.file_uploader(
            f"Please upload {doc_type} template file", key=f"{doc_type}_template"
        )
        if uploaded_file is not None:
            st.session_state.template_files[doc_type] = uploaded_file

# Suppliers Information
if "suppliers" not in st.session_state:
    st.session_state.suppliers = []

num_suppliers = st.number_input(
    "Number of Suppliers",
    min_value=1,
    step=1,
    value=max(1, len(st.session_state.suppliers)),
)


# Adjust supplier list size in state
if len(st.session_state.suppliers) != num_suppliers:
    if len(st.session_state.suppliers) < num_suppliers:
        st.session_state.suppliers.extend(
            [{"name": None, doc_type1: None, doc_type2: None}]
            * (num_suppliers - len(st.session_state.suppliers))
        )
    else:
        st.session_state.suppliers = st.session_state.suppliers[:num_suppliers]

# Supplier details
if event_option == "Combined Pricing and Questionnaire Document":
    st.write(
        "Please upload the combined Pricing and Questionnaire documents for each supplier"
    )
    for i in range(num_suppliers):
        with st.expander(f"Supplier {i+1} Information"):
            supplier_name = st.text_input(
                f"Supplier {i+1} Name",
                value=st.session_state.suppliers[i]["name"],
                key=f"name_{i}",
            )
            st.session_state.suppliers[i]["name"] = supplier_name

            uploaded_file = st.file_uploader(
                f"Please upload combined template file for Supplier {i+1}",
                key=f"combined_template_upload_{i}",
            )
            if uploaded_file is not None:
                st.session_state.suppliers[i][doc_type1] = uploaded_file
                st.session_state.suppliers[i][doc_type2] = uploaded_file
else:
    for i in range(num_suppliers):
        with st.expander(f"Supplier {i+1} Information"):
            supplier_name = st.text_input(
                f"Supplier {i+1} Name",
                value=st.session_state.suppliers[i]["name"],
                key=f"name_{i}",
            )
            st.session_state.suppliers[i]["name"] = supplier_name

            for doc_type in st.session_state.doc_types:
                uploaded_file = st.file_uploader(
                    f"Please upload {doc_type} template file for Supplier {i+1}",
                    key=f"{doc_type}_upload_{i}",
                )
                if uploaded_file is not None:
                    st.session_state.suppliers[i][doc_type] = uploaded_file

# Debugging Info
with st.expander("Debugging Info"):
    st.write("Event Name:", event_name)
    st.write("Number of Suppliers:", num_suppliers)
    st.write("Template Files:", st.session_state.template_files)
    st.write("Suppliers:", st.session_state.suppliers)
    st.write("Event Option:", event_option)
