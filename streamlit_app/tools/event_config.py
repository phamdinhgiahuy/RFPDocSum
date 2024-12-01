import streamlit as st

# Set initial configuration if not already set
if "event_name" not in st.session_state:
    st.session_state.event_name = ""
if "event_option" not in st.session_state:
    st.session_state.event_option = "In a Single File"
if "suppliers" not in st.session_state:
    st.session_state.suppliers = []
if "template_files" not in st.session_state:
    st.session_state.template_files = {"Pricing": None, "Questionnaire": None}
if "doc_types" not in st.session_state or len(st.session_state.doc_types) == 0:
    st.session_state.doc_types = ["Pricing", "Questionnaire"]

# Configuration form
st.write("# RFP Event Configuration")

# Event name
event_name = st.text_input("RFP Event Name", value=st.session_state.event_name)
st.session_state.event_name = event_name

# Event option: "In a Single File" or "In Separate Files"
event_option = st.radio(
    "Select the document configuration for this event",
    ("In a Single File", "In Separate Files"),
    index=0 if st.session_state.event_option == "In a Single File" else 1,
)


# Update session state when the event_option changes
if event_option != st.session_state.event_option:
    st.session_state.event_option = event_option

# Upload template files
if event_option == "In a Single File":
    combined_template = st.file_uploader(
        "Upload Combined Template File", type=["xlsx", "xls"]
    )
    if combined_template:
        st.session_state.template_files["Pricing"] = combined_template
        st.session_state.template_files["Questionnaire"] = combined_template
else:
    for doc_type in ["Pricing", "Questionnaire"]:
        uploaded_file = st.file_uploader(
            f"Upload {doc_type} Template File", type=["xlsx", "xls"]
        )
        if uploaded_file:
            st.session_state.template_files[doc_type] = uploaded_file

# Number of suppliers
num_suppliers = st.number_input("Number of Suppliers", min_value=1, step=1)
st.session_state.num_suppliers = num_suppliers

# Supplier details (ensure there is a supplier object for each supplier)
if len(st.session_state.suppliers) < num_suppliers:
    st.session_state.suppliers.extend(
        [{"name": "", "Pricing": None, "Questionnaire": None}]
        * (num_suppliers - len(st.session_state.suppliers))
    )
elif len(st.session_state.suppliers) > num_suppliers:
    st.session_state.suppliers = st.session_state.suppliers[:num_suppliers]

# Add Supplier Info Form
for i in range(num_suppliers):
    with st.expander(f"Supplier {i + 1} Information"):
        supplier_name = st.text_input(
            f"Supplier {i + 1} Name", value=st.session_state.suppliers[i]["name"]
        )
        st.session_state.suppliers[i]["name"] = supplier_name

        # Conditionally show file upload based on the event_option
        if event_option == "In a Single File":
            combined_file = st.file_uploader(
                f"Upload Combined File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"combined_{i}",
            )
            if combined_file:
                st.session_state.suppliers[i]["Pricing"] = combined_file
                st.session_state.suppliers[i]["Questionnaire"] = combined_file
        else:
            pricing_file = st.file_uploader(
                f"Upload Pricing File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"pricing_{i}",
            )
            if pricing_file:
                st.session_state.suppliers[i]["Pricing"] = pricing_file
            questionnaire_file = st.file_uploader(
                f"Upload Questionnaire File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"questionnaire_{i}",
            )
            if questionnaire_file:
                st.session_state.suppliers[i]["Questionnaire"] = questionnaire_file

# Submit Button to update session state and allow routing to the next page
# submit_disabled = getattr(st.session_state, "submitted", False)
if st.button("Submit Configuration"):
    st.session_state.submitted = True
    st.success(
        "Configuration updated successfully! Please proceed to the consolidation page."
    )

# Debugging Info - Show only after submission
if getattr(st.session_state, "submitted", False):
    with st.expander("Debugging Info"):
        st.write("### Event Details")
        st.write(f"- **Event Name**: {event_name}")
        st.write(f"- **Number of Suppliers**: {st.session_state.num_suppliers}")
        st.write("- **Template Files**:")
        for doc_type, file in st.session_state.template_files.items():
            st.write(f"  - {doc_type}: {file.name if file else 'Not Uploaded'}")
        st.write("- **Suppliers Info**:")
        for i, supplier in enumerate(st.session_state.suppliers):
            st.write(f"  - Supplier {i+1}:")
            st.write(f"    - Name: {supplier.get('name', 'Unnamed')}")
            for doc_type in ["Pricing", "Questionnaire"]:
                file = supplier.get(doc_type)
                st.write(f"    - {doc_type}: {file.name if file else 'Not Uploaded'}")
        st.write(f"- **Event Option**: {event_option}")
