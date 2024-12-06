import streamlit as st


# Set initial configuration if not already set
def initialize_session_state():
    """
    Initialize the session state with the required variables if they do not exist.

    The variables are:
    - event_name: The name of the RFP event, which will appear in the filename for the consolidated document.
    - event_option: The document configuration for this event, which can be either "In a Single File" or "In Separate Files".
    - suppliers: A list of supplier names.
    - template_files: A dictionary with keys "Pricing" and "Questionnaire", and values that are the uploaded template files.
    - doc_types: A list of document types, which is initially set to ["Pricing", "Questionnaire"].
    """
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


initialize_session_state()
# Configuration form
st.image(r"assets/kellanova_logo.png", width=200)
st.write("# RFP Event Configuration")

# Event name
event_name = st.text_input(
    "RFP Event Name",
    value=st.session_state.event_name,
    placeholder="This name will appear in the filename for the consolidated document.",
)

st.session_state.event_name = event_name

# Event option: "In a Single File" or "In Separate Files"
event_option = st.radio(
    "Select the document configuration for this event",
    ("In a Single File", "In Separate Files"),
)


# # Update session state when the event_option changes
if event_option != st.session_state.event_option:
    st.session_state.event_option = event_option


# Upload template files
st.write("### 🗝️ Template Files")
if event_option == "In a Single File":
    st.markdown("#### :blue[**Combined** template file]")
    combined_template = st.file_uploader(
        "Please upload Combined Template File", type=["xlsx", "xls"]
    )
    if combined_template:
        st.session_state.template_files["Pricing"] = combined_template
        st.session_state.template_files["Questionnaire"] = combined_template
else:
    for doc_type in ["Pricing", "Questionnaire"]:
        st.markdown(
            f"#### :{'green' if doc_type == 'Pricing' else 'orange'}[**{doc_type}** template file]"
        )
        uploaded_file = st.file_uploader(
            f"Please upload {doc_type} Template File", type=["xlsx", "xls"]
        )
        if uploaded_file:
            st.session_state.template_files[doc_type] = uploaded_file
st.write("### 🗃️ Suppliers Response Files")
# Number of suppliers
num_suppliers = st.number_input(
    "Number of Suppliers in this event", min_value=1, step=1
)
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
            f"Supplier {i + 1} Name",
            value=st.session_state.suppliers[i]["name"],
            placeholder="This name will be used to refer to the supplier in the consolidated document.",
        )
        st.session_state.suppliers[i]["name"] = supplier_name

        # Conditionally show file upload based on the event_option
        if event_option == "In a Single File":
            st.markdown("#### :blue[**Combined** response file]")
            combined_file = st.file_uploader(
                f"Please upload Combined File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"combined_{i}",
            )
            if combined_file:
                st.session_state.suppliers[i]["Pricing"] = combined_file
                st.session_state.suppliers[i]["Questionnaire"] = combined_file
        else:
            st.markdown("#### :green[**Pricing** file]")
            pricing_file = st.file_uploader(
                f"Please upload Pricing File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"pricing_{i}",
            )
            if pricing_file:
                st.session_state.suppliers[i]["Pricing"] = pricing_file

            st.markdown("#### :orange[**Questionnaire** file]")
            questionnaire_file = st.file_uploader(
                f"Please upload Questionnaire File for Supplier {i + 1}",
                type=["xlsx", "xls"],
                key=f"questionnaire_{i}",
            )
            if questionnaire_file:
                st.session_state.suppliers[i]["Questionnaire"] = questionnaire_file

if st.button("Submit Configuration"):
    st.session_state.submitted = True
    st.success(
        "Configuration updated successfully! Please proceed to the consolidation page."
    )
