import streamlit as st
import os
import glob

st.write("# RFP Event Configuration")
event_name = st.text_input("RFP Event Name")

rfp_folder_path = st.text_input(
    "RFP Folder Path", value=f"../Project Files/Datasets/{event_name}"
)

col1, col2 = st.columns(2)
with col1:
    doc_type1 = st.text_input("💰Pricing documents keyword", value="Pricing")
with col2:
    doc_type2 = st.text_input(
        "❓Questionnaire documents keyword", value="Questionnaire"
    )


if "files" not in st.session_state:
    st.session_state.files = {doc_type1: [], doc_type2: []}

if st.button("Search for Files"):
    if rfp_folder_path and os.path.isdir(rfp_folder_path):

        st.session_state.files = {doc_type1: [], doc_type2: []}

        for file in glob.glob(
            os.path.join(rfp_folder_path, "**", "*.xlsx"), recursive=True
        ):
            if doc_type1.lower() in file.lower():
                st.session_state.files[doc_type1].append(file)
            elif doc_type2.lower() in file.lower():
                st.session_state.files[doc_type2].append(file)

        if (
            len(st.session_state.files[doc_type1]) > 0
            or len(st.session_state.files[doc_type2]) > 0
        ):
            st.success("Files found!", icon="✅")
        else:
            st.error("No files found! Please verify the folder path.", icon="❌")

        with st.expander("Search Results"):
            if st.session_state.files[doc_type1]:
                st.write(f"Found the following '{doc_type1}' files:")
                for file in st.session_state.files[doc_type1]:
                    st.info(file, icon="📄")
            else:
                st.warning(f"No '{doc_type1}' files found.", icon="⚠️")

            if st.session_state.files[doc_type2]:
                st.write(f"Found the following '{doc_type2}' files:")
                for file in st.session_state.files[doc_type2]:
                    st.info(file, icon="📄")
            else:
                st.warning(f"No '{doc_type2}' files found.", icon="⚠️")
    else:
        st.error("Please provide a valid and existing folder path.", icon="❌")


st.write("### Suppliers Information")
num_suppliers = st.number_input(
    "Number of Suppliers", min_value=1, step=1, key="num_suppliers"
)


if (
    "suppliers" not in st.session_state
    or len(st.session_state.suppliers) != num_suppliers
):
    st.session_state.suppliers = [{} for _ in range(num_suppliers)]


for i in range(num_suppliers):
    with st.expander(f"Supplier {i+1} Information"):
        supplier_name = st.text_input(f"Supplier {i+1} Name", key=f"name_{i}")

        if not supplier_name:
            st.warning(f"Please enter a name for Supplier {i+1}.")
            continue

        st.session_state.suppliers[i]["name"] = supplier_name

        for doc_type in [doc_type1, doc_type2]:
            matching_files = [
                file
                for file in st.session_state.files[doc_type]
                if supplier_name.lower() in file.lower()
            ]

            if matching_files:
                if len(matching_files) == 1:
                    st.success(
                        f"Found file for {doc_type}: {matching_files[0]}", icon="✅"
                    )
                    st.session_state.suppliers[i][doc_type] = matching_files[0]
                else:
                    st.warning(
                        f"Multiple files found for {doc_type}. Please select the correct file.",
                        icon="⚠️",
                    )
                    selected_file = st.selectbox(
                        f"Select {doc_type} file",
                        matching_files,
                        key=f"{doc_type}_select_{i}",
                    )
                    st.session_state.suppliers[i][doc_type] = selected_file
            else:
                st.error(f"No file found for {doc_type}. Please upload one.", icon="📎")
                uploaded_file = st.file_uploader(
                    f"Upload {doc_type} file", key=f"{doc_type}_upload_{i}"
                )

                if uploaded_file is not None:
                    st.session_state.suppliers[i][doc_type] = uploaded_file


if st.button("Submit my config"):
    st.write("Event Name:", event_name)
    st.write("RFP Folder Path:", rfp_folder_path)
    st.write("Number of Suppliers:", num_suppliers)

if "suppliers" in st.session_state and len(st.session_state.suppliers) > 0:
    with st.expander("Submitted Suppliers Information"):
        st.write(st.session_state.suppliers)
