import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import openpyxl
from copy import copy
import os
import subprocess
import platform
import io

# def open_file_location(file_path):
#     if platform.system() == "Windows":
#         os.startfile(file_path)
#     elif platform.system() == "Darwin":
#         subprocess.Popen(["open", file_path])
#     else:
#         subprocess.Popen(["xdg-open", file_path])


def save_consolidated_file(consolidated):
    file_stream = io.BytesIO()
    consolidated.save(file_stream)
    file_stream.seek(0)
    return file_stream


def copy_sheet(source_sheet, target_sheet):
    copy_cells(source_sheet, target_sheet)  # copy all the cell values and styles
    copy_sheet_attributes(source_sheet, target_sheet)


def copy_sheet_attributes(source_sheet, target_sheet):
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)

    for rn in range(len(source_sheet.row_dimensions)):
        target_sheet.row_dimensions[rn] = copy(source_sheet.row_dimensions[rn])

    for key, value in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[key].min = copy(value.min)
        target_sheet.column_dimensions[key].max = copy(value.max)
        target_sheet.column_dimensions[key].width = copy(value.width)
        target_sheet.column_dimensions[key].hidden = copy(value.hidden)


def copy_cells(source_sheet, target_sheet):
    for (row, col), source_cell in source_sheet._cells.items():
        target_cell = target_sheet.cell(column=col, row=row)
        target_cell._value = source_cell._value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)


def get_files(supplier_info, sheet_indexes):
    dfs_price = {}
    worksheets_price = {}

    for supplier in supplier_info:
        if supplier.get("Pricing"):
            dfs_price[supplier["name"]] = []
            worksheets_price[supplier["name"]] = []
        else:
            st.error("No Pricing file found for supplier.", icon="❌")
            continue

        sup_excel = load_workbook(supplier["Pricing"])
        sheets_name = [sup_excel.sheetnames[i] for i in sheet_indexes]

        for sheet in sheets_name:
            temp_df = pd.read_excel(supplier["Pricing"], sheet_name=sheet, header=7)
            dfs_price[supplier["name"]].append(temp_df)
            worksheets_price[supplier["name"]].append(sup_excel[sheet])

    return dfs_price, worksheets_price


st.write("# RFP Suppliers' Response Consolidation")

if (
    "suppliers" not in st.session_state
    or st.session_state.suppliers[0].get("Pricing") is None
):
    st.error("No supplier data found. Please complete the setup first.")
    st.stop()

st.header("Pricing Sheets Consolidation")

first_supplier_file = st.session_state.suppliers[0]["Pricing"]
workbook = load_workbook(first_supplier_file)
available_sheets = workbook.sheetnames

st.write("Available sheets in the first supplier's Pricing file:")
pricing_sheets_list = st.multiselect(
    "Select Pricing sheet(s) to consolidate", available_sheets, available_sheets
)

chosen_sheet_indices = [available_sheets.index(sheet) for sheet in pricing_sheets_list]

pricing_option = st.radio(
    "Choose consolidation method for Pricing sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

if st.button("Consolidate Sheets"):
    consolidated = openpyxl.Workbook()
    if pricing_option == "Side by Side":
        st.write("Consolidating Pricing sheets side by side...")
        # Side by Side consolidation logic here
    else:
        st.write("Consolidating Pricing documents into separate sheets...")

        dfs_price, worksheets_price = get_files(
            st.session_state.suppliers, chosen_sheet_indices
        )

        for supplier in worksheets_price:
            for sheet in worksheets_price[supplier]:
                title = f"{supplier}_{sheet.title}"
                # trim the title if it's too long
                if len(title) > 30:
                    title = title[:30]
                target_sheet = consolidated.create_sheet(title)
                copy_sheet(sheet, target_sheet)
        if "Sheet" in consolidated.sheetnames:
            consolidated.remove(consolidated["Sheet"])
        file_stream = save_consolidated_file(consolidated)
        # export_path = "pricing_consolidated.xlsx"

        # consolidated.save(export_path)
        st.success(
            "Consolidated Pricing and Implementation sheets saved as 'pricing_consolidated.xlsx",
            icon="✅",
        )
        # if st.button("Show File in Folder"):
        #     open_file_location(os.path.abspath(export_path))
        st.download_button(
            label="Download Consolidated File",
            data=file_stream,
            file_name="pricing_consolidated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


st.header("Questionaire Sheets Consolidation")

questionnaire_option = st.radio(
    "Choose consolidation method for Questionnaire sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

first_supplier_file_q = st.session_state.suppliers[0]["Questionnaire"]
workbook_q = load_workbook(first_supplier_file_q)
available_sheets_q = workbook_q.sheetnames

st.write("Available sheets in the first supplier's Questionnaire file:")
questionnaire_sheets_list = st.multiselect(
    "Select Questionnaire sheet(s) to consolidate",
    available_sheets_q,
    available_sheets_q,
)

chosen_sheet_indices_q = [
    available_sheets_q.index(sheet) for sheet in questionnaire_sheets_list
]

if questionnaire_option == "Side by Side":
    st.write("Consolidating Questionnaire sheets side by side...")
    # Side by Side consolidation logic here
else:
    st.write("Consolidating Questionnaire sheets into separate sheets...")
    dfs_questionnaire, worksheets_questionnaire = get_files(
        st.session_state.suppliers, chosen_sheet_indices_q
    )

    for supplier in worksheets_questionnaire:
        for sheet in worksheets_questionnaire[supplier]:
            title = f"{supplier}_{sheet.title}"
            target_sheet = consolidated.create_sheet(title)
            copy_sheet(sheet, target_sheet)
    if "Sheet" in consolidated.sheetnames:
        consolidated.remove(consolidated["Sheet"])

    consolidated.save("questionnaire_consolidated.xlsx")
    st.success(
        "Consolidated Questionnaire sheets saved as 'consolidated.xlsx'", icon="✅"
    )
