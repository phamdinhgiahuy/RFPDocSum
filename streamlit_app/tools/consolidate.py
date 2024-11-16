import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import openpyxl
from copy import copy
import os
import subprocess
import platform
import io


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


def get_files(supplier_info, sheet_indexes, doc_type):
    dfs = {}
    worksheets = {}
    for supplier in supplier_info:
        if supplier.get(doc_type):
            dfs[supplier["name"]] = []
            worksheets[supplier["name"]] = []
        else:
            st.warning(
                f"No {doc_type} file found for supplier {supplier['name']}.", icon="⚠️"
            )
            continue
        sup_excel = load_workbook(supplier[doc_type])
        sheets_name = [sup_excel.sheetnames[i] for i in sheet_indexes]
        for sheet in sheets_name:
            temp_df = pd.read_excel(supplier[doc_type], sheet_name=sheet, header=7)
            dfs[supplier["name"]].append(temp_df)
            worksheets[supplier["name"]].append(sup_excel[sheet])
    return dfs, worksheets


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
pricing_sheets_list = st.multiselect(
    "Select Pricing sheet(s) to consolidate", available_sheets, available_sheets
)
chosen_sheet_indices = [available_sheets.index(sheet) for sheet in pricing_sheets_list]
pricing_option = st.radio(
    "Choose consolidation method for Pricing sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

if st.button("Consolidate Pricing Sheets"):
    consolidated = openpyxl.Workbook()
    consolidated.remove(consolidated["Sheet"])
    dfs_price, worksheets_price = get_files(
        st.session_state.suppliers, chosen_sheet_indices, "Pricing"
    )
    for supplier in worksheets_price:
        for sheet in worksheets_price[supplier]:
            title = f"{supplier}_{sheet.title}"[:30]
            target_sheet = consolidated.create_sheet(title)
            copy_sheet(sheet, target_sheet)
    file_stream = save_consolidated_file(consolidated)
    st.download_button(
        "Download Consolidated Pricing File",
        data=file_stream,
        file_name="pricing_consolidated.xlsx",
    )

st.header("Questionnaire Sheets Consolidation")
first_supplier_file_q = st.session_state.suppliers[0]["Questionnaire"]
workbook_q = load_workbook(first_supplier_file_q)
available_sheets_q = workbook_q.sheetnames
questionnaire_sheets_list = st.multiselect(
    "Select Questionnaire sheet(s) to consolidate",
    available_sheets_q,
    available_sheets_q,
)
chosen_sheet_indices_q = [
    available_sheets_q.index(sheet) for sheet in questionnaire_sheets_list
]
questionnaire_option = st.radio(
    "Choose consolidation method for Questionnaire sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

if st.button("Consolidate Questionnaire Sheets"):
    consolidated_q = openpyxl.Workbook()
    consolidated_q.remove(consolidated_q["Sheet"])
    dfs_questionnaire, worksheets_questionnaire = get_files(
        st.session_state.suppliers, chosen_sheet_indices_q, "Questionnaire"
    )
    for supplier in worksheets_questionnaire:
        for sheet in worksheets_questionnaire[supplier]:
            title = f"{supplier}_{sheet.title}"[:30]
            target_sheet = consolidated_q.create_sheet(title)
            copy_sheet(sheet, target_sheet)
    file_stream_q = save_consolidated_file(consolidated_q)
    st.download_button(
        "Download Consolidated Questionnaire File",
        data=file_stream_q,
        file_name="questionnaire_consolidated.xlsx",
    )
