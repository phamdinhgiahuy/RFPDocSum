import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import openpyxl
from copy import copy
import io
from io import BytesIO
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.styles import PatternFill, Font, Border, Fill
from difflib import SequenceMatcher
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from fuzzywuzzy import fuzz
import os


def clean_dataframe(df):
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df.replace(["", "None", "N/A", "NaN", " "], pd.NA, inplace=True)
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~(df.iloc[1:].isna().all() | (df.iloc[1:] == "").all())]
    return df


# Function to clean and normalize column names in all the files
def clean_column_names(df):
    df.columns = df.columns.str.strip().str.lower()
    return df


# Function to handle unnamed columns in a dataframe
def handle_unnamed_columns(df):
    df.columns = ["" if "unnamed" in str(col).lower() else col for col in df.columns]
    return df


# Function to find common columns by comparing values
def find_common_columns_by_values_worksheet(
    template_sheet, supplier_sheet, threshold=90
):

    common_columns = []
    supplier_value_columns = []
    mis_mat_rows = []

    for col in template_sheet.iter_cols():
        # Get the column letter of the current column
        col_letter = get_column_letter(col[0].column)
        # print(f"Column: {col_letter}")
        # Iterate over all rows in the current column
        row_values_template = [cell.value for cell in col if cell.value is not None]

        row_values_suppliers = []
        for row_sup in supplier_sheet.iter_rows(
            min_col=col[0].column, max_col=col[0].column
        ):
            # Get non-None values in the current column of the supplier sheet
            if row_sup[0].value is not None:
                row_values_suppliers.append(row_sup[0].value)
        # if both lists are empty, skip the column
        if not row_values_template and not row_values_suppliers:
            continue
        # fuzzy matching between the string joint from the template and the supplier list if the similarity exceeds the threshold of 70%
        temp_row_str = " ".join([str(val) for val in row_values_template])
        suppliers_row_str = " ".join([str(val) for val in row_values_suppliers])
        similarity = fuzz.ratio(temp_row_str, suppliers_row_str)
        if similarity > threshold:
            # print(f"Supplier Values: {row_values_suppliers}")
            # print(f"Template Values: {row_values_template}")
            # print(f"Matched with similarity: {similarity}")
            common_columns.append(col_letter)
            if similarity < 100:
                # highlight the row in the supplier sheet that does not match the template
                for row_sup in supplier_sheet.iter_rows(
                    min_col=col[0].column, max_col=col[0].column
                ):
                    if (
                        row_sup[0].value is not None
                        and row_sup[0].value not in row_values_template
                    ):
                        # print(f"Detected mismatch in row: {row_sup[0].coordinate}")
                        # print(f"Value: {row_sup[0].value}")
                        # add coordinates of the mismatched row
                        mis_mat_rows.append(row_sup[0].coordinate)
        else:
            # this column is not common, can be the column that contains the supplier values
            # print(f"Column {col_letter} is not common")
            # print(f"Supplier Values: {row_values_suppliers}")
            # print(f"Template Values: {row_values_template}")
            if row_values_suppliers:
                supplier_value_columns.append(col_letter)

    return common_columns, mis_mat_rows, supplier_value_columns


# Function to overwrite row values in target columns based on row-by-row similarity and highlight cells below threshold
def overwrite_columns_based_on_similarity(
    template_df, df, common_columns, threshold=0.9
):
    # Reset index to ensure that the rows are aligned for row-by-row comparison
    template_df = template_df.reset_index(drop=True)
    df = df.reset_index(drop=True)

    yellow_fill = PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )

    for col in common_columns:
        template_values = template_df[col].astype(str).str.lower().fillna("")
        df_values = df[col].astype(str).str.lower().fillna("")
        for idx, (template_val, df_val) in enumerate(zip(template_values, df_values)):
            similarity = SequenceMatcher(None, template_val, df_val).ratio()
            if similarity >= threshold:
                df.at[idx, col] = template_df.at[idx, col]
            else:
                df.at[idx, col] = df.at[idx, col]
                df.style.applymap(
                    lambda x: (
                        yellow_fill if x == df_val and similarity < threshold else None
                    ),
                    subset=[col],
                )

    return df


def combine_excel_files_with_similarity_worksheets(
    workbook,
    template_sheets,
    supplier_sheets_dict,
    threshold=80,
):
    for idx, template_sheet in enumerate(template_sheets):
        # create a new sheet in the workbook for each template sheet
        target_sheet = workbook.create_sheet(template_sheet.title)
        common_columns = []
        mis_mat_rows_dict = {}
        supplier_value_columns_dict = {}
        for supplier in supplier_sheets_dict:
            supplier_sheet = supplier_sheets_dict[supplier][idx]
            # find the common columns between the template sheet and the supplier sheet
            common_columns, mis_mat_rows, supplier_value_columns = (
                find_common_columns_by_values_worksheet(
                    template_sheet, supplier_sheet, threshold
                )
            )
            print(f"Common Columns: {common_columns}")
            print(f"Mismatched Rows: {mis_mat_rows}")
            print(f"Supplier Value Columns: {supplier_value_columns}")
            # store the results in dictionaries
            mis_mat_rows_dict[supplier] = mis_mat_rows
            supplier_value_columns_dict[supplier] = supplier_value_columns

        # write the common columns from the template sheet to the target sheet
        for col in common_columns:
            # column letter to index
            col_idx = openpyxl.utils.column_index_from_string(col)
            # copy the template value columns to the col idx
            copy_column(template_sheet, target_sheet, col_idx)
        # insert the supplier value columns to the target sheet
        for supplier, columns in supplier_value_columns_dict.items():
            for col in columns:
                col_idx = openpyxl.utils.column_index_from_string(col)
                # insert a new column at position of the supplier value column
                target_sheet.insert_cols(col_idx)
                copy_column(supplier_sheets_dict[supplier][idx], target_sheet, col_idx)
                # change the header of the column to include the supplier name
                header = target_sheet.cell(row=1, column=col_idx).value
                target_sheet.cell(
                    row=1,
                    column=col_idx,
                    value=f"{header} ({supplier})" if header else supplier,
                )
    # remove the default sheet
    if "Sheet" in workbook.sheetnames:
        workbook.remove(workbook["Sheet"])
    return workbook


# Function to summarize text using Sumy
def summarize_column_simple(text, sentence_count=3):
    try:
        parser = PlaintextParser.from_string(text, Tokenizer("english"))
        summarizer = LsaSummarizer()
        summarizer.stop_words = "english"
        summary = summarizer(parser.document, sentence_count)
        return " ".join(str(sentence) for sentence in summary)
    except Exception as e:
        return f"Error summarizing text: {e}"


def adjust_formula_references(formula, sheet_name_mapping):
    for old_name, new_name in sheet_name_mapping.items():
        if f"'{old_name}'!" in formula:  # Check for explicit sheet references
            formula = formula.replace(f"'{old_name}'!", f"'{new_name}'!")
    return formula


# function to save the consolidated file and let it be downloadable by the user
def save_consolidated_file(consolidated):
    file_stream = io.BytesIO()
    consolidated.save(file_stream)
    file_stream.seek(0)
    return file_stream


def format_workbook(workbook, image_path, image_width=100):
    sheet_names = workbook.sheetnames
    worksheets = [workbook[sheet_name] for sheet_name in sheet_names]
    for worksheet in worksheets:
        # Format header row (blue with white text)
        # header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        # font_style_header = Font(b=True, color="FFFFFF")
        # for col_idx in range(1, worksheet.max_column + 1):
        #     cell = worksheet.cell(row=4, column=col_idx)
        #     cell.fill = header_fill
        #     cell.font = font_style_header
        #     cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # # Format second header row (gray with black bold text)
        # second_header_fill = PatternFill(start_color="EEECE1", end_color="EEECE1", fill_type="solid")
        # font_style_second_header = Font(b=True)
        # for col_idx in range(1, worksheet.max_column + 1):
        #     cell = worksheet.cell(row=5, column=col_idx)
        #     cell.fill = second_header_fill
        #     cell.font = font_style_second_header
        #     cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        pass
        # Insert logo
        if image_path:
            from openpyxl.drawing.image import Image

            logo = Image(image_path)
            logo.width = image_width
            worksheet.add_image(logo, "A1")

    return workbook


def copy_column(source_sheet, target_sheet, column_index):
    """Copy values, styles, and attributes from one column to another."""
    col_letter = get_column_letter(column_index)

    for row in range(1, source_sheet.max_row + 1):
        source_cell = source_sheet.cell(row=row, column=column_index)
        target_cell = target_sheet.cell(row=row, column=column_index)

        target_cell.value = source_cell.value
        target_cell.data_type = source_cell.data_type

        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.hyperlink:
            target_cell.hyperlink = source_cell.hyperlink
        # Copy comment if present
        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)

    # Copy column dimensions (width, hidden status, etc.)
    if col_letter in source_sheet.column_dimensions:
        source_dim = source_sheet.column_dimensions[col_letter]
        target_dim = target_sheet.column_dimensions[col_letter]

        target_dim.width = source_dim.width
        target_dim.hidden = source_dim.hidden
        target_dim.outlineLevel = source_dim.outlineLevel
        target_dim.min = source_dim.min
        target_dim.max = source_dim.max


# function to copy the sheet from source to target
def copy_sheet(source_sheet, target_sheet):
    copy_cells(source_sheet, target_sheet)  # copy all the cell values and styles
    copy_sheet_attributes(source_sheet, target_sheet)


# function to copy the sheet attributes from source to target
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


# function to copy the cells values and attributes from source to target
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
    dfs_dict = {}
    worksheets_dict = {}
    for supplier in supplier_info:
        if supplier.get(doc_type):
            dfs_dict[supplier["name"]] = []
            worksheets_dict[supplier["name"]] = []
        else:
            st.warning(
                f"No {doc_type} file found for supplier {supplier['name']}.", icon="⚠️"
            )
            continue
        # read only
        sup_excel = load_workbook(supplier[doc_type])
        sup_sheets = [
            sup_excel[sheet]
            for sheet in sup_excel.sheetnames
            if sup_excel[sheet].sheet_state == "visible"
        ]
        for sheet_idx in sheet_indexes:
            sheet = sup_sheets[sheet_idx]
            title = f"{supplier['name']}_{sheet.title}"[:30]
            worksheets_dict[supplier["name"]].append(sheet)
            dfs_dict[supplier["name"]].append(
                pd.read_excel(
                    supplier[doc_type], sheet_name=sheet.title, engine="openpyxl"
                )
            )
    return dfs_dict, worksheets_dict


doc_type1, doc_type2 = st.session_state.doc_types
st.write("# RFP Suppliers' Response Consolidation")
if (
    "suppliers" not in st.session_state
    or st.session_state.template_files[doc_type2] is None
):
    st.error("No supplier data found. Please complete the setup first.")
    st.stop()

### Session State Variables Retrieval

event_name = st.session_state.event_name
event_option = st.session_state.event_option
# logo path
st.session_state.logo_path = r"kellanova_logo.png"


###Global Variables
file_colors = ["E4DFEC", "D4D5F8", "DFD2FA", "D9E5F3", "E0E0EC", "E7DAF2", "E3E9E9"]

### Pricing Sheets Consolidation

st.header("Pricing Sheets Consolidation")
template_pri = st.session_state.template_files[doc_type1]
wb_template_pri = load_workbook(template_pri)
all_sheets_pri = wb_template_pri.sheetnames

pricing_sheets_list = st.multiselect(
    "Select Pricing sheet(s) to consolidate", all_sheets_pri, all_sheets_pri
)

chosen_sheets_pri_idx = [all_sheets_pri.index(sheet) for sheet in pricing_sheets_list]
pri_comb_mode = st.radio(
    "Choose consolidation method for Pricing sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

if st.button("Consolidate Pricing Sheets"):
    consolidated_pri = openpyxl.Workbook()
    consolidated_pri.remove(consolidated_pri.active)
    dfs_pri_dict, sheets_pri_dict = get_files(
        st.session_state.suppliers, chosen_sheets_pri_idx, doc_type1
    )
    total_sheets = len(sheets_pri_dict)
    template_sheets = [wb_template_pri[sheet] for sheet in pricing_sheets_list]

    if pri_comb_mode == "Side by Side":
        progress = st.progress(0)
        consolidated_pri = combine_excel_files_with_similarity_worksheets(
            consolidated_pri, template_sheets, sheets_pri_dict
        )

        # fill the progress bar
        progress.progress(1)
        progress.empty()
    elif pri_comb_mode == "Separate Sheets":
        progress = st.progress(0)
        # total sheets to process
        total_sheets = sum(
            [len(sheets_pri_dict[supplier]) for supplier in sheets_pri_dict]
        )
        for supplier in sheets_pri_dict:
            for sheet in sheets_pri_dict[supplier]:
                title = f"{supplier}_{sheet.title}"[:30]
                target_sheet = consolidated_pri.create_sheet(title)
                copy_sheet(sheet, target_sheet)

                # update progress bar
                progress.progress(
                    (sheets_pri_dict[supplier].index(sheet) + 1) / total_sheets
                )

        progress.empty()
    consolidated_pri = format_workbook(consolidated_pri, st.session_state.logo_path)
    file_stream_p = save_consolidated_file(consolidated_pri)
    # save to session state
    st.session_state.consolidated_p = file_stream_p


if "consolidated_p" in st.session_state and st.session_state.consolidated_p is not None:
    # option to download the consolidated file if it exists
    download_pricing = True
    st.download_button(
        "Download Consolidated Pricing File",
        data=st.session_state.consolidated_p,
        file_name=f"{event_name}_{doc_type1}_consolidated.xlsx",
    )
else:
    st.session_state.consolidated_p = None
    download_pricing = False


### Questionnaire Sheets Consolidation

st.header("Questionnaire Consolidation")
template_ques = st.session_state.template_files[doc_type2]
wb_template_ques = load_workbook(template_ques)
all_sheets_ques = wb_template_ques.sheetnames

questionnaire_sheets_list = st.multiselect(
    "Select Questionnaire sheet(s) to consolidate",
    all_sheets_ques,
    all_sheets_ques,
)

chosen_sheets_ques_idx = [
    all_sheets_ques.index(sheet) for sheet in questionnaire_sheets_list
]

ques_comb_mode = st.radio(
    "Choose consolidation method for Questionnaire sheets:",
    ("Side by Side", "Separate Sheets"),
    index=0,
)

if st.button("Consolidate Questionnaire Sheets"):
    consolidated_ques = openpyxl.Workbook()
    consolidated_ques.remove(consolidated_ques.active)
    dfs_ques_dict, sheets_ques_dict = get_files(
        st.session_state.suppliers, chosen_sheets_ques_idx, doc_type2
    )
    total_sheets = len(sheets_ques_dict)
    template_sheets = [wb_template_ques[sheet] for sheet in questionnaire_sheets_list]

    if ques_comb_mode == "Side by Side":
        progress = st.progress(0)
        print(sheets_ques_dict)
        consolidated_ques = combine_excel_files_with_similarity_worksheets(
            consolidated_ques, template_sheets, sheets_ques_dict
        )
        # fill the progress bar
        progress.progress(1)
        progress.empty()
    elif ques_comb_mode == "Separate Sheets":
        progress = st.progress(0)
        # total sheets to process
        total_sheets = sum(
            [len(sheets_ques_dict[supplier]) for supplier in sheets_ques_dict]
        )

        for supplier in sheets_ques_dict:
            for sheet in sheets_ques_dict[supplier]:
                title = f"{supplier}_{sheet.title}"[:30]
                target_sheet = consolidated_ques.create_sheet(title)
                copy_sheet(sheet, target_sheet)

                # update progress bar
                progress.progress(
                    (sheets_ques_dict[supplier].index(sheet) + 1) / total_sheets
                )

        progress.empty()
    consolidated_ques = format_workbook(consolidated_ques, st.session_state.logo_path)
    file_stream_q = save_consolidated_file(consolidated_ques)
    # save to session state
    st.session_state.consolidated_q = file_stream_q

if "consolidated_q" in st.session_state and st.session_state.consolidated_q is not None:
    download_questionnaire = True
    st.download_button(
        "Download Consolidated Questionnaire File",
        data=st.session_state.consolidated_q,
        file_name=f"{event_name}_{doc_type2}_consolidated.xlsx",
    )
else:
    st.session_state.consolidated_q = None
    download_questionnaire = False
