import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import openpyxl
from copy import copy
import io
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from difflib import SequenceMatcher
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from fuzzywuzzy import fuzz
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


# Function to find common columns by comparing values
def find_matching_cols(template_sheet, supplier_sheet, threshold=90):

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


def separate_sheet_combine(
    workbook, template_sheets, supplier_sheets_dict, threshold=80
):
    st.toast(f"Combining files in progress...", icon="⏳")
    for idx, template_sheet in enumerate(template_sheets):
        # copy the template sheet to the workbook
        target_sheet = workbook.create_sheet(template_sheet.title)
        copy_sheet(template_sheet, target_sheet)
        for supplier in supplier_sheets_dict:
            print(f"Processing supplier: {supplier}")
            # add a new sheet for each supplier
            supplier_sheet = supplier_sheets_dict[supplier][idx]
            sheet_title = f"{supplier}_{template_sheet.title}"[:30]
            target_sheet = workbook.create_sheet(sheet_title)
            copy_sheet(supplier_sheet, target_sheet)
            _, mis_mat_rows, _ = find_matching_cols(
                template_sheet, supplier_sheet, threshold
            )
            # highligh the mismatched rows
            if mis_mat_rows:
                for mis_mat_row in mis_mat_rows:
                    col_mis, row_mis = openpyxl.utils.cell.coordinate_from_string(
                        mis_mat_row
                    )
                    row_mis = int(row_mis)  # Convert to integer
                    col_mis = column_index_from_string(col_mis)  # Convert to integer
                    if col_mis and row_mis:
                        target_sheet.cell(row=row_mis, column=col_mis).fill = (
                            PatternFill(
                                start_color="FAA0A0",
                                end_color="FAA0A0",
                                fill_type="solid",
                            )
                        )
                        print(
                            f"Filled cell: {col_mis} and {row_mis} in sheet: {sheet_title}"
                        )
                    print(
                        f"Detected mismatched row: {mis_mat_row}. Highlighting cell: {col_mis} and {row_mis}"
                    )
    # remove the default sheet
    if "Sheet" in workbook.sheetnames:
        workbook.remove(workbook["Sheet"])
    st.toast("Seperate-Sheet File combined successfully! Ready to download", icon="🎉")
    return workbook


def side_by_side_combine(
    workbook,
    template_sheets,
    supplier_sheets_dict,
    threshold=80,
):
    st.toast(f"Combining files in progress...", icon="⏳")
    for idx, template_sheet in enumerate(template_sheets):

        # create a new sheet in the workbook for each template sheet
        target_sheet = workbook.create_sheet(template_sheet.title)
        # copy sheet attributes from the template sheet to the target sheet

        common_columns = []
        uncommon_columns = []
        mis_mat_rows_dict = {}
        supplier_value_columns_dict = {}
        for supplier in supplier_sheets_dict:
            supplier_sheet = supplier_sheets_dict[supplier][idx]
            com_columns, mis_mat_rows, supplier_value_columns = find_matching_cols(
                template_sheet, supplier_sheet, threshold
            )
            # if values in com_columns not already in common_columns, add them
            common_columns.extend(
                [col for col in com_columns if col not in common_columns]
            )
            # if values in supplier_value_columns not already in uncommon_columns, add them
            uncommon_columns.extend(
                [col for col in supplier_value_columns if col not in uncommon_columns]
            )
            print(f"Common Columns: {com_columns}")
            print(f"Mismatched Rows: {mis_mat_rows}")
            print(f"Supplier Value Columns: {supplier_value_columns}")
            mis_mat_rows_dict[supplier] = mis_mat_rows
            supplier_value_columns_dict[supplier] = supplier_value_columns
        # write the common columns from the template sheet to the target sheet
        for col in common_columns:
            # column letter to index
            col_idx = openpyxl.utils.column_index_from_string(col)
            # copy the template value columns to the col idx
            copy_column(template_sheet, target_sheet, col_idx, col_idx)

        # insert the supplier value columns to the target sheet
        print(
            f"Mismatched Rows Dict for the template sheet {template_sheet.title}: {mis_mat_rows_dict}"
        )
        print("###########")
        print(
            f"Supplier Value Columns Dict for the template sheet {template_sheet.title}: {supplier_value_columns_dict}"
        )
        num_inserts = len(supplier_sheets_dict) - 1
        insert_offset = 0
        uncommon_columns = sorted(
            uncommon_columns, key=lambda x: openpyxl.utils.column_index_from_string(x)
        )
        print(f"Uncommon Columns: {uncommon_columns}")
        for col in uncommon_columns:
            # column letter to index
            col_idx = openpyxl.utils.column_index_from_string(col)
            # insert a new column at the position of the uncommon column
            for i in range(num_inserts):
                insert_idx = col_idx + insert_offset
                target_sheet.insert_cols(insert_idx + 1)
                insert_offset += 1
            # copy the template value columns to the col idx

        for i, (supplier, value_columns) in enumerate(
            supplier_value_columns_dict.items()
        ):
            # insert_offset = 0
            mis_mat_rows = mis_mat_rows_dict[supplier]
            for j, col in enumerate(value_columns):
                col_idx = openpyxl.utils.column_index_from_string(col)
                tar_idx = col_idx + j * num_inserts + i

                copy_column(
                    supplier_sheets_dict[supplier][idx], target_sheet, col_idx, tar_idx
                )
                print(
                    f"copy column {col_idx} for supplier: {supplier} to column: {tar_idx} in sheet: {target_sheet.title}"
                )
                # change the header of the column to include the supplier name
                target_sheet.cell(row=1, column=tar_idx, value=f"{supplier}")
                # Format the header cell
                header_cell = target_sheet.cell(row=1, column=tar_idx)
                header_cell.font = Font(
                    name="Arial", size=15, bold=True, color="FFFFFF"
                )
                header_cell.fill = PatternFill(
                    start_color="0080ff", end_color="0080ff", fill_type="solid"
                )
                header_cell.border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                )
                # center the header text
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

                # highlight the mismatched rows
                if mis_mat_rows:
                    for mis_mat_row in mis_mat_rows:
                        # check the coordinates of the mismatched row
                        col_mis, row_mis = openpyxl.utils.cell.coordinate_from_string(
                            mis_mat_row
                        )
                        st.toast(
                            f"Detected template discrepancy in cell: {mis_mat_row} sheet: {template_sheet.title} for supplier: {supplier}",
                            icon="⚠️",
                        )
                        # make the cell background red
                        target_sheet.cell(row=row_mis, column=tar_idx).fill = (
                            PatternFill(
                                start_color="FAA0A0",
                                end_color="FAA0A0",
                                fill_type="solid",
                            )
                        )

        # toast the completion of the sheet
        # st.toast(f"{template_sheet.title} consolidated!", icon="✔️")
    # remove the default sheet
    if "Sheet" in workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    st.toast("Side-By-Side File combined successfully! Ready to download", icon="🎉")

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


# function to save the consolidated file and let it be downloadable by the user
def save_consolidated_file(consolidated):
    file_stream = io.BytesIO()
    consolidated.save(file_stream)
    file_stream.seek(0)
    return file_stream


def append_logo(workbook, image_path, image_scale=0.8):
    sheet_names = workbook.sheetnames
    worksheets = [workbook[sheet_name] for sheet_name in sheet_names]
    for worksheet in worksheets:
        # add logo to the first cell of the worksheet
        if image_path:
            from openpyxl.drawing.image import Image

            logo = Image(image_path)
            logo.width = int(logo.width * image_scale)
            logo.height = int(logo.height * image_scale)
            worksheet.add_image(logo, "A1")

    return workbook


def copy_column(source_sheet, target_sheet, souce_col_idx, target_col_idx):

    source_col_letter = get_column_letter(souce_col_idx)
    target_col_letter = get_column_letter(target_col_idx)

    for row in range(1, source_sheet.max_row + 1):
        source_cell = source_sheet.cell(row=row, column=souce_col_idx)
        target_cell = target_sheet.cell(row=row, column=target_col_idx)

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

    source_dim = source_sheet.column_dimensions[source_col_letter]
    target_dim = target_sheet.column_dimensions[target_col_letter]

    target_dim.width = copy(source_dim.width)


# function to copy the sheet from source to target
def copy_sheet(source_sheet, target_sheet):
    # get all columns in the source sheet
    for col in source_sheet.iter_cols(max_col=50):
        col_idx = col[0].column
        copy_column(source_sheet, target_sheet, col_idx, col_idx)
    copy_sheet_attributes(source_sheet, target_sheet)


# function to copy the sheet attributes from source to target
def copy_sheet_attributes(source_sheet, target_sheet):
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)

    for rn, source_row in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[rn] = copy(source_row)

    for key, value in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[key].width = copy(value.width)


def get_files(supplier_info, sheet_indexes, doc_type):
    dfs_dict = {}
    worksheets_dict = {}
    with st.spinner("Reading files..."):
        for supplier in supplier_info:
            if supplier.get(doc_type):
                dfs_dict[supplier["name"]] = []
                worksheets_dict[supplier["name"]] = []
            else:
                st.warning(
                    f"No {doc_type} file found for supplier {supplier['name']}.",
                    icon="⚠️",
                )
                continue
            # read only
            sup_excel = load_workbook(
                supplier[doc_type], rich_text=True, data_only=True
            )
            sup_sheets = [
                sup_excel[sheet]
                for sheet in sup_excel.sheetnames
                if sup_excel[sheet].sheet_state == "visible"
            ]
            for sheet_idx in sheet_indexes:
                sheet = sup_sheets[sheet_idx]
                # title = f"{supplier['name']}_{sheet.title}"[:30]
                # sheet.title = title
                worksheets_dict[supplier["name"]].append(sheet)
                # dfs_dict[supplier["name"]].append(
                #     pd.read_excel(
                #         supplier[doc_type], sheet_name=sheet.title, engine="openpyxl"
                #     )
                # )
    st.toast("Supplier Files read successfully! 📚", icon="✅")

    return dfs_dict, worksheets_dict


doc_type1, doc_type2 = st.session_state.doc_types
st.write("# RFP Files Consolidator")
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


### Pricing Sheets Consolidation

st.write("### Pricing information")
template_pri = st.session_state.template_files[doc_type1]
wb_template_pri = load_workbook(template_pri, rich_text=True, data_only=True)
all_sheets_pri = wb_template_pri.sheetnames

pricing_sheets_list = st.multiselect(
    "Please select Pricing sheet(s) to consolidate", all_sheets_pri, all_sheets_pri
)

chosen_sheets_pri_idx = [all_sheets_pri.index(sheet) for sheet in pricing_sheets_list]

if "pri_comb_mode" not in st.session_state:
    st.session_state.pri_comb_mode = "Side by Side"

st.radio(
    "Please select consolidation method:",
    ("Side by Side", "Separate Sheets"),
    index=0,
    key="pri_comb_mode",
)

st.write(f"### Current Mode: **{st.session_state.pri_comb_mode}**")

if st.button("Consolidate", key="consolidate_pri"):
    consolidated_pri = openpyxl.Workbook()
    consolidated_pri.remove(consolidated_pri.active)
    dfs_pri_dict, sheets_pri_dict = get_files(
        st.session_state.suppliers, chosen_sheets_pri_idx, doc_type2
    )
    template_sheets_pri = [wb_template_pri[sheet] for sheet in pricing_sheets_list]
    with st.spinner("Processing... Please wait."):
        if st.session_state.pri_comb_mode == "Side by Side":
            consolidated_pri = side_by_side_combine(
                consolidated_pri, template_sheets_pri, sheets_pri_dict
            )
        elif st.session_state.pri_comb_mode == "Separate Sheets":
            consolidated_pri = separate_sheet_combine(
                consolidated_pri, template_sheets_pri, sheets_pri_dict
            )
        consolidated_pri = append_logo(consolidated_pri, st.session_state.logo_path)
        file_stream_p = save_consolidated_file(consolidated_pri)
    if file_stream_p is None:
        st.error("Failed to save the consolidated file. Please try again.")
        st.stop()
    # save to session state
    st.session_state.consolidated_p = file_stream_p
    st.success("Pricing sheets consolidated successfully!", icon="✅")


if st.session_state.get("consolidated_p"):
    st.download_button(
        f"💾 Download {event_name}_{doc_type1}_consolidated.xlsx",
        data=st.session_state.consolidated_p,  # Changed from consolidated_q
        file_name=f"{event_name}_{doc_type1}_consolidated.xlsx",
    )
else:
    st.session_state.consolidated_p = None


### Questionnaire Sheets Consolidation

st.write("### Questionnaire information")
template_ques = st.session_state.template_files[doc_type2]
wb_template_ques = load_workbook(template_ques, rich_text=True, data_only=True)
all_sheets_ques = wb_template_ques.sheetnames

questionnaire_sheets_list = st.multiselect(
    "Please select Questionnaire sheet(s) to consolidate",
    all_sheets_ques,
    all_sheets_ques,
)

chosen_sheets_ques_idx = [
    all_sheets_ques.index(sheet) for sheet in questionnaire_sheets_list
]

if "ques_comb_mode" not in st.session_state:
    st.session_state.ques_comb_mode = "Side by Side"

st.radio(
    "Please select consolidation method:",
    ("Side by Side", "Separate Sheets"),
    index=0,
    key="ques_comb_mode",
)

st.write(f"### Current Mode: **{st.session_state.ques_comb_mode}**")

if st.button("Consolidate", key="consolidate_ques"):
    consolidated_ques = openpyxl.Workbook()
    consolidated_ques.remove(consolidated_ques.active)
    dfs_ques_dict, sheets_ques_dict = get_files(
        st.session_state.suppliers, chosen_sheets_ques_idx, doc_type2
    )
    template_sheets_ques = [
        wb_template_ques[sheet] for sheet in questionnaire_sheets_list
    ]
    with st.spinner("Processing... Please wait."):
        if st.session_state.ques_comb_mode == "Side by Side":
            consolidated_ques = side_by_side_combine(
                consolidated_ques, template_sheets_ques, sheets_ques_dict
            )
        elif st.session_state.ques_comb_mode == "Separate Sheets":
            consolidated_ques = separate_sheet_combine(
                consolidated_ques, template_sheets_ques, sheets_ques_dict
            )
        consolidated_ques = append_logo(consolidated_ques, st.session_state.logo_path)
        file_stream_q = save_consolidated_file(consolidated_ques)

    # save to session state
    st.session_state.consolidated_q = file_stream_q
    st.success("Questionnaire sheets consolidated successfully!", icon="✅")

download_questionnaire = False

if st.session_state.get("consolidated_q"):
    st.download_button(
        f"💾 Download {event_name}_{doc_type2}_consolidated.xlsx",
        data=st.session_state.consolidated_q,
        file_name=f"{event_name}_{doc_type2}_consolidated.xlsx",
    )
    download_questionnaire = True
else:
    st.session_state.consolidated_q = None
