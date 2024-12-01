import streamlit as st
import re
import pandas as pd
from collections import defaultdict
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
from openpyxl.utils.dataframe import dataframe_to_rows
from fuzzywuzzy import fuzz
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from itertools import cycle
from collections import deque
from openpyxl.cell import Cell

# from openpyxl.cell.rich_text import RichText


def fill_color_switch():
    colors = [
        "E4DFEC",
        "D4D5F8",
        "DFD2FA",
        "D9E5F3",
        "E0E0EC",
        "E7DAF2",
        "E3E9E9",
        "CC79A7",
        "009E73",
        "0072B2",
    ]
    return cycle(colors)


def generate_merged_dict(sheet):
    merged_dict = {}

    for merged_range in sheet.merged_cells.ranges:
        # Only handle vertical merges (i.e., merged cells in the same column)
        if merged_range.size["columns"] == 1:
            start_cell = merged_range.start_cell
            column_idx = start_cell.column  # Get the column index
            merged_range_str = (
                merged_range.coord
            )  # Get the range as string (e.g., 'C14:C17')

            # If the column is not already in merged_dict, initialize an empty list
            if column_idx not in merged_dict:
                merged_dict[column_idx] = []

            # Add the merged range to the list of merged ranges for this column
            merged_dict[column_idx].append(merged_range_str)

    return merged_dict


def merge_columns_in_target_sheet(
    target_sheet, merged_dict, source_col_idx, target_col_idx
):
    if source_col_idx not in merged_dict:
        print(f"No merged ranges found for column {source_col_idx} in merged_dict.")
        return
    for merged_range in merged_dict[source_col_idx]:
        # Extract the start and end rows from the merged range (e.g., 'C14:C17' -> 14, 17)
        start_cell, end_cell = merged_range.split(":")
        start_row = int(start_cell[1:])
        end_row = int(end_cell[1:])
        start_column = column_index_from_string(start_cell[0])
        target_sheet.merge_cells(
            start_row=start_row,
            start_column=target_col_idx,  # Use target_col_idx for the target column
            end_row=end_row,
            end_column=target_col_idx,
        )


# Function to find common columns by comparing values
def find_matching_cols(template_sheet, supplier_sheet, threshold=90):
    common_columns = []
    supplier_value_columns = []
    mis_mat_rows = []

    for col in template_sheet.iter_cols():
        # Get the column letter of the current column
        col_letter = get_column_letter(col[0].column)
        row_values_template = [cell.value for cell in col if cell.value is not None]

        row_values_suppliers = []
        for row_sup in supplier_sheet.iter_rows(
            min_col=col[0].column, max_col=col[0].column
        ):
            # Get non-None values in the current column of the supplier sheet
            if row_sup[0].value is not None:
                row_values_suppliers.append(row_sup[0].value)

        # If both lists are empty, skip the column
        if not row_values_template and not row_values_suppliers:
            continue

        # Fuzzy matching between the string joint from the template and the supplier list
        temp_row_str = " ".join([str(val) for val in row_values_template])
        suppliers_row_str = " ".join([str(val) for val in row_values_suppliers])
        similarity = fuzz.ratio(temp_row_str, suppliers_row_str)

        if similarity > threshold:
            common_columns.append(col_letter)
            if similarity < 100:
                # Highlight the row in the supplier sheet that does not match the template
                for row_sup in supplier_sheet.iter_rows(
                    min_col=col[0].column, max_col=col[0].column
                ):
                    if row_sup[0].value is not None:
                        # Compare the pure text without formatting
                        cell_value = str(row_sup[0].value)
                        if cell_value not in row_values_template:
                            # Detected mismatch in row, add coordinates of the mismatched row
                            mis_mat_rows.append(row_sup[0].coordinate)
        else:
            # This column is not common, can be the column that contains the supplier values
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


def create_insertion_queue(common_columns, supplier_value_columns_dict):
    # Initialize a list to hold queue items
    queue_items = []

    # Add common columns from the template to the queue
    for col in common_columns:
        queue_items.append({"column_letter": col, "source": "template"})

    # Add supplier columns to the queue
    for supplier, supplier_columns in supplier_value_columns_dict.items():
        for col in supplier_columns:
            queue_items.append({"column_letter": col, "source": supplier})

    # Sort the queue by column letter and include a secondary sort by source type
    # Priority: template columns come before supplier columns with the same letter
    queue_items.sort(key=lambda x: (x["column_letter"], x["source"] != "template"))

    # Convert the sorted list into a deque (queue structure)
    insertion_queue = deque(queue_items)

    return insertion_queue


def side_by_side_combine(
    workbook,
    template_sheets,
    supplier_sheets_dict,
    threshold=80,
    summary_option=False,
):
    st.toast(f"Combining files in progress...", icon="⏳")

    # Iterate over each template sheet
    for idx, template_sheet in enumerate(template_sheets):
        print(f"Processing template sheet: {template_sheet.title}")

        # Create a new sheet in the workbook for each template sheet
        target_sheet = workbook.create_sheet(template_sheet.title)

        # Initialize variables to store column data and mismatched rows
        common_columns = []
        uncommon_columns = []
        mis_mat_rows_dict = {}
        supplier_value_columns_dict = {}
        supplier_colors = {}
        color_cycle = fill_color_switch()

        # Iterate over each supplier and process their sheet
        for supplier in supplier_sheets_dict:
            if supplier not in supplier_colors:
                supplier_colors[supplier] = next(color_cycle)

            supplier_sheet = supplier_sheets_dict[supplier][idx]
            com_columns, mis_mat_rows, supplier_value_columns = find_matching_cols(
                template_sheet, supplier_sheet, threshold
            )

            # Debug: Print matched and mismatched columns
            print(f"Supplier: {supplier}")
            print(f"  Common Columns: {com_columns}")
            print(f"  Mismatched Rows: {mis_mat_rows}")
            print(f"  Supplier Value Columns: {supplier_value_columns}")

            # Add matching columns to the final list (ensure no duplicates)
            common_columns.extend(
                [col for col in com_columns if col not in common_columns]
            )
            uncommon_columns.extend(
                [col for col in supplier_value_columns if col not in uncommon_columns]
            )

            # Store mismatched rows and value columns for the supplier
            mis_mat_rows_dict[supplier] = mis_mat_rows
            supplier_value_columns_dict[supplier] = supplier_value_columns

        # Copy common columns from template to target sheet
        print(f"Template Common Columns: {common_columns}")
        print(f"Supplier Value Columns: {supplier_value_columns_dict}")
        queue = create_insertion_queue(common_columns, supplier_value_columns_dict)
        if not queue:
            print("No columns to process for this template sheet.")
            continue

        for i, item in enumerate(queue):
            col_letter = item["column_letter"]
            source = item["source"]
            header_fill_color = None
            mis_mat_rows = None
            # Determine the source sheet
            if source == "template":
                source_sheet = template_sheet
            else:
                source_sheet = supplier_sheets_dict[source][idx]
                header_fill_color = supplier_colors[source]
                mis_mat_rows = mis_mat_rows_dict[source]

            # Get column indices
            col_idx_source = column_index_from_string(col_letter)
            col_idx_target = i + 1  # Insert in the order of the queue

            # Copy column
            print(
                f"Copying column {col_letter} from {source} to target at index {col_idx_target}"
            )
            end_row_write = copy_column(
                source_sheet, target_sheet, col_idx_source, col_idx_target
            )
            # format the header cell for the supplier if there are mismatched rows
            if header_fill_color:
                # header cell is the first bold cell in the column
                header_cell = target_sheet.cell(row=1, column=col_idx_target)
                header_cell.fill = PatternFill(
                    fill_type="solid", start_color=header_fill_color
                )
                # change the value of the header cell to include the supplier name
                if header_cell.value:
                    header_cell.value = f"{source}  {header_cell.value}"
                else:
                    header_cell.value = f"{source}"

                header_cell.font = Font(name="Arial", size=15, bold=True)
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                # apply color formatting to the data rows
                # empty_rows = 0
                for row in target_sheet.iter_rows(
                    min_row=2,
                    max_row=end_row_write,
                    min_col=col_idx_target,
                    max_col=col_idx_target,
                ):
                    for cell in row:
                        # if cell.value is None or cell.value == "":
                        #     empty_rows += 1
                        # if empty_rows > 20:
                        #     print("Too many empty rows")
                        #     break
                        cell.fill = PatternFill(
                            fill_type="solid", start_color=header_fill_color
                        )
                # apply color formatting to the mismatched rows
                if mis_mat_rows:
                    for row in mis_mat_rows:
                        col_mis, row_mis = openpyxl.utils.cell.coordinate_from_string(
                            row
                        )
                        row_mis = int(row_mis)  # Convert to integer
                        col_mis = column_index_from_string(
                            col_mis
                        )  # Convert to integer
                        if col_mis and row_mis:
                            target_sheet.cell(row=row_mis, column=col_mis).fill = (
                                PatternFill(
                                    start_color="FAA0A0",
                                    end_color="FAA0A0",
                                    fill_type="solid",
                                )
                            )
                # Add summary if requested
                if summary_option:
                    source_text = " ".join(
                        str(cell.value)
                        for row in target_sheet.iter_rows(
                            min_row=2,
                            max_row=end_row_write,
                            min_col=col_idx_target,
                            max_col=col_idx_target,
                        )
                        for cell in row
                        if cell.value
                        and not str(cell.value)
                        .replace(".", "", 1)
                        .isdigit()  # Exclude numbers
                    )
                    if len(source_text.split()) > 5:
                        summary = summarize_column_simple(source_text)
                        target_sheet.cell(
                            row=end_row_write + 1,
                            column=col_idx_target,
                            value="Summary:",
                        )
                        summary_cell = target_sheet.cell(
                            row=end_row_write + 2, column=col_idx_target
                        )
                        summary_cell.value = summary
                        summary_cell.font = Font(name="Arial", size=12, bold=False)
                        summary_cell.fill = PatternFill(
                            fill_type="solid", start_color="BFFFFF"
                        )
                        summary_cell.alignment = Alignment(
                            horizontal="center", vertical="center", wrap_text=True
                        )

        # Copy the template sheet attributes to the target sheet
        # copy_sheet_attributes(template_sheet, target_sheet)
        st.toast(f"{template_sheet.title} consolidated!", icon="✔️")
    # Remove the default sheet if it exists
    if "Sheet" in workbook.sheetnames:
        workbook.remove(workbook["Sheet"])

    # Final success toast
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


def copy_column(source_sheet, target_sheet, source_col_idx, target_col_idx):
    print(
        f"Copying column {source_col_idx} from {source_sheet.title} to {target_col_idx} in {target_sheet.title}"
    )
    source_col_letter = get_column_letter(source_col_idx)
    target_col_letter = get_column_letter(target_col_idx)

    # Initialize target_row at the first row of the target sheet

    hidden_count = 0
    end_row_idx = 300  # default value
    # Track merged cells in the source column
    # quick check the first 40 rows to see if there are any values, if all is empty, break the loop
    max_rows_to_check = min(source_sheet.max_row, 40)
    check_empty = 0
    for row in range(1, max_rows_to_check + 1):
        source_cell = source_sheet.cell(row=row, column=source_col_idx)
        if source_cell.value not in (
            None,
            "",
        ):  # If a non-empty cell is found, exit early
            print(f"Non-empty cell detected at row {row}.")
            break
        check_empty += 1

    # If all rows are empty, exit the function
    if check_empty == max_rows_to_check:
        print("All checked rows are empty. Exiting the function.")
        return end_row_idx
    target_row = 1
    merged_dict = generate_merged_dict(source_sheet)
    # Iterate over each row in the source sheet
    empty_rows_cont = 0
    for row in range(1, min(source_sheet.max_row + 1, 300)):
        source_cell = source_sheet.cell(row=row, column=source_col_idx)
        target_cell = target_sheet.cell(row=target_row, column=target_col_idx)
        if source_cell.value is None or source_cell.value == "":
            empty_rows_cont += 1
            if empty_rows_cont > 20:
                # print("More than 20 continuous empty rows detected. breaking the loop.")
                end_row_idx = row - empty_rows_cont + 1
                break
        else:
            empty_rows_cont = 0
        # Skip hidden rows in the source sheet
        if getattr(source_sheet.row_dimensions[row], "hidden", False):
            hidden_count += 1
            if hidden_count > 30:
                # print("More than 30 hidden rows detected. breaking the loop.")
                end_row_idx = row - hidden_count + 1
                break
        else:
            hidden_count = 0
        # if found more than 30 continuous empty rows, break the loop

        target_cell.value = source_cell.value
        target_cell.data_type = source_cell.data_type

        try:
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.border = copy(source_cell.border)
                target_cell.fill = copy(source_cell.fill)
                target_cell.number_format = copy(source_cell.number_format)
                target_cell.protection = copy(source_cell.protection)
                target_cell.alignment = copy(source_cell.alignment)
        except Exception as e:
            print(f"Error copying styles for cell {source_cell.coordinate}: {e}")

        if source_cell.hyperlink:
            target_cell.hyperlink = source_cell.hyperlink

        if source_cell.comment:
            target_cell.comment = copy(source_cell.comment)

        # Increment the target row index for the next copy
        target_row += 1

    if end_row_idx == 300:
        end_row_idx = target_row
    # Perform merging of cells after copying data
    merge_columns_in_target_sheet(
        target_sheet, merged_dict, source_col_idx, target_col_idx
    )

    # Copy column width and hidden property
    source_dim = source_sheet.column_dimensions[source_col_letter]
    target_dim = target_sheet.column_dimensions[target_col_letter]
    if hasattr(source_dim, "width"):
        target_dim.width = source_dim.width
    if hasattr(source_dim, "hidden"):
        target_dim.hidden = source_dim.hidden
    # for key, value in source_dim.__dict__.items():
    #     if key in [
    #         "width",
    #         "hidden",
    #     ]:  # , "min", "max", "hidden", "auto_size", "bestFit"]:
    #         target_dim.__dict__[key] = value

    return end_row_idx


# function to copy the sheet from source to target
def copy_sheet(source_sheet, target_sheet):
    # Copy all columns
    hidden_count = 0
    for col in source_sheet.iter_cols(max_col=min(source_sheet.max_column, 50)):
        col_idx = col[0].column
        # Check if the column is hidden
        col_letter = get_column_letter(col_idx)
        if source_sheet.column_dimensions[col_letter].hidden:
            hidden_count += 1
            if hidden_count > 20:
                # print("More than 20 hidden columns detected. Stopping the loop.")
                break

            continue

        # Copy the column to the same index in the target
        _ = copy_column(source_sheet, target_sheet, col_idx, col_idx)

    # Copy sheet-level attributes
    copy_sheet_attributes(source_sheet, target_sheet)


def copy_sheet_attributes(source_sheet, target_sheet):
    # Copy basic sheet attributes

    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.page_setup = copy(source_sheet.page_setup)
    target_sheet.print_options = copy(source_sheet.print_options)
    target_sheet.auto_filter = copy(source_sheet.auto_filter)
    target_sheet.print_area = source_sheet.print_area
    target_sheet.freeze_panes = source_sheet.freeze_panes

    # Copy row dimensions
    # for rn, source_row in source_sheet.row_dimensions.items():
    #     target_sheet.row_dimensions[rn] = copy(source_row)

    # Copy column dimensions
    # for cn, source_col in source_sheet.column_dimensions.items():
    #     target_sheet.column_dimensions[cn] = copy(source_col)
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))
    target_sheet.protection = copy(source_sheet.protection)


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


def create_summary_price_table(summary_sheet, price_sheet, supplier_names):
    # Extract headers and map columns to suppliers
    headers_dict = {}
    for col in price_sheet.iter_cols():
        col_letter = get_column_letter(col[0].column)
        headers = [(cell.value, cell.row) for cell in col if cell.font.bold]
        headers_dict[col_letter] = headers

    # Identify columns for price data
    price_label_col = None
    max_len = 0
    supplier_cols_dict = defaultdict(list)
    for key, value in headers_dict.items():
        if value and value[0][1] == 1 and value[0][0] in supplier_names:
            supplier_cols_dict[value[0][0]].append(key)
        else:
            if len(value) > max_len:
                max_len = len(value)
                price_label_col = key

    # Aggregate price data for suppliers
    agg_dict = defaultdict(list)
    for key, value in supplier_cols_dict.items():
        for col in value:
            count_not_none = sum(
                1
                for _, row in headers_dict[price_label_col]
                if price_sheet[f"{col}{row}"].value is not None
            )
            if count_not_none / len(headers_dict[price_label_col]) > 0.7:
                agg_dict[key].append(col)

    # Compile summary data
    summary_data = []
    for supplier, cols in agg_dict.items():
        for col in cols:
            col_headers = headers_dict.get(col, [])
            if len(col_headers) < 2:
                continue
            category = col_headers[1][0]
            for cate, row in headers_dict[price_label_col]:
                value = price_sheet[f"{col}{row}"].value
                price_value = re.sub(r"[^\d.]", "", str(value))
                if price_value:
                    summary_data.append(
                        {"Category": category, "Subcategory": cate, supplier: value}
                    )

    # Convert summary data into DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Pivot the DataFrame and reset the index
    if not summary_df.empty:
        merged_df = summary_df.pivot_table(
            index=["Category", "Subcategory"],
            values=[
                col
                for col in summary_df.columns
                if col not in ["Category", "Subcategory"]
            ],
            aggfunc="first",
        ).reset_index()
        # Ensure that only existing suppliers are used
        existing_suppliers = [
            supplier for supplier in supplier_names if supplier in merged_df.columns
        ]
        if not existing_suppliers:
            print("No valid suppliers found in the DataFrame columns.")
            return summary_sheet

        # Filter merged_df with the existing suppliers
        try:
            merged_df_filtered = merged_df.loc[:, existing_suppliers]
            # create a new df to store the max value of the three suppliers
            max_rows = (
                merged_df_filtered.max(axis=1).groupby(merged_df["Category"]).idxmax()
            )
            max_df = merged_df.loc[max_rows].reset_index(drop=True)
            print(merged_df)
        except Exception as e:
            print(f"Error filtering DataFrame: {e}")
            return 0
    else:
        print("The summary_df is empty; ensure your input data is correct.")
        return 0

    # Write the dataframe to the sheet
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        summary_sheet.append(r)

    # Merge cells with the same value in column A
    current_value = None
    start_row = None
    for row in range(2, summary_sheet.max_row + 1):  # Skip header row
        cell_value = summary_sheet.cell(row=row, column=1).value
        if cell_value != current_value:
            if start_row and current_value is not None:
                summary_sheet.merge_cells(
                    start_row=start_row, start_column=1, end_row=row - 1, end_column=1
                )
            current_value = cell_value
            start_row = row
    if start_row and current_value is not None:
        summary_sheet.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=summary_sheet.max_row,
            end_column=1,
        )

    # Format the sheet
    header_fill = PatternFill(
        start_color="4472C4", end_color="4472C4", fill_type="solid"
    )
    font_style_header = Font(b=True, color="FFFFFF")
    border_style = Side(border_style="thin", color="000000")

    # Apply borders and styles to all cells
    for row in summary_sheet.iter_rows(
        min_row=1,
        max_row=summary_sheet.max_row,
        min_col=1,
        max_col=summary_sheet.max_column,
    ):
        for cell in row:
            cell.border = Border(
                left=border_style,
                right=border_style,
                top=border_style,
                bottom=border_style,
            )
            if cell.row == 1:
                cell.fill = header_fill
                cell.font = font_style_header

    # Set the column width to auto based on the content
    for col in range(1, summary_sheet.max_column + 1):
        column = get_column_letter(col)
        max_length = max(
            len(str(summary_sheet.cell(row=row, column=col).value) or "")
            for row in range(1, summary_sheet.max_row + 1)
        )
        summary_sheet.column_dimensions[column].width = max_length + 2
    if not summary_df.empty and not max_df.empty:
        print(max_df)
        # start_row = summary_sheet.max_row + 5
        # chart_data_start_row = start_row + 1

        # Write chart data header
        chart_headers = ["Category"] + list(merged_df.columns[2:])  # Supplier names
        summary_sheet.append(chart_headers)

        # Write chart data (categories and corresponding supplier prices)
        for index, row in max_df.iterrows():
            row_data = [row["Category"]] + list(
                row.iloc[2:]
            )  # Include only relevant columns
            summary_sheet.append(row_data)

        # # Create a bar chart
        # chart = BarChart()
        # chart.type = "col"  # Column chart
        # chart.style = 13
        # chart.grouping = "clustered"  # Set grouping to clustered
        # chart.title = "Supplier Price Comparison by Category"
        # chart.x_axis.title = "Category"
        # chart.y_axis.title = "Price"

        # # Define data range for chart
        # data = Reference(
        #     summary_sheet,
        #     min_row=chart_data_start_row,
        #     max_row=summary_sheet.max_row,
        #     min_col=2,  # Supplier data starts from the second column
        #     max_col=1 + len(agg_dict),  # Number of suppliers
        # )
        # categories = Reference(
        #     summary_sheet,
        #     min_row=chart_data_start_row,
        #     max_row=summary_sheet.max_row,
        #     min_col=1,  # Category column
        # )

        # # Add data and categories to chart
        # chart.add_data(data, titles_from_data=True)
        # chart.set_categories(categories)

        # # Position the chart below the data
        # summary_sheet.add_chart(chart, f"A{summary_sheet.max_row + 2}")
    return 1


st.write("# RFP Files Consolidator")
# Check if suppliers are set up
if "suppliers" not in st.session_state or len(st.session_state.suppliers) == 0:
    st.error("No supplier data found. Please complete the setup first.")
    st.stop()

# Check if document types are configured
if "doc_types" not in st.session_state or len(st.session_state.doc_types) == 0:
    st.error("Document types not configured. Please complete the setup first.")
    st.stop()

# Check if each supplier has the necessary files and names
for i, supplier in enumerate(st.session_state.suppliers):
    if (
        not supplier.get("name")
        or not supplier.get(st.session_state.doc_types[0])
        or not supplier.get(st.session_state.doc_types[1])
    ):
        st.error(
            f"Supplier {i+1} is missing required data (name or template files). Please complete the setup first."
        )
        st.stop()

### Session State Variables Retrieval

event_name = st.session_state.event_name
event_option = st.session_state.event_option
# logo path
st.session_state.logo_path = r"kellanova_logo.png"
doc_type1, doc_type2 = st.session_state.doc_types

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
        st.session_state.suppliers, chosen_sheets_pri_idx, doc_type1
    )
    template_sheets_pri = [wb_template_pri[sheet] for sheet in pricing_sheets_list]
    with st.spinner("Processing... Please wait."):
        if st.session_state.pri_comb_mode == "Side by Side":
            consolidated_pri = side_by_side_combine(
                consolidated_pri, template_sheets_pri, sheets_pri_dict
            )
            supplier_names = list(sheets_pri_dict.keys())
            # iterate over the sheets in consolidated_pri and create a summary sheet
            for sheet in consolidated_pri.worksheets:
                # Skip template or non-price sheets if needed
                if sheet.title in pricing_sheets_list:
                    # Create a new summary sheet
                    summary_sheet = consolidated_pri.create_sheet(
                        title=f"Summary of {sheet.title}"[:30]
                    )
                    status_sum = create_summary_price_table(
                        summary_sheet, sheet, supplier_names
                    )
                    # Move the summary sheet to the leftmost position
                    if status_sum:
                        consolidated_pri._sheets.remove(summary_sheet)
                        consolidated_pri._sheets.insert(0, summary_sheet)
                    else:
                        # remove the sheet if the summary is not created
                        consolidated_pri.remove(summary_sheet)

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


# add a summary option tickbox
summary_option = st.checkbox("Questionnaire summary included", value=False)


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
                consolidated_ques,
                template_sheets_ques,
                sheets_ques_dict,
                summary_option=summary_option,
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
