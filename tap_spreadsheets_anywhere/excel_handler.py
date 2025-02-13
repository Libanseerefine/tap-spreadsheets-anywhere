import re
import logging
import openpyxl


LOGGER = logging.getLogger(__name__)

def generator_wrapper(reader, encapsulate_with_brackets=False, excluded_columns=None, skip_initial=0, included_columns=None, filtered_columns=None):
    # Define the list of columns to filter on
    filtered_columns = filtered_columns or []
    excluded_columns = [] if excluded_columns is None or included_columns else excluded_columns
    included_columns = included_columns or []

    # Lowercase and format excluded columns for consistent comparison
    included_columns_lower = [col.lower() for col in included_columns]
    excluded_columns_lower = [re.sub(r"\s+", '_', col.lower()) for col in excluded_columns]

    _skip_count = 0
    header_row = None
    for row in reader:
        if _skip_count < skip_initial:
            LOGGER.debug("Skipped (%d/%d) row: %r", _skip_count, skip_initial, row)
            _skip_count += 1
            continue

        to_return = {}
        if header_row is None:
            header_row = row
            continue

        # Check if the row should be skipped based on filtered_columns
        if filtered_columns:
            header_map = {}
            for idx, cell in enumerate(header_row):
                header_val = cell.value if isinstance(cell.value, str) else f"Column{idx+1}"
                header_map[header_val.lower()] = idx

            filter_column_indices = []
            for col in filtered_columns:
                col_lower = col.lower().strip()

                # Case 1: Direct match against header names or "ColumnX"
                if col_lower in header_map:
                    filter_column_indices.append(header_map[col_lower])
                    continue

                # Case 2: "ColumnX" format (e.g. "Column2") -> numeric index
                match = re.match(r"^column(\d+)$", col_lower)
                if match:
                    col_index = int(match.group(1)) - 1
                    if col_index < len(header_row):
                        filter_column_indices.append(col_index)
                        continue
                    else:
                        LOGGER.warning("Filtered column '%s' is out of range.", col)

                LOGGER.warning(
                    "Filtered column '%s' not found in header and not recognized as 'ColumnX'. Ignoring.",
                    col
                )
                
            # Skip the row if all specified filter columns are empty
            if all(not row[i].value for i in filter_column_indices):
                LOGGER.debug("Row skipped due to empty values in filtered_columns '%s': %r", filtered_columns, row)
                continue

        for index, cell in enumerate(row):
            header_cell = header_row[index]

            # Determine header value or generate "ColumnX" if empty
            if isinstance(header_cell.value, str):
                formatted_key = header_cell.value
            else:
                formatted_key = f"Column{index + 1}"  # Column numbering starts from 1

            if encapsulate_with_brackets:
                # Encapsulate with square brackets, leave content unchanged
                formatted_key = f"[{formatted_key}]"
            else:
                # Remove non-word, non-whitespace characters
                formatted_key = re.sub(r"[^\w\s]", '', formatted_key)

                # Replace whitespace with underscores
                formatted_key = re.sub(r"\s+", '_', formatted_key)

                # Convert to lowercase
                formatted_key = formatted_key.lower()

            # Direct exact match logic for included columns
            
            if included_columns:
                if isinstance(header_cell.value, str) and header_cell.value in included_columns:
                    pass  # Direct match; include this column
                else:
                    # Ensure the header is a string; otherwise, use "ColumnX"
                    formatted_key = header_cell.value.lower() if isinstance(header_cell.value, str) else f"Column{index + 1}"

                    # Replace spaces with underscores (only applicable for string headers)
                    formatted_key = re.sub(r"\s+", '_', formatted_key)

                    # Check if formatted_key contains any of the included columns
                    if not any(inc in formatted_key for inc in included_columns_lower):
                        continue

            # Exclude logic with consistent formatting
            elif excluded_columns:
                # Ensure excluded_columns_lower is a set before updating
                excluded_columns_lower = {re.sub(r"\s+", '_', col.lower()) for col in excluded_columns}

                # Add "ColumnX" format *only* for explicitly excluded columns
                excluded_columns_lower.update({f"column{idx+1}" for idx in range(len(header_row)) if f"column{idx+1}" in excluded_columns_lower})

                formatted_key = (re.sub(r"\s+", '_', header_cell.value.lower())
                    if isinstance(header_cell.value, str)
                    else f"column{index+1}")

                column_index_key = f"column{index+1}"  # ColumnX format

                # Exclude if header name or column index (ColumnX) matches
                if formatted_key in excluded_columns_lower or column_index_key in excluded_columns_lower:
                    continue  # Skip excluded columns

            # Final formatting for the output
            formatted_key = re.sub(r"[^\w\s]", '', formatted_key).replace(' ', '_')

            to_return[formatted_key] = cell.value

        yield to_return


def get_row_iterator(table_spec, file_handle):
    encapsulate_with_brackets = table_spec.get('encapsulate_with_brackets', False)
    excluded_columns = table_spec.get('excluded_columns', [])
    included_columns = table_spec.get('included_columns', [])
    skip_initial = table_spec.get("skip_initial", 0)
    filtered_columns = table_spec.get("filtered_columns", 0)
    workbook = openpyxl.load_workbook(file_handle.name, read_only=True, data_only = True)
    if "worksheet_name" in table_spec:
        possible_sheet_names = [name.strip() for name in table_spec["worksheet_name"].split(',')]
        active_sheet = None
        
        for sheet_name in possible_sheet_names:
            if sheet_name in workbook.sheetnames:
                # Found a matching worksheet in the workbook
                active_sheet = workbook[sheet_name]
                break

        if not active_sheet:
            # None of the candidate sheets were found
            LOGGER.error(
                "Unable to open any of the specified sheets '%s'. "
                "Did you check for typos or extra spaces?",
                table_spec["worksheet_name"],
            )
            raise ValueError(
                f"No valid sheet found in workbook from list: {table_spec['worksheet_name']}"
            )
    else:
        try:
            worksheets = workbook.worksheets
            if len(worksheets) == 1:
                active_sheet = worksheets[0]
            else:
                max_row = 0
                longest_sheet_index = 0
                for i, sheet in enumerate(worksheets):
                    if sheet.max_row > max_row:
                        max_row = i.max_row
                        longest_sheet_index = i
                active_sheet = worksheets[longest_sheet_index]
        except Exception as e:
            LOGGER.info(e)
            active_sheet = worksheets[0]
    return generator_wrapper(active_sheet, encapsulate_with_brackets, excluded_columns, skip_initial, included_columns, filtered_columns)
