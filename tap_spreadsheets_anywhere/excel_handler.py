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
            filter_column_indices = [
                index for index, header_cell in enumerate(header_row)
                if header_cell.value and header_cell.value.lower() in [col.lower() for col in filtered_columns]
            ]

            # Skip the row if all specified filter columns are empty
            if all(not row[i].value for i in filter_column_indices):
                LOGGER.debug("Row skipped due to empty values in filtered_columns '%s': %r", filtered_columns, row)
                continue

        for index, cell in enumerate(row):
            header_cell = header_row[index]

            # Determine header value or generate "ColumnX" if empty
            if header_cell.value:
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
                if header_cell.value and header_cell.value in included_columns:
                    pass  # Direct match; include this column
                else:
                    # Process header for substring matching
                    formatted_key = header_cell.value.lower() if header_cell.value else ""
                    formatted_key = re.sub(r"\s+", '_', formatted_key)
                    if not any(inc in formatted_key for inc in included_columns_lower):
                        continue

            # Exclude logic with consistent formatting
            elif excluded_columns:
                formatted_key = re.sub(r"\s+", '_', header_cell.value.lower() if header_cell.value else "")
                if any(exc in formatted_key for exc in excluded_columns_lower):
                    continue  # Exclude matching headers

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
        try:
            active_sheet = workbook[table_spec["worksheet_name"]]
        except Exception as e:
            LOGGER.error("Unable to open specified sheet '"+table_spec["worksheet_name"]+"' - did you check the workbook's sheet name for spaces?")
            raise e
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
