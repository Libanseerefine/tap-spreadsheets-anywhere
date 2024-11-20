import re
import logging
import openpyxl


LOGGER = logging.getLogger(__name__)

def generator_wrapper(reader, encapsulate_with_brackets=False):
    header_row = None
    for row in reader:
        to_return = {}
        if header_row is None:
            header_row = row
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

            to_return[formatted_key] = cell.value

        yield to_return


def get_row_iterator(table_spec, file_handle):
    encapsulate_with_brackets = table_spec.get('encapsulate_with_brackets', False)
    workbook = openpyxl.load_workbook(file_handle.name, read_only=True)
    if "worksheet_name" in table_spec:
        try:
            active_sheet = workbook[table_spec["worksheet_name"]]
        except Exception as e:
            LOGGER.error("Unable to open specified sheet '"+table_spec["worksheet_name"]+"' - did you check the workbook's sheet name for spaces?")
            raise e
    else:
        try:
            worksheets = workbook.worksheets
            if(len(worksheets) == 1):
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
    return generator_wrapper(active_sheet, encapsulate_with_brackets)
