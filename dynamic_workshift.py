import xlwings as xw


def main():
    # Connect to the active workbook and sheet
    wb = xw.Book.caller()  # Connect to the calling workbook
    sheet = wb.sheets.active  # Use the active sheet

    # Define ranges
    fd_limit = sheet.range("B2").value  # FD Soll
    sd_limit = sheet.range("B3").value  # SD Soll
    input_range = sheet.range("B6:D8")  # Input area
    output_range = sheet.range("B11:D13")  # Output area

    # Loop through each column (date) in the output range
    for col_idx in range(input_range.columns.count):
        fd_count = 0
        sd_count = 0

        # Calculate existing FD and SD counts from the input range
        for cell in input_range.columns[col_idx]:
            if cell.value == "FD":
                fd_count += 1
            elif cell.value == "SD":
                sd_count += 1

        # Fill the output range dynamically
        for cell in output_range.columns[col_idx]:
            input_cell = input_range[cell.row - output_range.row, col_idx]
            if input_cell.value:
                # Copy value from input if present
                cell.value = input_cell.value
            elif fd_count < fd_limit:
                # Fill with FD if FD quota is not met
                cell.value = "FD"
                fd_count += 1
            elif sd_count < sd_limit:
                # Fill with SD if SD quota is not met
                cell.value = "SD"
                sd_count += 1
            else:
                # Leave blank if quotas are met
                cell.value = ""


if __name__ == "__main__":
    xw.Book("dynamic_workshift.xlsm").set_mock_caller()
    main()
