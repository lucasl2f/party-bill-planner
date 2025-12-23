import xlsxwriter


def create_excel(servico, credores):
    # Create a new workbook and add a worksheet
    workbook = xlsxwriter.Workbook("my_spreadsheet.xlsx")
    worksheet = workbook.add_worksheet()

    # Define formats
    header_format = workbook.add_format(
        {
            "bg_color": "#82cfe8",  # same as your PatternFill
            "bold": True,
            "align": "center",
            "valign": "vcenter",
        }
    )

    currency_format = workbook.add_format(
        {
            "num_format": "R$ #,##0.00",
        }
    )

    gray_row_format = workbook.add_format(
        {
            "bg_color": "#DDDDDD",
        }
    )

    # Set headers
    worksheet.write("A1", "Serviço", header_format)
    worksheet.write("B1", "Credor", header_format)

    # Set column width
    worksheet.set_column("A:A", 50)
    worksheet.set_row(0, 30)  # row 1 (0-indexed) height

    # Calculate column letters
    num_credores = len(credores)
    end_credor_col = 1 + num_credores  # B + N credores
    total_col = end_credor_col + 1  # Total column
    participants_col = total_col + 1  # Participants start
    participants_total_col = (
        participants_col + num_credores - 1 + 1
    )  # Participants Total
    balance_col = participants_total_col + 1  # Balance start
    balance_end_col = balance_col + num_credores - 1  # Balance end

    # Merge cells for Credor header
    worksheet.merge_range(0, 1, 0, end_credor_col, "Credor", header_format)

    # Total column
    worksheet.write(0, total_col, "Total", header_format)

    # Participants header
    worksheet.merge_range(
        0,
        participants_col,
        0,
        participants_total_col - 1,
        "Participantes",
        header_format,
    )
    worksheet.write(0, participants_total_col, "Total", header_format)

    # Balance header
    worksheet.merge_range(
        0, balance_col, 0, balance_end_col, "Balanço geral", header_format
    )

    # Add Credores (row 2, 0-indexed row 1)
    for idx, nome in enumerate(credores):
        col = 1 + idx
        worksheet.write(1, col, nome, header_format)
        worksheet.set_column(col, col, 12)

    # Add Participants (row 2, same row as credores)
    for idx, nome in enumerate(credores):
        col = participants_col + idx
        worksheet.write(1, col, nome, header_format)
        worksheet.set_column(col, col, 12)

    # Add Balance credores (row 2)
    for idx, nome in enumerate(credores):
        col = balance_col + idx
        worksheet.write(1, col, nome, header_format)
        worksheet.set_column(col, col, 12)

    # Total formula for each row (row 4 to 50, 0-indexed row 3 to 49)
    total_col_letter = xlsxwriter.utility.xl_col_to_name(total_col)
    for row in range(3, 50):
        start_col_letter = xlsxwriter.utility.xl_col_to_name(1)
        end_col_letter = xlsxwriter.utility.xl_col_to_name(end_credor_col)
        formula = f"=SUM({start_col_letter}{row + 1}:{end_col_letter}{row + 1})"
        worksheet.write_formula(row, total_col, formula, currency_format)

    # Sum of Total column (row 3, 0-indexed row 2)
    sum_total_formula = f"=SUM({total_col_letter}4:{total_col_letter}10000)"
    worksheet.write_formula(2, total_col, sum_total_formula, currency_format)

    # Participants Total formula for each row
    participants_total_col_letter = xlsxwriter.utility.xl_col_to_name(
        participants_total_col
    )
    participants_start_col_letter = xlsxwriter.utility.xl_col_to_name(participants_col)
    participants_end_col_letter = xlsxwriter.utility.xl_col_to_name(
        participants_total_col - 1
    )
    for row in range(3, 50):
        formula = f"=SUM({participants_start_col_letter}{row + 1}:{participants_end_col_letter}{row + 1})"
        worksheet.write_formula(row, participants_total_col, formula)

    # Sum of Participants Total column
    sum_total_participants_formula = (
        f"=SUM({participants_total_col_letter}4:{participants_total_col_letter}10000)"
    )
    worksheet.write_formula(2, participants_total_col, sum_total_participants_formula)

    # Balance formulas
    total_col_letter = xlsxwriter.utility.xl_col_to_name(total_col)
    participants_total_col_letter = xlsxwriter.utility.xl_col_to_name(
        participants_total_col
    )
    for i in range(num_credores):
        row_credor_letter = xlsxwriter.utility.xl_col_to_name(1 + i)
        row_participant_letter = xlsxwriter.utility.xl_col_to_name(participants_col + i)
        balance_col_idx = balance_col + i
        balance_col_letter = xlsxwriter.utility.xl_col_to_name(balance_col_idx)

        # Sum formula for balance header (row 3, 0-indexed row 2)
        sum_balance_formula = f"=SUM({balance_col_letter}4:{balance_col_letter}50)"
        worksheet.write_formula(
            2, balance_col_idx, sum_balance_formula, currency_format
        )

        # Balance formula for each row
        for row in range(3, 50):
            formula = f"=IFS({row_participant_letter}{row + 1} >= 1, {row_credor_letter}{row + 1}-({row_participant_letter}{row + 1}*(${total_col_letter}{row + 1}/${participants_total_col_letter}{row + 1})), {row_participant_letter}{row + 1} = 0, {row_credor_letter}{row + 1})"
            worksheet.write_formula(row, balance_col_idx, formula)

    # Apply gray fill to row 3 (0-indexed row 2)
    for col in range(0, balance_end_col + 1):
        worksheet.write(2, col, "", gray_row_format)

    # Close the workbook
    workbook.close()

    print(
        f"Excel file created with structure: Serviço, Credor, {', '.join(credores)}, Total"
    )


# Example usage
servico = "Consultoria"
credores = ["John", "Jane", "Alice", "Juquinha"]
create_excel(servico, credores)
