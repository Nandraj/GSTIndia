import xlrd
import xlsxwriter
import datetime

col_list_std_format = [
    "GSTIN",
    "NAME",
    "INOVICE NO",
    "INVOICE TYPE",
    "INOVICE DT",
    "VALUE",
    "POS",
    "REV. CHRG",
    "RATE",
    "TAX. VAL",
    "IGST",
    "CGST",
    "SGST",
    "CESS",
    "FILLING STATUS",
]
col_head_not_in_scigst_format = ["INVOICE TYPE", "POS"]


def get_date_in_format_from_excel_cell(cell_value, workbook):
    try:
        cell_value_as_datetime = datetime.datetime(
            *xlrd.xldate_as_tuple(cell_value, workbook.datemode))
        date_in_format = cell_value_as_datetime.strftime("%d-%m-%Y")
    except Exception:
        pass
    try:
        date_in_format = datetime.datetime.strptime(
            cell_value, "%d-%b-%Y").strftime("%d-%m-%Y")
    except Exception:
        pass
    try:
        date_in_format = datetime.datetime.strptime(
            cell_value, "%d/%b/%Y").strftime("%d-%m-%Y")
    except Exception:
        pass
    try:
        date_in_format = datetime.datetime.strptime(
            cell_value, "%d-%m-%Y").strftime("%d-%m-%Y")
    except Exception:
        pass
    try:
        date_in_format = datetime.datetime.strptime(
            cell_value, "%d/%m/%Y").strftime("%d-%m-%Y")
    except Exception:
        pass
    try:
        date_in_format = datetime.datetime.strptime(
            cell_value, "%d.%m.%Y").strftime("%d-%m-%Y")
    except Exception:
        pass
    return date_in_format


def parse_excel_file_and_return_list_of_dict(workbook):
    sheet = workbook.sheet_by_index(0)
    num_rows = sheet.nrows

    col_head = [cell for cell in sheet.row_values(0)]

    data_list = []

    for n in range(1, num_rows):
        row_values_list = [cell for cell in sheet.row_values(n)]
        row_data_dict = {}
        for i in col_head[1:]:
            row_data_dict[i] = row_values_list[col_head.index(i)]
        data_list.append(row_data_dict)
    return data_list


def get_list_of_dict_from_scigst_and_generate_list_of_row_values(
    list_of_dict, col_list_std_format, col_head_not_in_scigst_format, workbook
):
    list_of_row_value = []
    list_of_row_value.append(col_list_std_format)
    for dict_ in list_of_dict:
        temp_raw_list = []
        for col_head in col_list_std_format:
            if col_head in col_head_not_in_scigst_format:
                if col_head == "INVOICE TYPE":
                    temp_raw_list.append("R")
                elif col_head == "POS":
                    temp_raw_list.append("Gujarat")
            elif col_head == "NAME":
                temp_raw_list.append(dict_[col_head].rstrip())
            elif col_head == "FILLING STATUS":
                if dict_[col_head] == "Y":
                    temp_raw_list.append("Submitted")
                else:
                    temp_raw_list.append("Not-Submitted")
            elif col_head == "INOVICE DT":
                temp_raw_list.append(
                    get_date_in_format_from_excel_cell(
                        dict_[col_head], workbook)
                )
            else:
                temp_raw_list.append(dict_[col_head])
        list_of_row_value.append(temp_raw_list)
    return list_of_row_value


def generate_std_format_xl_file(xl_file, list_of_rows):
    workbook = xlsxwriter.Workbook(xl_file)
    sheet = workbook.add_worksheet("Sheet1")
    header_format = workbook.add_format(
        {"align": "center", "valign": "vcenter", "bold": True}
    )
    bold_format = workbook.add_format({"bold": True})
    head_col = 0
    for header in list_of_rows[0]:
        sheet.write(0, head_col, header, header_format)
        head_col += 1

    row_no = 1
    for row_values in list_of_rows[1:]:
        col_no = 0
        for col_value in row_values:
            if col_no in [10, 11, 12, 13]:
                if col_value == "":
                    sheet.write(row_no, col_no, 0)
                else:
                    sheet.write(row_no, col_no, col_value)
            else:
                sheet.write(row_no, col_no, col_value)
            if col_no == 2:
                try:
                    sheet.write(
                        row_no +
                        1, col_no, str(int(col_value)) + "-Total", bold_format
                    )
                except:
                    sheet.write(
                        row_no + 1, col_no, str(col_value) +
                        "-Total", bold_format
                    )
            elif col_no == 8:
                sheet.write(row_no + 1, col_no, "-", bold_format)
            elif col_no in [10, 11, 12, 13]:
                if col_value == "":
                    sheet.write(row_no + 1, col_no, 0, bold_format)
                else:
                    sheet.write(row_no + 1, col_no, col_value, bold_format)
            else:
                sheet.write(row_no + 1, col_no, col_value, bold_format)
            col_no += 1
        row_no += 3
    workbook.close()


def gstr2_xl_generator(input_file, output_folder):
    try:
        xl_file_name_with_ext = input_file.split("/")[-1]  # .split(".")[0]
        out_file = output_folder + xl_file_name_with_ext
        book = xlrd.open_workbook(input_file)  # , ragged_rows=True)
        list_of_dict = parse_excel_file_and_return_list_of_dict(book)
        list_of_rows = get_list_of_dict_from_scigst_and_generate_list_of_row_values(
            list_of_dict, col_list_std_format, col_head_not_in_scigst_format, book
        )
        generate_std_format_xl_file(out_file, list_of_rows)
    except Exception as e:
        print(e)
