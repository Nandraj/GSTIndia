import datetime
import xlrd
import json
import PySimpleGUI as sg


def get_date_in_format_from_excel_cell(cell_value, workbook):
    # Convert excel cell date data into relevant format for json generation
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


def get_month_in_format(cell_value):
    cell_value_str = str(int(cell_value))
    if len(cell_value_str) < 6:
        return "0" + cell_value_str
    else:
        return cell_value_str


def get_master_details(book):
    master_sheet = book.sheet_by_name("master")
    GSTIN = master_sheet.cell_value(rowx=0, colx=1)
    fromFp = get_month_in_format(master_sheet.cell_value(rowx=1, colx=1))
    toFp = get_month_in_format(master_sheet.cell_value(rowx=2, colx=1))
    return [GSTIN, fromFp, toFp]


def is_int(n):
    try:
        float_n = float(n)
        int_n = int(float_n)
    except ValueError:
        return False
    else:
        return float_n == int_n


def process_row_data(row_value_list, id_list, workbook):
    invoice_dict = {}
    for index, field in enumerate(id_list):
        # integers
        if field in ["sno"]:
            invoice_dict[field] = int(row_value_list[index])
        # tax field if nil then 0 else float given number
        elif field in ["val", "iamt", "camt", "samt", "oval", "oiamt", "ocamt", "osamt"]:
            if row_value_list[index] in ['', 0.0]:
                invoice_dict[field] = 0
            else:
                if is_int(row_value_list[index]) == True:
                    invoice_dict[field] = int(row_value_list[index])
                else:
                    invoice_dict[field] = round(
                        float(row_value_list[index]), 2)
        # dates
        elif field in ["idt", "oidt"]:
            invoice_dict[field] = get_date_in_format_from_excel_cell(
                row_value_list[index], workbook)
        # invoice numbers
        elif field in ["inum", "oinum"]:
            try:
                invoice_dict[field] = str(int(row_value_list[index]))
            except:
                invoice_dict[field] = row_value_list[index]
        #istype and ostype
        elif field in ["istype", "ostype", "idtype", "odtype"]:
            invoice_dict[field] = str(row_value_list[index])
        # remaining normal strings
        else:
            invoice_dict[field] = str(row_value_list[index]).upper()
    return invoice_dict


def process_data_sheet_and_return_dictionery(book):
    full_id_list = ["sno", "istype", "stin", "idtype", "inum", "idt", "portcd", "val", "iamt",
                    "camt", "samt", "ostype", "odtype", "oinum", "oidt", "oval", "oiamt", "ocamt", "osamt"]
    first_half_id_list = ["sno", "istype", "stin", "idtype",
                          "inum", "idt", "portcd", "val", "iamt", "camt", "samt"]
    second_half_id_list = ["ostype", "odtype", "oinum",
                           "oidt", "oval", "oiamt", "ocamt", "osamt"]
    data_sheet = book.sheet_by_name("stmt1A")
    num_rows = data_sheet.nrows
    main_dict = {}
    master_details = get_master_details(book)
    main_dict["gstin"] = master_details[0]
    main_dict["fromFp"] = master_details[1]
    main_dict["toFp"] = master_details[2]
    main_dict["refundRsn"] = "INVITC"
    main_dict["version"] = "1.3"
    stmt01a_dict_list = []
    for n in range(2, num_rows):
        row_value_list = [cell for cell in data_sheet.row_values(n)]
        # print(row_value_list)
        if row_value_list[11] == "":
            processed_dict = process_row_data(
                row_value_list, first_half_id_list, book)
            # print(processed_dict)
            stmt01a_dict_list.append(processed_dict)
        elif row_value_list[1] == "":
            processed_dict = process_row_data(
                row_value_list[11:], second_half_id_list, book)
            # print(processed_dict)
            stmt01a_dict_list.append(processed_dict)
        else:
            processed_dict = process_row_data(
                row_value_list, full_id_list, book)
            # print(processed_dict)
            stmt01a_dict_list.append(processed_dict)
    main_dict["stmt01A"] = stmt01a_dict_list
    return main_dict


def get_excel_file_gui_window():
    xl_file = sg.PopupGetFile(
        "********* NR *********\nStatement 1A Json Generation Utitility\nSelect Stmt 1A Excel Template", grab_anywhere=True, no_titlebar=True)
    return xl_file


def main():
    # file selector
    xl_file = get_excel_file_gui_window()
    xl_file_name = xl_file.split("/")[-1].split(".")[0]
    xl_file_path = ""
    for index, value in enumerate(xl_file.split("/")[:-1]):
        if index == 0:
            xl_file_path += value
        else:
            xl_file_path += "\\"
            xl_file_path += value
    try:
        book = xlrd.open_workbook(xl_file)
        final_dict = process_data_sheet_and_return_dictionery(book)
        with open(xl_file_path + '\\' + xl_file_name + '.json', 'w') as file:
            json.dump(final_dict, file, indent=None, separators=(',', ':'))
        sg.PopupOK("Json Generated Successfully!")
    except:
        # print(e)
        sg.PopupError("Error Occurred")


if __name__ == "__main__":
    main()
