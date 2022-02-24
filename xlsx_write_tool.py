import datetime
import xlsxwriter
import os


def _write_to_xlsx(sig_data: list, sh_name: list, header_name_list: list):

    date = datetime.datetime.now().date()
    workbook = xlsxwriter.Workbook(
        os.path.join('.', 'exported_data', f"PathwayTraversal_{date.strftime('%Y%m%d')}.xlsx"))
    col_len_width = []
    for idx, data_ in enumerate(zip(sig_data, sh_name)):
        header_format = workbook.add_format(
            {'align': 'center', 'valign': 'vcenter', 'size': 12, 'color': 'black', 'bold': 4})
        header = []
        for h in header_name_list:
            header.append({'header': h, "format": header_format})

        # print(sheet_name)
        sheet = workbook.add_worksheet(data_[1])
        if len(data_[0]):

            sheet.add_table(0, 0, len(data_[0]), len(data_[0][0]) - 1,
                            {'data': sorted(data_[0], key=lambda x: x[3], reverse=True), 'autofilter': True,
                             'columns': header})

            for m in range(len(data_[0]) + 1):
                sheet.set_row(m, 30, cell_format=header_format)

            col_len_width.append([len(j) for j in header_name_list])
            for n, l in enumerate(zip(col_len_width[idx], header_name_list)):
                sheet.set_column(n, n, max(l[0], len(l[1])) * 5)
    workbook.close()

    return os.path.join('.', 'exported_data', f"PathwayTraversal_{date.strftime('%Y%m%d')}.xlsx")