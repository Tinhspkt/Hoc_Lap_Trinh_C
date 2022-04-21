import os
from os import path
import xlwings as xw

def create_name_in_chill_wb(wb_path):
    wb = xw.Book(wb_path)
    headers_map = {
        'day_lan': 'Dây cáp mạng cat6 UTP',
        'day_loa': 'Cáp 18AWG 1PR',
        'day_quang': 'Dây cáp quang 4Fo',
        'day_dien': 'Dây dẫn 2 ruột Cu/PVC 2x1,5mm2 Cadivi',
        'mang_be': 'PVC 24x14 mm',
        'mang_to': 'PVC 60x40 mm',
        'ong_d20': 'PVC D20',

    }
    index_col = 'B'  # CỘT CHỨA TIÊU ĐỀ CẦN SO SÁNH ĐỂ LẤY KHỐI LƯỢNG
    quantity_col = 'K'  # CỘT CHỨA KHỐI LƯỢNG

    sht = wb.sheets[0]
    quantities = []
    for k, v in headers_map.items():
        # tối đa 1000 dòng
        for i in range(6, 1000):
            index_col_value = sht.range('%s%s' % (index_col, i)).value
            quantity_col_value = sht.range('%s%s' % (quantity_col, i)).value
            if index_col_value and v in index_col_value:
                quantities.append({
                    'type': k,
                    'quantity': quantity_col_value or 0
                })

                wb.api.Names.Add(
                    Name=k,
                    RefersTo=sht.range('%s%s' % (quantity_col, i)).api
                )
                break
    wb.save()
    wb.close()


def main():
    folder = r'D:\HTL\Desktop\CAMERA 70QN BVHC\test_collect'
    main_xls = os.path.join(folder, 'thong ke khoi luong.xlsx')
    main_wb = xw.Book(main_xls)
    main_sht = main_wb.sheets[0]

    # CỘT CHỨA TÊN FILE XLS DIỄN GIẢI TƯƠNG ỨNG
    xls_filename_column = 'AN'

    data = []

    min_row = 2
    max_row = 84
    for i in range(min_row, max_row + 1):
        if main_sht.range('C%d' % i).value and main_sht.range('%s%s' % (xls_filename_column, i)).value:
            data.append({
                'row_index': i,
                'xls_file': main_sht.range('%s%s' % (xls_filename_column, i)).value
            })
    for n in data:
        xls_path = os.path.join(folder, '%s.xls' % n['xls_file'])
        create_names_in_child_wb(xls_path)


if __name__ == '__main__':
    main()