from openpyxl import load_workbook
import json
import urllib3


def read_excel(excel_file):
    out_struct = {}
    wb = load_workbook(excel_file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_arr = []
        for cell in row[1:]:
            if cell:
                row_arr.append(cell)
        out_struct[row[0]] = row_arr

    return out_struct


def download_art_photo(link_struct, path='.', folders=False):
    http = urllib3.PoolManager()

    for art, photos_arr in link_struct.items():
        files_iterator = 0
        for photo in photos_arr:
            files_iterator += 1
            file_format = 'jpg'
            filename_art = art + '_' + str(files_iterator) + '.' + file_format

            r = http.request('GET', photo, preload_content=False)

            with open(filename_art, 'wb+') as out:
                while True:
                    data = r.read(65536)  # 2**16
                    if not data:
                        break
                    out.write(data)

            r.release_conn()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    file = "input.xlsx"
    file_struct = read_excel(file)
    download_art_photo(file_struct)

    # with open('temp.txt', 'w+') as f:
    #    f.write(json.dumps(file_struct))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
