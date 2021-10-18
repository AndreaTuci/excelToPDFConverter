import os
from win32com import client

excel = client.Dispatch("Excel.Application")

# print(filter(lambda x: x.endswith(('.txt','.py')), os.listdir(os.curdir)))
file_list = []
for f in os.listdir(os.curdir):
    if f.endswith(('.xlsx', '.xls', '.ods')):
        file_list.append(f)

main_dir = os.path.dirname(os.path.realpath('__file__'));


print(f'Cartella di lavoro: {main_dir}. Cerco fogli di calcolo')

for file in file_list:
    try:
        excel_file_path = os.path.join(main_dir, file)
        print(f'File {file} trovato!')

        if file.endswith('.xlsx'):
            new_dir = file[:-5]
        else:
            new_dir = file[:-4]

        new_dir = os.path.join(main_dir, new_dir)
        os.makedirs(new_dir)

        sheets = excel.Workbooks.Open(excel_file_path)

        i = 0

        while i < len(sheets.Worksheets):
            try:
                work_sheet = sheets.Worksheets[i]
                title = work_sheet.Name.replace(" ", "_")
                pdf_file_path = os.path.join(new_dir, f'{title}.pdf')
                work_sheet.ExportAsFixedFormat(0, pdf_file_path)
                print(f'Foglio {work_sheet.Name} convertito')
            except:
                print(f'Errore per il foglio n.{i}')
            i += 1

        sheets.Close(True)
        print('-'*45)

    except:
        print('Non trovo il file!')

