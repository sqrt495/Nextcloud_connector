import zipfile
import os
import xlwings as xw
import psutil

def unzip(file, unpack_dir):
    with zipfile.ZipFile(file) as zip:
        for zip_info in zip.infolist():
            if zip_info.filename[-1] == '/':
                continue
            zip_info.filename = os.path.basename(zip_info.filename)
            zip.extract(zip_info, unpack_dir)

def replacer(path, name, method):
    try:
        os.mkdir(os.path.join(path, method))
        os.replace(os.path.join(path, name), os.path.join(path, method, name))
    except:
        os.replace(os.path.join(path, name), os.path.join(path, method, name))

def resaver(path, file):
    dir_file = os.path.join(path, file)
    try:
        excel_app = xw.App(visible=False)
        excel_app.display_alerts = False
        wb = xw.Book(dir_file)
        new_name = file
        wb.save(os.path.join(path, new_name))
        wb.close()
        # xw.apps.active.api.quit()
        if file != new_name:
            os.remove(dir_file)
    except:
        try:
            excel_app = xw.App(visible=False)
            excel_app.display_alerts = False
            wb = xw.Book(dir_file, corrupt_load=2)
            new_name = "MOD - " + file
            wb.save(os.path.join(path, new_name))
            wb.close()
            os.remove(dir_file)
        except:
            pass


# rewrite it
def check_headers(path, file, dict):
    file = os.path.join(path, file)
    try:
        excel_app = xw.App(visible=False)
        excel_app.display_alerts = False
        wb = xw.Book(file)
        mass = []
        for ws in wb.sheets:
            for i in range(1, 20):
                try:
                    values = ws.range((i, 1), (i, 50)).value
                    for value in values:
                        if value == 'Наименование "Павильона Здоровья" в парке, где выявлено подозрение на ЗНО':
                            wb.close()
                            return "ЦАОП"
                        mass.append(value)
                except:
                    pass

        wb.close()
        # xw.apps.active.api.quit()

        for i in mass:
            if i in dict.keys():
                return dict[i]
        return 'Контроль_МО-Прочие_отсортированные_файлы'
    except:
        return 'Контроль_МО-Прочие_отсортированные_файлы'

def trash_sort(path, file):
    if file.split('.')[-1] in ['pdf', 'jpeg', 'jpg', 'png', 'zip', 'PDF', 'JPEG', 'JPG', 'PNG', 'ZIP']:
        replacer(path, file, 'Контроль_МО-Прочие_отсортированные_файлы')

def die():
    try:
        for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                proc.kill()
    except:
        pass