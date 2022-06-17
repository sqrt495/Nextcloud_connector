import xlwings as xw
import psutil
import zipfile
from pathlib import Path
from datetime import datetime
import shutil
import config
import os

today = datetime.now().strftime('%d.%m.%Y')

# TODO: RENAME dictionary = check_handler_dictionary
check_handler_dictionary ={}
for i in ["Дата направления запроса в МО",
          "Дата направления ответа от МО в Дирекцию МГКР",
          "Комментарий персонального помощника",
          "Ответ от МО"]:
    check_handler_dictionary[i] = "Ответы_Контроль_МО"
for i in ["Группа проблем",
          "Тип ошибки",
          "Назначить помощника",
          "Результат от МО"]:
    check_handler_dictionary[i] = "Дельта"
for i in ["ФИО персонального помощника (выбирается из списка только в случае подписания соглашения на сопровождение)",
          "Подписано соглашение на сопровождение пациента персональными помощниками",
          "Дата записи к онкологу (при наличии)"]:
    check_handler_dictionary[i] = "Новые_пациенты_с_подозрением"



# перед формированием словаря, где ключ = путь_к_паке, а значения = список имен файлов в папке
# формирование словаря, где ключ = путь_к_паке, а значения = список имен файлов в папке
def unpacked_tree(extract_dir):
    unpacked_dict = {}
    unpacked_list = []
    unpacked_dir_tree = os.walk(extract_dir)
    for i in unpacked_dir_tree:
        dir_structure = list(i)
        if dir_structure[-1]:
            unpacked_dict[dir_structure[0]+'\\'] = dir_structure[-1]
            for f in dir_structure[-1]:
                unpacked_list.append(dir_structure[0]+ '\\' + f)
    return unpacked_dict, unpacked_list


########### DEF unpack_zipfile #################################
def unpack_zipfile(dir_with_archive, extract_dir, encoding='cp866'):
    for file in os.listdir(dir_with_archive):
        if '.zip' in file:
            with zipfile.ZipFile(dir_with_archive+'\\'+file) as archive:
                for entry in archive.infolist():
                    try:
                        name = entry.filename.encode('cp437').decode(encoding)
                    except:
                        name = entry.filename.encode('cp866').decode('cp866')
                    target = os.path.join(extract_dir, *name.split('/'))
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    if not entry.is_dir():  # file
                        with archive.open(entry) as source, open(target, 'wb') as dest:
                            shutil.copyfileobj(source, dest)
            os.remove(dir_with_archive+'\\'+file)

# поиск архивов в папке unzip
# и разархивирование их в папку 'unzip'
def find_nested_zip(unpacked_dict, extract_dir):
    try:
        for k,v in unpacked_dict.items():
            res = [string for string in v if '.zip' in string]
            for _ in res:
                match = os.path.join(k, _)
                match = os.path.abspath(match)
                if match:
                    unpack_zipfile(k, extract_dir, encoding='cp866')
                    os.remove(os.path.abspath(match))
    except:
        print('no zip in dir')

def rename_files_by_folder(unpacked_list):
    for path_to_file in unpacked_list:
        name_listing = path_to_file.split('\\')
        res = any('кжз' in string.lower() for string in name_listing)
        kjz_str = 'КЖЗ-'
        if res:
            name_listing[-1] = kjz_str+name_listing[-2]+'-'+name_listing[-1]
        else:
            name_listing[-1] = name_listing[-2]+'-'+name_listing[-1]

        new_path_to_file = '\\'.join([str(n) for n in name_listing]).replace(kjz_str*2,
                                                                             kjz_str)
        new_path_to_file = new_path_to_file.replace(name_listing[-2]+'-'+name_listing[-2],
                                                    name_listing[-2])
        os.rename(path_to_file,
                  new_path_to_file)

def sort_files(original_folder_name,
               extract_dir,
               unpacked_list,
               not_table_file_types,
               table_file_types,
               not_table_kjz_dir,
               tables_dir,
               others_files_dir):
    for path_to_file in unpacked_list:
        name_listing = path_to_file.split('\\')
        res = any('кжз' in string.lower() for string in name_listing)
        try:
            if res:
                if path_to_file.endswith(tuple(not_table_file_types)):
                    try:
                        os.makedirs(not_table_kjz_dir)
                    except:
                        pass
                    os.rename(path_to_file, Path(config.not_table_kjz_dir + '\\' + name_listing[-1]))
                else:
                    try:
                        os.makedirs(tables_dir)
                    except:
                        pass
                    os.rename(path_to_file, Path(config.tables_dir + '\\' + name_listing[-1]))
            elif path_to_file.endswith(tuple(table_file_types)):
                try:
                    os.makedirs(tables_dir)
                except:
                    pass
                os.rename(path_to_file, Path(config.tables_dir + '\\' + name_listing[-1]))
            else:
                path_to_file.endswith(tuple(not_table_file_types))
                try:
                    os.makedirs(others_files_dir)
                except:
                    pass
                os.rename(path_to_file, Path(config.others_files_dir + '\\' + name_listing[-1]))
        except:
            print(path_to_file)
    shutil.rmtree(extract_dir+f'\\{original_folder_name}\\')


def unzip(file, unpack_dir):
    with zipfile.ZipFile(file) as zip:
        for zip_info in zip.infolist():
            if zip_info.filename[-1] == '/':
                continue
            zip_info.filename = os.path.basename(zip_info.filename)
            zip.extract(zip_info, unpack_dir)


def replacer(current_dir, new_dir, name, method):
    try:
        os.mkdir(os.path.join(new_dir, method))
        os.replace(os.path.join(current_dir, name), os.path.join(new_dir, method, name))
    except:
        os.replace(os.path.join(current_dir, name), os.path.join(new_dir, method, name))


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
        return 'Прочие_отсортированные_файлы'  # изменил "Прочее" на "Не таблицы"
    except:
        return 'Прочие_отсортированные_файлы'  # изменил "Прочее" на "Не таблицы"


def trash_sort(path, file):
    if file.split('.')[-1] in ['pdf', 'jpeg', 'jpg', 'png', 'zip', 'PDF', 'JPEG', 'JPG', 'PNG',
                               'ZIP']:  # добавил в список ", 'PDF', 'JPEG', 'JPG', 'PNG', 'ZIP'"
        replacer(path, file, 'Не таблицы')  # изменил "Прочее" на "Не таблицы"


def die():
    try:
        for proc in psutil.process_iter():
            if proc.name() == "EXCEL.EXE":
                proc.kill()
    except:
        pass