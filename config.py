# libs
import os
from pathlib import Path
import secret_keys

#  TODO: create requirements.txt

# local folder settings - parent folder is main
main_dir = Path(os.path.dirname(os.path.abspath(__file__))).parent.absolute()

#current folder
modules_dir = os.path.dirname(os.path.abspath(__file__))

# settings environments dirs:
temp_dir = os.path.join(main_dir, 'temp')

archive_dir = os.path.join(main_dir, 'archive')
cloud_dir = os.path.join(archive_dir, 'cloud')
sort_dir = os.path.join(archive_dir, 'sorted')
# result_dir = os.path.join(archive_dir, 'result')

tables_dir = os.path.join(temp_dir, 'Таблицы')
control_mo_tables_dir = os.path.join(temp_dir, 'Ответы_Контроль_МО')
delta_tables_dir = os.path.join(temp_dir, 'Дельта')
new_patients_tables_dir = os.path.join(temp_dir, 'Новые_пациенты_с_подозрением')
not_table_kjz_dir = os.path.join(temp_dir, 'КЖЗ')
others_files_dir = os.path.join(temp_dir, 'Прочие_отсортированные_файлы')


def rise_up_project_architecture(temp_dir=temp_dir,
                                 archive_dir=archive_dir,
                                 cloud_dir=cloud_dir,
                                 sort_dir=sort_dir,
                                 tables_dir=tables_dir,
                                 control_mo_tables_dir=control_mo_tables_dir,
                                 delta_tables_dir=delta_tables_dir,
                                 new_patients_tables_dir=new_patients_tables_dir,
                                 not_table_kjz_dir=not_table_kjz_dir,
                                 others_files_dir=others_files_dir):

    """create dirs if need"""
    for d in [temp_dir,
             archive_dir,
             cloud_dir,
             sort_dir,
             tables_dir,
             control_mo_tables_dir,
             delta_tables_dir,
             new_patients_tables_dir,
             not_table_kjz_dir,
             others_files_dir]:
        try:
            os.makedirs(d)
            print(f'{d} was created!')
        except Exception as e: print(f'{d} already available?', e, sep='\n')


# telegram bot settings
bot_token = secret_keys.bot_token
admins_tg_id_list = secret_keys.admins_tg_id_list

# chromdriver connect
for file in os.listdir():
    if file.endswith('chromedriver.exe'):
        driver_source = os.path.abspath(file)
        # print(driver_source)
