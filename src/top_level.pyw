
import threading
from pathlib import Path
import datetime
import os

import directory_fetcher
import user_interface.UI_master


load_regnskapsprogram_thrd = None
regnskapsprogram = None
program_is_running = False

def load_regnskapsprogram():
    global regnskapsprogram
    import regnskapsprogram as regnskapsprogram_tmp
    regnskapsprogram = regnskapsprogram_tmp


def call_main_program(*args):
    global program_is_running
    global load_regnskapsprogram_thrd
    global regnskapsprogram

    if program_is_running:
        success = False
        message = 'Kan ikke kjøre flere kall til programmet samtidig'
        title = 'Avbrutt'
        return success, message, title

    program_is_running = True
    load_regnskapsprogram_thrd.join()

    ret = regnskapsprogram.run_main_program(*args)
    program_is_running = False
    return ret


def main():
    global load_regnskapsprogram_thrd

    downloads_dir = str(Path.home() / "Downloads")
    documents_dir = str(Path.home() / "Documents")

    default_create_new_account = True
    default_year = str(datetime.date.today().year)
    default_name = ''
    default_new_transactions_file_dir = downloads_dir
    default_account_location = documents_dir
    default_exsisting_form = ''

    tmp_dir = directory_fetcher.get_tmp_dir()

    # Get default exsisting account
    if os.path.isfile(tmp_dir / 'last_account_file.txt'):
        f = open(tmp_dir / 'last_account_file.txt')
        default_exsisting_form = f.read()
        default_account_location = os.path.dirname(default_exsisting_form)
        f.close()
        default_create_new_account = False

    # Get default account name
    if os.path.isfile(tmp_dir / 'last_account_name.txt'):
        f = open(tmp_dir / 'last_account_name.txt')
        default_name = f.read()
        f.close()
    load_regnskapsprogram_thrd = threading.Thread(target=load_regnskapsprogram, args=())
    load_regnskapsprogram_thrd.start()

    user_interface.UI_master.run_GUI(default_create_new_account,
                                     default_year,
                                     default_name,
                                     default_new_transactions_file_dir,
                                     default_account_location,
                                     default_exsisting_form,
                                     directory_fetcher.open_backups_dir,
                                     call_main_program)


if __name__ == "__main__":
    main()
