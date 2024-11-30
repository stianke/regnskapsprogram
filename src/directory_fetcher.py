
from pathlib import Path
import sys
import os

appdata_folder_name = 'Regnskapsprogram'
template_path = 'template/Regnskap_template.xlsx'

def get_data_dir():
    data_dir = Path.home()

    if sys.platform == "win32":
        data_dir = data_dir / "AppData/Roaming"
    elif sys.platform == "linux":
        data_dir = data_dir / ".local/share"
    elif sys.platform == "darwin":
        data_dir = data_dir / "Library/Application Support"

    data_dir = data_dir / appdata_folder_name
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    return data_dir

def get_tmp_dir():
    return get_data_dir() / 'tmp'

def get_backups_dir():
    return get_data_dir() / 'backups'

def get_template_dir():
    global template_path
    return template_path

def open_backups_dir():
    import subprocess
    directory = get_backups_dir()

    if sys.platform == 'win32':
        subprocess.Popen(['start', directory], shell=True)

    elif sys.platform == 'darwin':
        subprocess.Popen(['open', directory])

    else:
        try:
            subprocess.Popen(['xdg-open', directory])
        except OSError:
            # er, think of something else to try
            # xdg-open *should* be supported by recent Gnome, KDE, Xfce
            pass

def get_last_account_file():
    return get_tmp_dir() / 'last_account_file.txt'
def get_last_account_name_file():
    return get_tmp_dir() / 'last_account_name.txt'
