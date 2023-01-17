# imports:
import os
import shutil
import sys
from winreg import HKEY_CURRENT_USER, OpenKey, QueryValueEx, HKEY_CURRENT_USER



def create_folders(directories, DownloadPath):
    for key in directories:
        if key not in os.listdir(DownloadPath):
            os.mkdir(os.path.join(DownloadPath, key))
    if "ALTRO" not in os.listdir(DownloadPath):
        os.mkdir(os.path.join(DownloadPath, "ALTRO"))


def organize_folders(directories, DownloadPath):
    for file in os.listdir(DownloadPath):
        if os.path.isfile(os.path.join(DownloadPath, file)):
            src_path = os.path.join(DownloadPath, file)
            for key in directories:
                extension = directories[key]
                if file.endswith(extension) and file != os.path.basename(sys.executable):
                    dest_path = os.path.join(DownloadPath, key, file)
                    shutil.move(src_path, dest_path)
                    break


def organize_remaining_files(DownloadPath):
    for file in os.listdir(DownloadPath):
        if os.path.isfile(os.path.join(DownloadPath, file)) and file != os.path.basename(sys.executable):
            src_path = os.path.join(DownloadPath, file)
            dest_path = os.path.join(DownloadPath, "ALTRO", file)
            shutil.move(src_path, dest_path)


def organize_remaining_folders(directories, DownloadPath):
    list_dir = os.listdir(DownloadPath)
    organized_folders = []
    for folder in directories:
        organized_folders.append(folder)
    organized_folders = tuple(organized_folders)
    for folder in list_dir:
        if folder not in organized_folders:
            src_path = os.path.join(DownloadPath, folder)
            dest_path = os.path.join(DownloadPath, "CARTELLE", folder)
            try:
                shutil.move(src_path, dest_path)
            except shutil.Error:
                shutil.move(src_path, dest_path + " - copy")
                print("La cartella esiste gia' nella cartella di destinazione."
                      "\nCartella rinominata con'{}'".format(folder + " - copia"))


if __name__ == '__main__':


    with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
        DownloadPath = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
    
    directories = {
        "HTML": (".html5", ".html", ".htm", ".xhtml"),
        "IMMAGINI": (".jpeg", ".jpg", ".tiff", ".gif", ".bmp", ".png", ".bpg",
                   ".svg",
                   ".heif", ".psd"),
        "VIDEO": (".avi", ".flv", ".wmv", ".mov", ".mp4", ".webm", ".vob",
                   ".mng",
                   ".qt", ".mpg", ".mpeg", ".3gp", ".mkv"),
        "DOCUMENTI": (".oxps", ".epub", ".pages", ".docx", ".doc", ".fdf",
                      ".ods",
                      ".odt", ".pwi", ".xsn", ".xps", ".dotx", ".docm", ".dox",
                      ".rvg", ".rtf", ".rtfd", ".wpd", ".xls", ".xlsx", ".ppt",
                      "pptx"),
        "ARCHIVI": (".a", ".ar", ".cpio", ".iso", ".tar", ".gz", ".rz", ".7z",
                     ".dmg", ".rar", ".xar", ".zip"),
        "AUDIO": (".aac", ".aa", ".aac", ".dvf", ".m4a", ".m4b", ".m4p",
                  ".mp3",
                  ".msv", "ogg", "oga", ".raw", ".vox", ".wav", ".wma"),
        "TESTO": (".txt", ".in", ".out"),
        "PDF": ".pdf",
        "EXE": ".exe",
        "ALTRO": "",
        "CARTELLE": ""
    }
    try:
        create_folders(directories, DownloadPath)
        organize_folders(directories, DownloadPath)
        organize_remaining_files(DownloadPath)
        organize_remaining_folders(directories, DownloadPath)
    except shutil.Error:
        print("Errore nello spostamento di alcuni file")
