from config import UPLOAD_FOLDER, DOWNLOAD_FOLDER
import sys
import os
sys.path.append(os.path.abspath(os.path.join('..', 'config')))


def delete_folder_content(folder):
    for _, _, files in os.walk(folder):
        if len(files) > 0:
            for file in files:
                os.remove(folder + "//" + file)


def delete_upload_download_folder_data():
    """DELETE CONTENT IN UPLOAD AND DOWNLOAD FOLDER"""
    try:
        delete_folder_content(UPLOAD_FOLDER)
        delete_folder_content(DOWNLOAD_FOLDER)
    except:
        pass
