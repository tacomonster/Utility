from os.path import isfile, join
from os import listdir
import os
import shutil
from win32com.client import Dispatch
import zipfile
import re
import pickle


"""Nice functions that make life easier"""


def pickle_object(obj, path):
    """Sends objects to pickle file"""
    with open(path, 'wb') as file:
        pickle.dump(obj, file)


def unpickle_object(path):
    """Unpickles a pickle file"""
    with open(path, 'rb') as file:
        return pickle.load(file)


def dir_files(dir):
    """
    :param dir: file path to directory
    :return:list of all the files in a directory
    """
    return [dir + os.sep + f for f in listdir(dir) if isfile(join(dir, f))]


def dir_subdir_files(dir):
    """ Files in current directory and subfolders of directory
    :param dir: fill path to directory
    :return: list of all the files in a directory
    """
    files = []
    for path, subdirs, files in os.walk(dir):
        for name in files:
            files.append(os.path.join(path, name))
    return files


def copy_file(file_path, to_dir):
    """Copies a file from its location and pastes it into a new directory"""
    f_name = file_path.split(os.sep)[-1]
    new_path = to_dir + os.sep + f_name
    shutil.copy(file_path, new_path)


def extract_deal_id(path):
    """Etracts Deal ID from file name"""
    path = path.split(os.sep)[-1]       # Filename
    deal_id = re.search(r'[0-9]{4,10}', path)
    if bool(deal_id):
        return deal_id.group()
    return None


def unzip_file(path_to_file):
    """Unzips file and returns path, if file cannot be unzipped the returns path."""
    try:
        zip_ref = zipfile.ZipFile(path_to_file, 'r')
        extract_to = path_to_file.replace('.zip', '')
        zip_ref.extractall(extract_to)
        zip_ref.close()
        return extract_to
    except PermissionError:
        return path_to_file


def transfer_excel_data(agg_path, out_path, loss_trk_path):
    """ VBA Macro Injector
    Transfering data from aggragate excel sheet over to Loss Tracker Excel sheet by
    injecting an excel macro and then running it in the Loss Tracker sheet
    :returns Excel File with updated data"""
    macro_path = str(os.path.abspath(os.path.dirname(sys.argv[0])) + os.sep + "vbaMacro.txt")
    vba_macro = open(macro_path).read()
    vba_macro = vba_macro.format(loss_trk_path, agg_path, datetime.today().strftime("%m/25/%Y"))
    # Create Excel Object
    # try:
    com_instance = Dispatch("Excel.Application")
    com_instance.Visible = True
    com_instance.DisplayAlerts = False
    # Open Excel Files
    print('[CLICK] Accept Updates')
    com_instance.Workbooks.Open(agg_path)
    objworkbook = com_instance.Workbooks.Open(loss_trk_path)
    # Inject Macro & Run Macro
    xlmodule = objworkbook.VBProject.VBComponents.Add(1)
    xlmodule.CodeModule.AddFromString(vba_macro.strip())
    com_instance.Application.Run('FilePath')
    objworkbook.SaveAs(out_path)
    com_instance.Quit()
