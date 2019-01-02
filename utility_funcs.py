from os.path import isfile, join
from os import listdir
import os
import shutil
from win32com.client import Dispatch
import zipfile
import re
import pickle
import sys

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
    
    
def get_size(obj, seen=None):
    """Recursively finds size of objects"""
    size = sys.getsizeof(obj)
    if seen is None:
        seen = set()
    obj_id = id(obj)
    if obj_id in seen:
        return 0
    seen.add(obj_id)
    if isinstance(obj, dict):
        size += sum([get_size(v, seen) for v in obj.values()])
        size += sum([get_size(k, seen) for k in obj.keys()])
    elif hasattr(obj, '__dict__'):
        size += get_size(obj.__dict__, seen)
    elif hasattr(obj, '__iter__') and not isinstance(obj, (str, bytes, bytearray)):
        size += sum([get_size(i, seen) for i in obj])
    return size
