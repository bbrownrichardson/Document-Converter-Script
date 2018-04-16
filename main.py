import os
import docx2txt
import shutil
from docx import *
import win32com.client


def func(directory, src_directory):
    dir_list = os.listdir(directory)
    copy_files(directory, src_directory)
    for document in dir_list:
        print(document)
        if document[-4:] == ".doc":
            # doc = win32com.client.GetObject(directory)
            # data = doc.Range().Text
            # data_to_txt(document, data)
            pass
        elif document[-5:] == ".docx":
            data = docx2txt.process(document)
            data_to_txt(document, data)
        elif document[-4:] == ".odt":
            # data = docx2txt.process(document)
            # data_to_txt(document, data)
            pass
        else:
            pass

    delete_files(dir_list)
    mv_n_rmv_txt(directory, src_directory)


def copy_files(directory, src_directory):
    src_files = os.listdir(directory)
    for file_name in src_files:
        full_file_name = os.path.join(directory, file_name)
        if os.path.isfile(full_file_name):
            shutil.copy(full_file_name, src_directory)


def delete_files(cpy_file_list):
    for i in cpy_file_list:
        os.remove(i)


def mv_n_rmv_txt(desired_dir, src_directory):
    files = os.listdir(src_directory)
    for i in files:
        if i[-4:] == '.txt':
            files.append(i)
        else:
            pass
    copy_files(src_directory, desired_dir)
    delete_files(files)


def data_to_txt(doc, data):
    filename, file_extension = os.path.splitext(doc)
    f = open(filename + '.txt', "w+")
    # print(data)
    f.write(data.encode('utf-8'))
        # u' '.join(i).encode('utf-8').strip()
        #print(i)
    f.close()


# IMPORTANT: REPLACE the variables get_dir AND src_code_dir with correct
# directories by copy and pasting the desired directories from YOUR
# computer and simply run

# This is the directory that contains the all the docx files to be converted
get_dir = None

# This is the directory that this python file is in
src_code_dir = None

func(get_dir, src_code_dir)
dir_list_temp = os.listdir(get_dir)

# document = opendocx('Heddy 11-97_ 1 corrected.doc')
# fullText=getdocumenttext(document)
# print(fullText)
# a = 'Heddy 11-97_ 1 corrected.doc'
# c = subprocess.call(
#                 ['soffice', '--headless', '--convert-to', 'docx', a])
# print(c)

# doc = win32com.client.GetObject(r'C:\Users\Brianna\Documents\Danielle Montagne Help\BRIANNA')
# data = doc.Range().Text
# print(data)
