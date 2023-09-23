import os
import glob

folder_path = './chiDocx'
docx_files = glob.glob(os.path.join(folder_path, '*.docx'))

print(docx_files)