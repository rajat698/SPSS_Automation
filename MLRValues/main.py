import docx
import os
import glob
import MLR_table
import paragraph

folder_path = './docx/linearRegressionDocx'
files = glob.glob(os.path.join(folder_path, '*.docx'))

for i in files:
    read_doc = docx.Document(i)
    write_doc = docx.Document()
    MLR_table.populate_coefficients_table(read_doc, write_doc )
    paragraph.adjustedRsquareLine(read_doc, write_doc)
    write_doc.save(f'{i[23:]}')
