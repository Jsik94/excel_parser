import os
from excel_parser.parse import ExcelExtractor


e =ExcelExtractor(__file__)
e.getFileList()
e.extract()
e.saveCSV()