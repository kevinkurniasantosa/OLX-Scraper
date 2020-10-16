import unicodedata
import glob
from pyexcel.cookbook import merge_all_to_a_book

filename_csv = 'OLX Cbcb.csv'
filename_excel = 'OLX Cbcb.xlsx'
merge_all_to_a_book(glob.glob(filename_csv), filename_excel)
