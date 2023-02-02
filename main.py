import artikelen
import stuklijst
from tkinter import filedialog

if __name__ == '__main__':
    bom_path = filedialog.askopenfilename(filetypes=(('Excel', 'xlsx'),))
    steel_code = 'A4'
    artikelen.main(bom_path, steel_code)
    stuklijst.main(bom_path, steel_code)
