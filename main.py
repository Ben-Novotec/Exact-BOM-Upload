import artikelen
import stuklijst

if __name__ == '__main__':
    bom_path = 'Bill of Materials SQRF02_B voor exact.xlsx'
    steel_code = 'A4'
    artikelen.main(bom_path, steel_code)
    stuklijst.main(bom_path, steel_code)
