import pandas as pd
from utils import make_exact_code


def main(bom_path: str, steel_code: str):
    bom = pd.read_excel(bom_path, sheet_name=0, header=1, index_col=0)
    bom = bom.fillna('')
    n_items = len(bom.index)

    # add exact code to bom
    bom = make_exact_code(bom, steel_code)

    # makes a list with the supplier codes from Exact which match the supplier in the BOM
    supplier_excel = pd.read_excel(bom_path, sheet_name=1)
    supplier_codes = []
    for s in bom['PartSupplier']:
        if not s:  # if s is an empty string
            supplier_codes.append('')
            continue
        mask = supplier_excel['Naam'].str.contains(s, case=False, regex=False)
        supplier_code = supplier_excel['Code'][mask].values[0]
        supplier_codes.append(supplier_code)

    # TODO eenheid code verschillend van 'pc', zoals 'mt' voor buizen, en deelbaar verschillend van 0
    # makes a new dataframe which can be uploaded to Exact, and writes it as a new sheet to the Excel file
    exact_df = pd.DataFrame({'Code': bom['Code'],
                             'Leveranciercode': supplier_codes,
                             'Artikelgroep: Code': [7] * n_items,
                             'Eenheid: Code': ['pc'] * n_items,
                             'BTW-code: Verkoop': [5] * n_items,
                             'Aankoopeenheid': ['pc'] * n_items,
                             'BTW-code aankoop': [5] * n_items,
                             'Rechtstreekse levering': [0] * n_items,
                             'Aankoopeenheidsfactor': [1] * n_items,
                             'Valuta: Aankoop': ['EUR'] * n_items,
                             'Verkoop': [1] * n_items,
                             'Voorraad': [1] * n_items,
                             'Aankoop': [1] * n_items,
                             'Deelbaar': [0] * n_items,
                             'Ordergestuurd': [0] * n_items,
                             'Webwinkel': [0] * n_items},
                            index=bom.index)
    exact_df = exact_df.drop_duplicates()
    with pd.ExcelWriter(bom_path, mode='a', if_sheet_exists='replace') as writer:
        exact_df.to_excel(writer, sheet_name='Exact artikelen')


if __name__ == '__main__':
    main('Bill of Materials SQRF02_B voor exact.xlsx', 'A4')

