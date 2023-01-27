import pandas as pd
from utils import make_exact_code


def main(bom_path: str, steel_code: str):
    bom = pd.read_excel(bom_path, sheet_name=0, header=1)
    bom = bom.rename(columns={'No.': 'Nr'})
    bom = bom.fillna('')
    bom = make_exact_code(bom, steel_code)

    # checks if the component is an assembly
    # checks if the next label has more '.' than the current, or the label is '0'
    more_dots = (bom['Nr'].shift(-1).str.count('\\.') > bom['Nr'].str.count('\\.')) | (bom['Nr'] == '')
    asms = bom[more_dots]
    # dataframe with all assemblies, with the columns needed for Exact
    asms_ex = pd.DataFrame({'label': 'H',
                            'Hoofdcomponent': asms['Code'],
                            'Omschrijving': asms['Component description'],
                            'Seriegrootte': 1,
                            'Besteldoorlooptijd': 1,
                            'Status': 30,
                            'Versienummer': 1})

    # the components on the top level
    top_comp = bom[~bom['Nr'].str.contains('.', regex=False) & bom['Nr']]
    exact_top_asm = asms_ex.iloc[[0]]
    exact_top_comp = top_comp[['Code', 'Quantity']].rename(columns={'Code': 'Onderdeel'})
    exact_top = pd.concat([exact_top_asm, exact_top_comp], axis=1)

    # assemblies with components
    exact_dfs = [exact_top]
    for i in asms.index[1:]:
        asm = asms.loc[[i]]
        asm_nr = asm.at[asm.index[0], 'Nr']
        # subcomponents: everything that starts with the nr, and has one more '.'
        mask = bom['Nr'].str.startswith(asm_nr) & (bom['Nr'].str.count('\\.') == asm_nr.count('.')+1)
        comps = bom[mask]
        exact_comps = comps[['Code', 'Quantity']].rename(columns={'Code': 'Onderdeel'})
        # adds a dataframe with the correct format to the list
        exact_dfs.append(pd.concat([asms_ex.loc[[i]], exact_comps], axis=1))

    exact_df = pd.concat(exact_dfs)

    with pd.ExcelWriter(bom_path, mode='a', if_sheet_exists='replace') as writer:
        exact_df.to_excel(writer, sheet_name='Exact stuklijst', index=False)


if __name__ == '__main__':
    main('Bill of Materials SQRF02_B voor exact.xlsx', 'A4')
