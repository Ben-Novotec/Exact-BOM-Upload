import pandas as pd


def make_exact_code(bom: pd.DataFrame, steel_code: str) -> pd.DataFrame:
    """Determines the exact code and adds it as a column to a dataframe"""
    code = bom['Exact Code'] + bom['ExactCode']
    material = [steel_code if 'Stainless steel' in s else '' for s in bom['Physical material']]
    bom['Code'] = [c + m if c[-2:] != steel_code else c for c, m in zip(code, material)]
    return bom
