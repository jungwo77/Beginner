# ---
# jupyter:
#   jupytext:
#     formats: ipynb,py:light
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.5'
#       jupytext_version: 1.14.1
#   kernelspec:
#     display_name: Python 3 (ipykernel)
#     language: python
#     name: python3
# ---

import pandas as pd
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

tb1 = pd.read_excel("item_base.xlsx", sheet_name=2, header=6, usecols=[1,3])
tb1

tb2 = pd.read_excel("String.xlsx", header=1, usecols=[1,2])
tb2.rename(columns = {'stringkey':'strItemName'},inplace=True)
tb2

tb3 = tb1.merge(tb2, on=["strItemName"], how="left")
tb3

inventors.to_excel('ItemStrng.xlsx', index=False))
tb


