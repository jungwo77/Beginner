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

# 1테이블 정규화
tb1 = pd.read_excel("item_base.xlsx", sheet_name=2, header=6, usecols=[1,3])
tb1

# 2테이블 정규화 / 머지할 칼럼명이 다르기 때문에 칼럼명 변경
tb2 = pd.read_excel("String.xlsx", header=1, usecols=[1,2])
tb2.rename(columns = {'stringkey':'strItemName'},inplace=True)
tb2

# 병합 3테이블 신규 생성 strItemName 필드 기준으로 합병
tb3 = tb1.merge(tb2, on=["strItemName"], how="left")
tb3

#신규 파일 생성
inventors.to_excel('ItemStrng.xlsx', index=False))


