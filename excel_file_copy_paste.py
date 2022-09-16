import pandas as pd
import os
import xlwings as xw
from pathlib import Path
import re
from typing import Tuple, Union
import numpy as np

from typing import List, Tuple
from enum import Enum


class CopyOptions(Enum):
    # https://docs.microsoft.com/en-us/office/vba/api/excel.xlpastetype

    xlPasteAll = -4104  # Everything will be pasted.
    xlPasteAllExceptBorders = 7  # Everything except borders will be pasted.
    # Everything will be pasted and conditional formats will be merged.
    xlPasteAllMergingConditionalFormats = 14
    # Everything will be pasted using the source theme.
    xlPasteAllUsingSourceTheme = 13
    xlPasteColumnWidths = 8  # Copied column width is pasted.
    xlPasteComments = -4144  # Comments are pasted.
    xlPasteFormats = -4122  # Copied source format is pasted.
    xlPasteFormulas = -4123  # Formulas are pasted.
    # Formulas and Number formats are pasted.
    xlPasteFormulasAndNumberFormats = 11
    xlPasteValidation = 6  # Validations are pasted.
    xlPasteValues = -4163  # Values are pasted.
    xlPasteValuesAndNumberFormats = 12  # Values and Number formats are pasted.


def range_to_range(
        source_sheet: xw.main.Sheet, source_address: str,
        target_sheet: xw.main.Sheet, target_address: str, options: CopyOptions) -> None:
    ''' source_range의 데이터, 서식을 target_range에 복사 '''
    source_sheet.range(source_address).copy()
    target_sheet.range(target_address).api.PasteSpecial(options.value)


def split_ranges_by_space(
        sheet: xw.main.Sheet, address: str, ref_col: str, name_col: str) -> Tuple[List[xw.main.Range], List[str]]:
    ''' 공백을 이용하여 range를 address 
        빠른 처리를 위해서 excel 접근은 최소화
    '''
    # 주소 분리
    p = re.compile('[\d]+')
    rows = p.findall(address)

    if len(rows) == 2:
        top = rows[0]
        bottom = rows[1]

    p = re.compile('[a-zA-Z]+')
    cols = p.findall(address)

    if len(cols) == 2:
        left = cols[0]
        right = cols[1]

    # 주소를 증가하면서 공백으로 구분되는 구간을 확인
    # 공백 제거(위에서 아래방향으로 채움)
    ref_address = ref_col + top + ":" + ref_col + bottom
    ref_data = sheet.range(ref_address).value
    ref_series = pd.Series(ref_data, name='ref').ffill()

    # 이름 컬럼 처리
    name_address = name_col + top + ":" + name_col + bottom
    name_data = sheet.range(name_address).value
    name_series = pd.Series(name_data, name='name').ffill()

    values = pd.concat([name_series, ref_series], axis=1)

    spans = []
    name_ref = []
    for row in values.drop_duplicates().itertuples(index=False):
        spans.append(
            (values[(values['name'] == row.name) & (values['ref'] == row.ref)].index.min(),
             values[(values['name'] == row.name) & (values['ref'] == row.ref)].index.max()))
        name_ref.append(row.name+'_'+row.ref)

    ranges = [left + str(int(top)+span[0])+":"+right +
              str(int(top)+span[1]) for span in spans]

    return ranges, name_ref


def used_range(sheet: xw.main.Sheet) -> xw.Range:

    used_range_rows = (sheet.api.UsedRange.Row,
                       sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count)
    used_range_cols = (sheet.api.UsedRange.Column,
                       sheet.api.UsedRange.Column + sheet.api.UsedRange.Columns.Count)
    used_range = xw.Range(*zip(used_range_rows, used_range_cols))
    return used_range
    # used_range.select()


def district_heating_file_split(
    path=Path(r'D:\python_dev\excel'),
    filename='집단에너지 전력배출계수 검토_220208.xlsx',
    sheet_name='district',
    introduction='ch8:cz56',
    header='r6:ce7',
    content='r8:ce664'
):
    ''' 집단에너지 배출계수 검토 결과를 각 사업장에 안내할 수 있도록 개별 파일로 작성 '''

    i = 0
    with xw.App() as app:

        try:
            filename_ext = filename[:filename.rfind('.')]
            book = app.books[filename_ext]
            print(f'{book.name} 파일 작업 중')

        except KeyError:
            book = app.books.open(path/filename)
            print(f'{book.name} 로드')

        except FileNotFoundError:
            book = app.books.open(path/filename)
            print(f'{book.name} 로드')

        try:
            source_sheet = book.sheets[sheet_name]
            print(f'sheet : {source_sheet.name}')

        # sheet 명이 잘못된 경우
        except:
            raise ValueError('sheet 이름 확인')

        ranges, values = split_ranges_by_space(
            sheet=source_sheet, address=content, ref_col='v', name_col='s')

        for range, value in zip(ranges, values):
            target_book = app.books.add()
            target_sheet = target_book.sheets.active

            # introduction
            range_to_range(
                source_sheet=source_sheet,
                source_address=introduction,
                target_sheet=target_sheet,
                target_address='b2',
                options=CopyOptions.xlPasteAllUsingSourceTheme)

            # header
            range_to_range(
                source_sheet=source_sheet,
                source_address=header,
                target_sheet=target_sheet,
                target_address='b53',
                options=CopyOptions.xlPasteAllUsingSourceTheme)

            # content
            range_to_range(
                source_sheet=source_sheet,
                source_address=range,
                target_sheet=target_sheet,
                target_address='b55',
                options=CopyOptions.xlPasteAllUsingSourceTheme)

            filename = value + '.xlsx'
            target_book.save(path/filename)
            target_book.close()
