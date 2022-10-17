#!/usr/bin/env python
# coding: utf-8
# Delete raws in excel files that do not corresponded the conditions

import pandas as pd
import string


sourse_file_name = input('Input source file name without format: ')
sourse_file_name = sourse_file_name + '.xlsx'

ad_name = input('Input name of ad campaign: ')

source_file = pd.read_excel(sourse_file_name)

source_file['Наименование брони']  = source_file['Наименование брони'].astype(str)
new_file = source_file[source_file['Наименование брони'].str.contains(ad_name)]

new_ad_name = ''.join([i for i in ad_name if i not in string.punctuation])
new_file_name = 'Erid_' + new_ad_name + '.xlsx'

writer = pd.ExcelWriter(new_file_name)
new_file.to_excel(writer, sheet_name='Erid', index=False)

# auto-adjust columns width
for column in new_file:
    column_width = max(new_file[column].astype(str).map(len).max(), len(column) + 7)
    col_idx = new_file.columns.get_loc(column)
    writer.sheets['Erid'].set_column(col_idx, col_idx, column_width)

writer.save()

