import json
import openpyxl as pyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Font, Border, colors
from datetime import datetime

def add(JSON,Excel):
    month_dict = {
        1: 'January',
        2: 'February',
        3: 'March',
        4: 'April',
        5: 'May ',
        6: 'June ',
        7: 'July',
        8: 'August',
        9: 'September',
        10: 'October',
        11: 'November',
        12: 'December'
    }

    bank_mask = {
        'A':	'ABB',
        'B':	'ABMB',
        'C':	'AGRO',
        'D':    'AMBB',
        'E':	'ARM',
        'F':	'BIMB',
        'G':	'BKRB',
        'H':	'BMMB',
        'I':	'BOC',
        'J':	'BSN',
        'K':	'CIMB',
        'L':	'CITI',
        'M':	'HLB',
        'N':	'HSBC',
        'O':	'KFH',
        'P':	'MBB',
        'Q':	'MBSB',
        'R':	'OCBC',
        'S':	'PBB',
        'T':	'RHB',
        'U':	'SCB',
        'V':	'UOB',
        'W':	'CPAY',
        'X':	'FASP',
        'Y':	'MOB1',
        'Z':	'NETS',
        'A1':	'RMS',
        'A2':	'RSSB'
    }

    loc_dict = {
        "issuer_transaction": {
            'start': {'row': 11, 'col': 6},
            'cmp head': [False],
            'id start': {'row': 11, 'col': 3},
            'id length': 3,
            'order': [
                "contact_volume",
                "contact_value",
                "contactless_volume",
                "contacless_value",
                "sales_volume",
                "sales_value"
            ]
        },
        "acquirer_transaction": {
            'start': {'row': 11, 'col': 6},
            'cmp head': [False],
            'id start': {'row': 11, 'col': 3},
            'id length': 3,
            'order': [
                "contact_volume",
                "contact_value",
                "contactless_volume",
                "contacless_value",
                "sales_volume",
                "sales_value"
            ]
        },
        "denial_codes": {
            'start': {'row': 6, 'col': 5},
            'cmp head': [True, {'row': 5,'col': 5}],
            'id start': {'row': 6, 'col': 2},
            'id length': 3,
            'order': [
                False
            ]
        },
        'rest': [
            ['title','Summary','C4'],
            ['title','issuer_transaction','F4'],
            ['title','acquirer_transaction','F4'],
        ],
        "system_availability":{
            'each':{
                'Service Maintenance': {
                    
                }
            }
        }
    }

    def iter_tool(start, finish, row):
        for i in range(start, finish):
            yield get_column_letter(i) + str(row)

    identifier = ['year', 'month', 'bank']

    JSON = json.load(JSON)
    WB = pyxl.load_workbook(Excel)

    if 'year' not in JSON.keys():
        JSON['year'] = '2021'

    JSON['month'] = month_dict[JSON['month']]

    for key in JSON:
        if key in identifier:
            continue
        key_space = key.replace('_',' ').title()
        if key_space not in WB.sheetnames or key not in loc_dict:
            print(f'{key_space} not in')
            continue

        WS = WB[key_space]
        max_row = WS.max_row

        for row in range(1,max_row):
            adress = get_column_letter(loc_dict[key]['start']['col']-1)+str(row)
            if loc_dict[key]['id length'] == 3 and WS[adress].value in JSON[key]:
                the_cell = True
                i = 0
                for col in WS.iter_cols(min_col=loc_dict[key]['id start']['col'],max_col=loc_dict[key]['id start']['col'] + loc_dict[key]['id length'] - 2, min_row=row,max_row=row):
                    if str(col[0].value) != JSON[identifier[i]]:
                        the_cell = False
                        break
                    i+= 1
                if the_cell and not loc_dict[key]['cmp head'][0]:
                    i = 0
                    for cell in iter_tool(loc_dict[key]['start']['col'], loc_dict[key]['start']['col'] + len(loc_dict[key]['order']),row):
                        WS[cell].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        WS[cell].value = loc_dict[key]['order'][i]
                        i+= 1
                elif the_cell:
                    header = iter_tool(loc_dict[key]['start']['col'], loc_dict[key]['start']['col'] + len(JSON[key][WS[adress].value].keys()), loc_dict[key]['cmp head'][1]['row'])
                    for cell in iter_tool(loc_dict[key]['start']['col'], loc_dict[key]['start']['col'] + len(JSON[key][WS[adress].value].keys()), row):
                        WS[cell].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        WS[cell].value = WS[next(header)].value
            if bank_mask[WS[adress].value] == JSON['bank']:
                WS[adress].value = JSON['bank']
                WS[adress].fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    for target, sheet, coor in loc_dict['rest']:
        if target == 'title':
            sheet = sheet.replace('_', ' ').title()
            WS = WB[sheet]
            WS[coor].value = WS[coor].value.replace('ABC', JSON['bank'])
    WB.save(f'{datetime.now().strftime("%m-%d-%Y_%H-%M-%S")}_{JSON[identifier[2]]}.xlsx')

add(open('req.json'),'excel.xlsx')