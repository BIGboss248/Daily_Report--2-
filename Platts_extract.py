import re
import os
import datetime
from datetime import date
import string
import tkinter
from tkinter import filedialog
import pandas as pd
import openpyxl as xl
from openpyxl import formatting, styles, Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.worksheet.table import Table, TableStyleInfo


def find_commodity_price_row(Platts_String: str, commodity_symbol: str, second_pattern='\s+\d+.+'):
    """ Finds the row corosponding to the commodity inside the daily report and returns a re.match"""
    # the pattern is based on the symbol so in the program what the  function will recive is the symbol
    # corosponding to the needed commodity
    pattern = "".join([rf'{commodity_symbol}', second_pattern])
    compiled_pattern = re.compile(pattern)
    matches = re.search(pattern=compiled_pattern, string=Platts_String)
    return matches


def extract_numbers(Platts_String: str, commodity_symbol: str, second_pattern='\s+\d+.+'):
    """gets the file  and commodity symbol and uses the find_commodity_price_row gets the row of the commodity inside the daily report and
    removes the name symbol and spaces of the row and retruns a list of floats containing only the numbers meaning prices and changes"""
    out_match = find_commodity_price_row(
        Platts_String, commodity_symbol, second_pattern)
    # creating a list of english letters to remove the name and symbol
    alphabet = list(string.ascii_letters)
    # extracting the string from re,match and turning it into a list of strings
    in_list = out_match.group().split(" ")
    # removing the space betwean the numbers and name and symbol
    in_list = list(dict.fromkeys(in_list))
    # creating a list of items that should be removed from the original re.match string
    remove_list = []
    for i in in_list:
        if i.isalpha() or i == "" or '%' in i:
            remove_list.append(i)
    for i in in_list:
        for a in alphabet:
            if a in i:
                remove_list.append(i)
                break
    # removing duplicates from remove_list
    remove_list = list(dict.fromkeys(remove_list))
    # removing the items in in_list that are in remove_list
    for i in remove_list:
        in_list.remove(i)
    # in special cases there will be a '.' that will cause error so we check for it and  delete it if its there
    if '.' in in_list:
        in_list.remove('.')
    # turning the list of strings into float
    in_list = [float(x) for x in in_list]
    return in_list


def generate_report(Platts_String: str, commodity: dict, second_pattern='\s+\d+.+'):
    """ taking the generated string from reading platts file and taking a list of required commodity information and Using
    1-find_commodity_price_row 2-extract_numbers it generates a list of lists containing commodity name, price and changes """
    result = []
    for i in list(commodity.keys()):
        numbers = extract_numbers(Platts_String, commodity[i], second_pattern)
        numbers.insert(0, i)
        result.append(numbers)
    return (result)


def final_report(Platts_String: str, commodity: dict, index: int, headers: list, second_pattern='\s+\d+.+'):
    """Using generate_report it takes the output list and extracts the numbers we need plus adding the headers for pandas data farme"""
    report = generate_report(Platts_String, commodity, second_pattern)
    result = []
    for i in report:
        temp = []
        for j in range(index):
            temp.append(i[j])
        result.append(temp)
    output = dict.fromkeys(headers)
    for i in output.keys():
        output[i] = []
    j = 0
    for i in output.keys():
        for k in result:
            output[i].append(k[j])
        j = j+1
    df_output = pd.DataFrame(output)
    df_output.set_index('Commodity', inplace=True)
    return df_output


def get_Volume_Issue_Date(Platts_String: str):
    """gets the file string returns a list first one will be the volume second the issue number and third a date object"""
    pattern = re.compile(r'Volume.+')
    match = re.search(pattern, Platts_String)
    match_string = match.group()
    Volume = int(match_string.split('/')[0].split(' ')[1])
    Issue = int(match_string.split('/')[1].split(' ')[2])
    month = match_string.split('/')[2].split(' ')[1]
    month_dict = {'January': 1, 'February': 2, 'March': 3,
                  'April': 4, 'May': 5, 'June': 6, 'July': 7,
                  'August': 8, 'September': 9, 'October': 10,
                  'November': 11, 'December': 12}
    day = int(match_string.split('/')[2].split(' ')[2].replace(',', ""))
    year = int(match_string.split('/')[2].split(' ')[3])
    date_stamp = datetime.date(year, month_dict[month], day)
    return [Volume, Issue, date_stamp]


def export_to_excel(excel_file_address: str, dataframe_dict: dict):
    """ With a openpyxl.writer writes the dataframes to excel worksheets"""
    excel_writer = pd.ExcelWriter(excel_file_address, engine='openpyxl')
    for dataframe in dataframe_dict:
        dataframe_dict[dataframe].to_excel(excel_writer, sheet_name=dataframe)
    excel_writer.close()


def excel_set_font(excel_file_address: str, font=Font(name='IRNazanin', size=16)):
    """Takes the address of an excel file Adjusts font"""
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for i in range(1, ws.max_row+1):
            for j in range(1, ws.max_column+1):
                selected_cell = ws.cell(row=i, column=j)
                selected_cell.font = font
    wb.save(excel_file_address)


def excel_set_alignment(excel_file_address: str, alignment=Alignment(horizontal='center', vertical='center')):
    """Takes the address of an excel file Adjusts alignment"""
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for i in range(1, ws.max_row+1):
            for j in range(1, ws.max_column+1):
                selected_cell = ws.cell(row=i, column=j)
                selected_cell.alignment = alignment
    wb.save(excel_file_address)


def excel_set_border(excel_file_address: str,
                     border=Border(left=Side(border_style="thin", color='000000'),
                                   right=Side(border_style="thin",
                                              color='000000'),
                                   top=Side(border_style="thin",
                                            color='000000'),
                                   bottom=Side(border_style="thin", color='000000'))):
    """Takes the address of an excel file Adjusts border"""
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for i in range(1, ws.max_row+1):
            for j in range(1, ws.max_column+1):
                selected_cell = ws.cell(row=i, column=j)
                selected_cell.border = border
    wb.save(excel_file_address)


def excel_set_tables(excel_file_address: str, Table_Style=TableStyleInfo(name="TableStyleMedium19")):
    """ Takes the address of an excel file and defines Tables with styles"""
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        # Create Table and set a table style
        tab = Table(displayName=sheet,
                    ref=rf'A1:{string.ascii_uppercase[ws.max_column-1]}{ws.max_row}', tableStyleInfo=Table_Style)
        ws.add_table(tab)
    wb.save(excel_file_address)


def excel_set_number_formats(excel_file_address: str):
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        # set the price and change format to currency
        # we start row from 2 to ignore headers and we set columns 2 and 4 since range() dosent count the last number
        # (price and change columns) to change to currenct format
        for i in range(2, ws.max_row+1):
            for j in range(2, 4):
                selected_cell = ws.cell(row=i, column=j)
                selected_cell.number_format = '"$"#,##0.00_-'
        for i in range(2, ws.max_row+1):
            selected_cell = ws.cell(row=i, column=4)
            if selected_cell.value != None:
                selected_cell.value = selected_cell.value/100.00
                selected_cell.number_format = '0.00%'
    wb.save(excel_file_address)


def excel_set_column_width(excel_file_address: str):
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        # Set column width to auto size
        for i in string.ascii_uppercase:
            max_width = 0
            for row in range(1, ws.max_row+1):
                row_len = len(str(ws[f'{i}{row}'].value))
                if row_len > max_width:
                    max_width = row_len
            ws.column_dimensions[i].width = max_width + 13
    wb.save(excel_file_address)

def excel_set_conditional_formatting(excel_file_address: str):
    red_color = 'ffc7ce'
    red_color_font = '9c0103'
    green_color = 'C6EFCE'
    green_color_font = '006100'
    yellow_color = 'FFEB9C'
    yellow_color_font = '9C5700'
    red_font = styles.Font(color=red_color_font)
    red_fill = styles.PatternFill(start_color=red_color, end_color=red_color, fill_type='solid')
    green_font = styles.Font(color=green_color_font)
    green_fill = styles.PatternFill(start_color=green_color, end_color=green_color, fill_type='solid')
    yellow_font = styles.Font(color=yellow_color_font)
    yellow_fill =styles.PatternFill(start_color=yellow_color, end_color=yellow_color, fill_type='solid')
    wb = xl.load_workbook(excel_file_address)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        range = rf"B2:{string.ascii_uppercase[ws.max_column]}{ws.max_row}"
        # start from row 2 and column 2 to ignore headers and indexes
        ws.conditional_formatting.add(range,formatting.rule.CellIsRule(operator='lessThan', formula=['0'], fill=red_fill, font=red_font))
        ws.conditional_formatting.add(range, formatting.rule.CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))
        ws.conditional_formatting.add(range,formatting.rule.CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill, font=green_font))
        ws.conditional_formatting.add(range, formatting.rule.CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
        ws.conditional_formatting.add(range,formatting.rule.CellIsRule(operator='equal', formula=['0'], fill=yellow_fill, font=yellow_font))
        ws.conditional_formatting.add(range, formatting.rule.CellIsRule(operator='equal', formula=['0'], fill=yellow_fill))
    wb.save(excel_file_address)


def excel_format(excel_file_address: str):
    """ Takes the address of an excel file and Adjusts column width font and format alignment border and table style"""
    excel_set_font(excel_file_address)
    excel_set_alignment(excel_file_address)
    excel_set_border(excel_file_address)
    excel_set_tables(excel_file_address)
    excel_set_number_formats(excel_file_address)
    excel_set_column_width(excel_file_address)
    excel_set_conditional_formatting(excel_file_address)


# declare addresses
Platts_file_full_address = r'G:\Shared drives\Unlimited Drive\Scripts\1-Global_Resourses\Platts-text.txt'

# Open Platts file
Platts_Daily_Report_File = open(
    Platts_file_full_address, 'rt', encoding='utf-8')
Platts_Daily_Report_String = Platts_Daily_Report_File.read()
Platts_Daily_Report_File.close()

# list of commoditys
indexes = {'IODEX 62% Fe CFR North China': 'IODBZ00',
           '65% Fe CFR North China': 'IOPRM00',
           '58% Fe CFR North China': 'IODFE00'}
lump = {'Lump outright': 'IOCLS00'}
pellet = {'Weekly CFR China 65% Fe': 'IOBFC04',
          'Daily CFR China 63% Fe spot fixed price assessment': 'IOCQR04',
          'Atlantic Basin 65% Fe Blast Furnace pellet FOB Brazil': 'SB01095',
          'Direct Reduction 67.5% Fe pellet premium (65% Fe basis)': 'IODBP00'}
ore_brands = {'Pilbara Blend Fines (PBF) CFR Qingdao': 'IOPBQ00',
              'Brazilian Blend Fines (BRBF) CFR Qingdao': 'IOBBA00',
              'Newman High Grade Fines (NHGF) CFR Qingdao': 'IONHA00',
              'Mining Area C Fines (MACF) CFR Qingdao': 'IOMAA00',
              'Jimblebar Fines (JMBF) CFR Qingdao': 'IOJBA00',
              '57% Fe Yandi Fines (YDF) CFR Qingdao': 'IOYFA00'}
Asia_Pacific_coking_coal = {'HCC Peak Downs Region FOB Australia': 'HCCGA00',
                            'HCC Peak Downs Region CFR China': 'HCCGC00',
                            'HCC Peak Downs Region CFR India': 'HCCGI00',
                            'Premium Low Vol FOB Australia': 'PLVHA00',
                            'Premium Low Vol CFR China': 'PLVHC00',
                            'Premium Low Vol CFR India': 'PLVHI00',
                            'Low Vol HCC FOB Australia': 'HCCAU00',
                            'Low Vol HCC CFR China': 'HCCCH00',
                            'Low Vol HCC CFR India': 'HCCIN00',
                            'Low Vol PCI FOB Australia': 'MCLVA00',
                            'Low Vol PCI CFR China': 'MCLVC00',
                            'Low Vol PCI CFR India': 'MCLVI00',
                            'Mid Vol PCI FOB Australia': 'MCLAA00',
                            'Mid Vol PCI CFR China': 'MCLAC00',
                            'Mid Vol PCI CFR India': 'MCVAI00',
                            'Semi Soft FOB Australia': 'MCSSA00',
                            'Semi Soft CFR China': 'MCSSC00',
                            'Semi Soft CFR India': 'MCSSI00'}
Asia_Pacific_brand_relativities_Premium_Low_Vol = {'Peak Downs FOB Australia': 'HCPDA00',
                                                   'Peak Downs CFR China': 'MCBAA00',
                                                   'Saraji FOB Australia': 'HCSAA00',
                                                   'Saraji CFR China': 'MCBAB00',
                                                   'Oaky North FOB Australia': 'HCOKA00',
                                                   'Oaky North CFR China': 'MCBAR00',
                                                   'Illawarra FOB Australia': 'HCIWA00',
                                                   'Illawarra CFR China': 'MCBAH00',
                                                   'Moranbah North FOB Australia': 'HCMOA00',
                                                   'Moranbah North CFR China': 'MCBAG00',
                                                   'Goonyella FOB Australia': 'HCGOA00',
                                                   'Goonyella CFR China': 'MCBAE00',
                                                   'Peak Downs North FOB Australia': 'HCPNA00',
                                                   'Peak Downs North CFR China': 'MCBAJ00',
                                                   'Goonyella C FOB Australia': 'HCGNA00',
                                                   'Goonyella C CFR China': 'MCBAI00',
                                                   'Riverside FOB Australia': 'HCRVA00',
                                                   'Riverside CFR China': 'MCRVR00',
                                                   'GLV FOB Australia': 'HCHCA00',
                                                   'GLV CFR China': 'MCBAF00'}
Asia_Pacific_brand_relativities_Low_Vol_HCC = {'Lake Vermont HCC': 'MCBAN00',
                                               'Carborough Downs': 'MCBAO00',
                                               'Middlemount Coking': 'MCBAP00',
                                               'Poitrel Semi Hard': 'MCBAQ00'}
Dry_bulk_freight_assessments = {'Australia-China-Capesize': 'CDANC00',
                                'Australia-Rotterdam-Capesize': 'CDARN00',
                                'Australia-China-Panamax': 'CDBFA00',
                                'Australia-India-Panamax': 'CDBFAI0',
                                'USEC-India-Panamax': 'CDBUI00',
                                'USEC-Rotterdam-Panamax': 'CDBUR00',
                                'USEC-Brazil-Panamax': 'CDBUB00',
                                'US Mobile-Rotterdam-Panamax': 'CDMAR00'}


df_indexes = pd.DataFrame(final_report(Platts_Daily_Report_String, indexes,
                                       4, ['Commodity', 'Price', 'Change', 'Change %']))

df_lump = pd.DataFrame(final_report(Platts_Daily_Report_String, lump, 3, [
                       'Commodity', 'Price', 'Change']))

df_pellet = pd.DataFrame(final_report(
    Platts_Daily_Report_String, pellet, 3, ['Commodity', 'Price', 'Change']))
# Adding the IODEX Price to Premiums
df_pellet.loc['Weekly CFR China 65% Fe', 'Price'] = df_pellet.loc['Weekly CFR China 65% Fe', 'Price'] + \
    df_indexes.loc['IODEX 62% Fe CFR North China', 'Price']
df_pellet.loc['Direct Reduction 67.5% Fe pellet premium (65% Fe basis)', 'Price'] = df_pellet.loc[
    'Direct Reduction 67.5% Fe pellet premium (65% Fe basis)', 'Price'] + df_indexes.loc['IODEX 62% Fe CFR North China', 'Price']


df_ore_brands = pd.DataFrame(final_report(
    Platts_Daily_Report_String, ore_brands, 3, ['Commodity', 'Price', 'Change']))

df_Asia_Pacific_coking_coal = pd.DataFrame(final_report(
    Platts_Daily_Report_String, Asia_Pacific_coking_coal, 3, ['Commodity', 'Price', 'Change']))

df_Asia_Pacific_brand_relativities_Premium_Low_Vol = pd.DataFrame(final_report(
    Platts_Daily_Report_String, Asia_Pacific_brand_relativities_Premium_Low_Vol, 2, ['Commodity', 'Price']))

df_Asia_Pacific_brand_relativities_Low_Vol_HCC = pd.DataFrame(final_report(
    Platts_Daily_Report_String, Asia_Pacific_brand_relativities_Low_Vol_HCC, 2, ['Commodity', 'Price']))

df_Dry_bulk_freight_assessments = pd.DataFrame(final_report(
    Platts_Daily_Report_String, Dry_bulk_freight_assessments, 3, ['Commodity', 'Price', 'Change'], second_pattern='.+'))

# List of data frames for exporting to excel file
dataframe_dict = {'indexes': df_indexes, 'lump': df_lump, 'pellet': df_pellet, 'ore_brands': df_ore_brands,
                  'coking_coal': df_Asia_Pacific_coking_coal,
                  'Premium_Coal': df_Asia_Pacific_brand_relativities_Premium_Low_Vol,
                  'HCC_Coal': df_Asia_Pacific_brand_relativities_Low_Vol_HCC,
                  'freight_assessments': df_Dry_bulk_freight_assessments}

excel_file_address = r'G:\Shared drives\Unlimited Drive\Global trading\Platts-Daily-Report\Platts-Data-V2.xlsx'
export_to_excel(excel_file_address, dataframe_dict)
excel_format(excel_file_address)
