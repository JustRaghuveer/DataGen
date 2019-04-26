import pandas as pd
from faker.providers import address, bank
from flask import Flask, render_template, request, send_file, flash  # importing flask module
from faker import Faker, Factory
import xlrd
import xlwt
import os
import csv
import sys
from pathlib import Path

global errormsg
errormsg = []

# initializing a variable of Flask
app = Flask(__name__)

AU_DataGen = Faker("en_US")

AU_DataGen = Factory.create()
AU_DataGen.add_provider(address)
AU_DataGen.add_provider(bank)

Desktop = os.path.expanduser("~\\Desktop")
Dir = '{}\\DataGen_Templates'.format(Desktop)
os.makedirs(Dir, exist_ok=True)

FOLDER = 'C:\\Users\\rhebbar\\Desktop\\DataGen_Templates\\'

def DataGen(filenames):
    # Give the location of the file
    filelocation = ("{}\\{}".format(Dir, filenames))
    # To open Workbook in read mode
    wb = xlrd.open_workbook(filelocation)
    sheet = wb.sheet_by_index(0)

    # open workbook to write data
    wb = xlwt.Workbook()
    ws = wb.add_sheet("DataGenerator", cell_overwrite_ok=True)
    style_header = "font: bold on, color black; borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color teal; align: horiz center, wrap yes,vert centre;"
    StyleHeader = xlwt.easyxf(style_header)
    style_cells = "borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color white; align: horiz center, wrap yes,vert centre;"
    StyleCells = xlwt.easyxf(style_cells)

    StyleDateCells = xlwt.easyxf(style_cells, num_format_str='YYYY-MM-DD')
    # wb.row.height_mismatch = True
    # wb.row.height = 256*20

    # For row 0 and column 0
    print(sheet.cell_value(0, 0))
    print(sheet.cell_value(1, 0))

    # NumberOfTimes = int(sheet.cell_value(2, 0))

    NumberOfTimes = request.form.get('NumberOfTimes')

    # OTB = request.form.get('OTB')
    #
    # Custom = request.form.get('Custom')

    OTB = request.form.get('OTB')

    print(OTB)

    if OTB == '1':
        print("Hurray")

    print(NumberOfTimes)
    print(OTB)

    HeaderNames = sheet.col_values(0)
    # HeaderNames = [x for x in HeaderNames if x]
    HeaderNames = HeaderNames[2:]  # print list starting from 2nd element
    print(HeaderNames)
    print(HeaderNames[1])

    print(sheet.nrows)  # print number of rows in excel that have data


    # list value of 4th row in excel
    DataList = sheet.row_values(4)
    DataList = [x for x in DataList if x]
    print(DataList)

    for x in range(sheet.nrows):
        ws.col(x).width = int(20 * 260)
        ws.row(0).height_mismatch = True
        ws.row(0).height = 20 * 22

    for rowheight in range(int(NumberOfTimes)):
        ws.row(rowheight + 1).height_mismatch = True
        ws.row(rowheight + 1).height = 20 * 30

    for i in range(5, sheet.nrows):
        DataList = sheet.row_values(i)
        print(DataList[1])
        # DataList = [x for x in DataList if x]  # removes empty spaces in list
        # print(DataList[0])
        # print(DataList[0]), DataList[0] has the first column value of list and Datalist[1] is having second caloumn value.

        if 'FullName' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.name(), style=StyleCells)
                    print(AU_DataGen.name())
            except Exception as e:
                print(ws.write(times + 1, i - 5, "Used the keyword wrongly in template", style=StyleCells))
        else:
            pass

        if 'FirstName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                # ws.write(0, i - 5, DataList[0])
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.first_name(), style=StyleCells)
                print(AU_DataGen.first_name())
        else:
            pass

        if 'LastName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.last_name(), style=StyleCells)
                print(AU_DataGen.last_name())
        else:
            pass

        if 'NumberLength' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.random_number(digits=int(DataList[2]), fix_len=DataList[2]),
                             style=StyleCells)
                    print(AU_DataGen.random_number(digits=int(DataList[2]), fix_len=DataList[2]))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> length(int), input 2 -> fix_len(boolean)", style=StyleCells)
        else:
            pass

        if 'NumberRange' in DataList[1]:
            try:
                if None in (DataList[2], DataList[3]):
                    pass
                else:
                    pass
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.random.randint(int(DataList[2]), int(DataList[3])),
                             style=StyleCells)
                    print(AU_DataGen.random.randint(int(DataList[2]), int(DataList[3])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> from(int), input 2 -> to(int)", style=StyleCells)
        else:
            pass

        if 'Email' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(30 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.email(), style=StyleCells)
                print(AU_DataGen.email())
        else:
            pass

        if 'Safe_Mail' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(30 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.safe_email(), style=StyleCells)
                print(AU_DataGen.safe_email())
        else:
            pass

        if 'Country' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(24 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.country(), style=StyleCells)
                print(AU_DataGen.country())
        else:
            pass

        if 'City' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.city(), style=StyleCells)
                print(AU_DataGen.city())
        else:
            pass

        if 'Address' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.col(i - 5).width = int(40 * 260)
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.address(), style=StyleCells)
                print(AU_DataGen.address())
        else:
            pass

        if 'Zipcode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.zipcode(), style=StyleCells)
                print(AU_DataGen.zipcode())
        else:
            pass

        if 'CustomString' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.bothify(text=DataList[2]), style=StyleCells)
                    print(AU_DataGen.bothify(text=DataList[2]))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> # or ?", style=StyleCells)
        else:
            pass

        if 'SecondaryAddress' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.secondary_address(), style=StyleCells)
                print(AU_DataGen.secondary_address())
        else:
            pass

        if 'State' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.state(), style=StyleCells)
                print(AU_DataGen.state())
        else:
            pass

        if 'StreetAddress' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.street_address(), style=StyleCells)
                print(AU_DataGen.street_address())
        else:
            pass

        if 'CountryCode' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.country_code(representation="alpha-{}".format(int(DataList[2]))),
                             style=StyleCells)
                    print(AU_DataGen.country_code(representation="alpha-{}".format(int(DataList[2]))))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> 2 or 3", style=StyleCells)
        else:
            pass

        if 'LicencePlate' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.license_plate(), style=StyleCells)
                print(AU_DataGen.license_plate())
        else:
            pass

        if 'BBAN' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.bban(), style=StyleCells)
                print(AU_DataGen.bban())
        else:
            pass

        if 'IBAN' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.iban(), style=StyleCells)
                print(AU_DataGen.iban())
        else:
            pass

        if 'EAN' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.ean(length=int(DataList[2])), style=StyleCells)
                    print(AU_DataGen.ean(length=int(DataList[2])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> length(int)", style=StyleCells)
        else:
            pass

        if 'Colour' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.color_name(), style=StyleCells)
                print(AU_DataGen.color_name())
        else:
            pass

        if 'Hex_Colour' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.hex_color(), style=StyleCells)
                print(AU_DataGen.hex_color())
        else:
            pass

        if 'CreditCard' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.credit_card_number(card_type=None), style=StyleCells)
                print(AU_DataGen.credit_card_number(card_type=None))
        else:
            pass

        if 'CreditCard_Provider' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.credit_card_provider(card_type=None), style=StyleCells)
                print(AU_DataGen.credit_card_provider(card_type=None))
        else:
            pass

        if 'Currency' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.currency_name(), style=StyleCells)
                print(AU_DataGen.currency_name())
        else:
            pass

        if 'CurrencyCode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.currency_code(), style=StyleCells)
                print(AU_DataGen.currency_code())
        else:
            pass

        if 'CryptoCurrency' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_name(), style=StyleCells)
                print(AU_DataGen.cryptocurrency_name())
        else:
            pass

        if 'CryptoCurrency_Code' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_code(), style=StyleCells)
                print(AU_DataGen.cryptocurrency_code())
        else:
            pass

        if 'Date' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.date(pattern=DataList[2], end_datetime=None), style=StyleCells)
                    print(AU_DataGen.date(pattern=DataList[2], end_datetime=None))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> pattern(Date)", style=StyleCells)
        else:
            pass

        if 'Time' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.time(pattern=DataList[2], end_datetime=None), style=StyleCells)
                    print(AU_DataGen.time(pattern=DataList[2], end_datetime=None))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> pattern(Time)", style=StyleCells)
        else:
            pass

        if 'TimeZone' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.timezone(), style=StyleCells)
                print(AU_DataGen.timezone())
        else:
            pass

        if 'FileName' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.file_name(category=None, extension=DataList[2]), style=StyleCells)
                    print(AU_DataGen.file_name(category=None, extension=DataList[2]))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> Extension(String)", style=StyleCells)
        else:
            pass

        if 'FileExtension' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.file_extension(category=None), style=StyleCells)
                print(AU_DataGen.file_extension(category=None))
        else:
            pass

        if 'FilePath' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.file_path(depth=int(DataList[2]), category=None, extension=DataList[3]), style=StyleCells)
                    print(AU_DataGen.file_path(depth=int(DataList[2]), category=None, extension=DataList[3]))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> Depth(int), input 2 ->Extention(Int)", style=StyleCells)
        else:
            pass

        if 'UpperLetter' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.random_uppercase_letter(), style=StyleCells)
                print(AU_DataGen.random_uppercase_letter())
        else:
            pass

        if 'CustomLetter' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.random_element(elements=list(DataList[2])), style=StyleCells)
                    print(AU_DataGen.random_element(elements=list(DataList[2])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> Depth(int)", style=StyleCells)
        else:
            pass

        if 'CustomWord' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    custword = []
                    custword.append(str(DataList[2]))
                    ws.write(times + 1, i - 5, AU_DataGen.word(ext_word_list=custword), style=StyleCells)
                    print(AU_DataGen.word(ext_word_list=custword))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> list(string)", style=StyleCells)

        if 'FormatedString' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.password(length=int(DataList[2]), special_chars=DataList[3], digits=DataList[4], upper_case=DataList[5], lower_case=DataList[6]), style=StyleCells)
                    print(AU_DataGen.password(length=20, special_chars=False, digits=False, upper_case=True, lower_case=True))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> length(int)", style=StyleCells)
        else:
            pass

        if 'Latitude' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.latitude(), style=StyleCells)
                print(AU_DataGen.latitude())
        else:
            pass

        if 'Longitude' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.longitude(), style=StyleCells)
                print(AU_DataGen.longitude())
        else:
            pass

        if 'StateCode' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.state_abbr(include_territories=True), style=StyleCells)
                print(AU_DataGen.state_abbr(include_territories=True))
        else:
            pass

        if 'Boolean' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.boolean(chance_of_getting_true=50), style=StyleCells)
                print(AU_DataGen.boolean(chance_of_getting_true=50))
        else:
            pass

        if 'DecimalNumber' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.pydecimal(left_digits=int(DataList[2]), right_digits=int(DataList[3]), positive=int(DataList[4])), style=StyleCells)
                    print(AU_DataGen.pydecimal(left_digits=int(DataList[2]), right_digits=int(DataList[3]), positive=int(DataList[4])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> length(int), input 2 -> right_digits(int)", style=StyleCells)
        else:
            pass

        if 'DOB' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.date_of_birth(tzinfo=None, minimum_age=int(DataList[2]), maximum_age=int(DataList[3])), style=StyleDateCells)
                    print(AU_DataGen.date_of_birth(tzinfo=None, minimum_age=int(DataList[2]), maximum_age=int(DataList[3])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> min_age(int), input 2 -> max_age(int)", style=StyleCells)
        else:
            pass

        if 'FutureDate' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.future_date(end_date="+{}".format(DataList[2]), tzinfo=None), style=StyleDateCells)
                    print(AU_DataGen.future_date(end_date="+{}".format(DataList[2]), tzinfo=None))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> end_date(String), eg. -> 20d, 3m, 2y", style=StyleCells)

        else:
            pass

        if 'StoreName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.company(), style=StyleDateCells)
                print(AU_DataGen.company())
        else:
            pass

        if 'StringRange' in DataList[1]:
            try:
                for times in range(int(NumberOfTimes)):
                    ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                    ws.write(times + 1, i - 5, AU_DataGen.pystr(min_chars=int(DataList[2]), max_chars=int(DataList[3])), style=StyleDateCells)
                    print(AU_DataGen.pystr(min_chars=int(DataList[2]), max_chars=int(DataList[3])))
            except Exception as e:
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(1, i - 5, "Used the keyword wrongly in template", style=StyleCells)
                ws.write(2, i - 5, "Plz follow keyword instructions: \n input 1 -> min_chars(int), input 2 ->max_chars(int)", style=StyleCells)
        else:
            pass

        if 'StreetName' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.street_name(), style=StyleDateCells)
                print(AU_DataGen.street_name())
        else:
            pass

        if 'BuildingNumber' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.building_number(), style=StyleDateCells)
                print(AU_DataGen.building_number())
        else:
            pass

        if 'AL_AddressLine1' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, "{} {}".format(AU_DataGen.building_number(), AU_DataGen.street_name()), style=StyleDateCells)
                print("{} {}".format(AU_DataGen.building_number(), AU_DataGen.street_name()))
        else:
            pass

        if 'AL_AddressLine3' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, "{} {}".format(AU_DataGen.city(), AU_DataGen.zipcode()), style=StyleDateCells)
                print("{} {}".format(AU_DataGen.city(), AU_DataGen.zipcode()))
        else:
            pass

        if 'PhoneNumber' in DataList[1]:
            for times in range(int(NumberOfTimes)):
                ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
                ws.write(times + 1, i - 5, AU_DataGen.phone_number(), style=StyleDateCells)
                print(AU_DataGen.phone_number())
        else:
            pass

    wb.save("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls")





@app.route('/')
def index():
    return render_template('DGenTest1.html')


def csv_from_excel():
    # wb = xlrd.open_workbook("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls")
    # sh = wb.sheet_by_index(0)
    # your_csv_file = open("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.csv", 'w')
    # wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    #
    # for rownum in range(sh.nrows):
    #     wr.writerow(sh.row_values(rownum))
    #
    # your_csv_file.close()
    pd.read_excel("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls", sheetname='DataGenerator').to_csv("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator1.csv", index=False)

@app.route('/', methods=['POST'])
def DataGenrator():
    if request.form['action'] == 'Generate':
        if request.method == 'POST':
            if request.form.get('OTB') == '1':
                if request.form.get('Templates') == "AL_Template":
                    listfiles = "Example2.xlsm"   # exact path is in desktop/datagen_templates/Example2.xlsm
                    DataGen(listfiles)
                    if request.form.get('selectoutputtype') == "CSV":
                        csv_from_excel()
                        print("Data Generated Successfully")
                        return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator1.csv',
                                     # '{}'.format(Dir)
                                     mimetype='text/xls',
                                     attachment_filename='DataGenByRaghu.csv',
                                     as_attachment=True)

                    elif request.form.get('selectoutputtype') == "EXCEL":
                        print("Data Generated Successfully")
                        return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls',
                                     # '{}'.format(Dir)
                                     mimetype='text/xls',
                                     attachment_filename='DataGenByRaghu.xls',
                                     as_attachment=True)

                    else:
                        pass

                if request.form.get('Templates') == "AL_StoreImport":
                    listfiles = "AL_StoreImport.xlsm"   # exact path is in desktop/datagen_templates/AL_StoreImport.xlsm
                    DataGen(listfiles)
                    if request.form.get('selectoutputtype') == "CSV":
                        csv_from_excel()
                        print("Data Generated Successfully")
                        return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator1.csv',
                                     # '{}'.format(Dir)
                                     mimetype='text/xls',
                                     attachment_filename='DataGenByRaghu.csv',
                                     as_attachment=True)

                    elif request.form.get('selectoutputtype') == "EXCEL":
                        return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls',
                                     # '{}'.format(Dir)
                                     mimetype='text/xls',
                                     attachment_filename='DataGenByRaghu.xls',
                                     as_attachment=True)
                    else:
                        pass

            else:
                listfiles = request.files['file']
                filenames = listfiles.filename
                xyz = os.path.dirname(os.path.abspath(filenames))
                print(xyz)
                print(filenames)
                print("Hello")
                DataGen(filenames)
                if request.form.get('selectoutputtype') == "CSV":
                    csv_from_excel()
                    print("Data Generated Successfully")
                    return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator1.csv',  # '{}'.format(Dir)
                                 mimetype='text/xls',
                                 attachment_filename='DataGenByRaghu.csv',
                                 as_attachment=True)
                elif request.form.get('selectoutputtype') == "EXCEL":
                    return send_file('C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\FlaskAutoDataGenerator.xls',
                                     # '{}'.format(Dir)
                                     mimetype='text/xls',
                                     attachment_filename='DataGenByRaghu.xls',
                                     as_attachment=True)
                else:
                    pass





    if request.form['action'] == 'Download Template':
        if request.method == 'POST':
            combobox1 = request.form.getlist('Templates')
            print(combobox1)
            i=0
            for combobox1[i] in combobox1:
                if "AL_Template" in combobox1[i]:
                    return send_file('C:\\Users\\rhebbar\\Desktop\\DataGen_Templates\\Example2.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='AL_Template.xlsm',
                                     as_attachment=True)
                if "Default_Template" in combobox1[i]:
                    return send_file('C:\\Users\\rhebbar\\Desktop\\DataGen_Templates\\Template.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='DataGen_Template.xlsm',
                                     as_attachment=True)
                if "AL_StoreImport" in combobox1[i]:
                    return send_file('C:\\Users\\rhebbar\\Desktop\\DataGen_Templates\\AL_StoreImport.xlsm',
                                     mimetype='text/xlsm',
                                     attachment_filename='AL_StoreImport_Template.xlsm',
                                     as_attachment=True)

    if request.form['action'] == 'User Guide':
        if request.method == 'POST':
            return send_file('C:\\Users\\rhebbar\\Desktop\\DataGen_Templates\\Document\\DataGen Usage Guide.pdf',
                             mimetype='text/xlsm',
                             attachment_filename='DataGen User Guide.pdf',
                             as_attachment=True)


@app.route('/errpage')
def erpage():
    return render_template('New500.html')

if __name__ == "__main__":
    app.run(debug=True)