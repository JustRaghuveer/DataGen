import xlrd
from faker import Faker, Factory
import xlwt
from faker.providers import address, bank

AU_DataGen = Faker("en_US")

AU_DataGen = Factory.create()
AU_DataGen.add_provider(address)
AU_DataGen.add_provider(bank)

# Give the location of the file
filelocation = ("C:\\Users\\rhebbar\\Desktop\\Example.xlsm")

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
# wb.row.height_mismatch = True
# wb.row.height = 256*20

# For row 0 and column 0
print(sheet.cell_value(0, 0))
print(sheet.cell_value(1, 0))

NumberOfTimes = int(sheet.cell_value(2, 0))

print(NumberOfTimes)

HeaderNames = sheet.col_values(0)
# HeaderNames = [x for x in HeaderNames if x]
HeaderNames = HeaderNames[2:] # print list starting from 2nd element
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

for rowheight in range(NumberOfTimes):
    ws.row(rowheight + 1).height_mismatch = True
    ws.row(rowheight + 1).height = 20 * 30

for i in range(5, sheet.nrows):
    DataList = sheet.row_values(i)
    print(DataList[1])
    # DataList = [x for x in DataList if x]  # removes empty spaces in list
    # print(DataList[0])
    # print(DataList[0]), DataList[0] has the first column value of list and Datalist[1] is having second caloumn value.
    if 'FullName' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.name(), style=StyleCells)
            print(AU_DataGen.name())
    else:
        pass

    if 'FirstName' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, DataList[0])
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.first_name(), style=StyleCells)
            print(AU_DataGen.first_name())
    else:
        pass

    if 'LastName' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.last_name(), style=StyleCells)
            print(AU_DataGen.last_name())
    else:
        pass

    if 'NumberLength' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.random_number(digits=int(DataList[2]), fix_len=True), style=StyleCells)
            print(AU_DataGen.random_number(digits=int(DataList[2]), fix_len=True))
    else:
        pass

    if 'NumberRange' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.random.randint(int(DataList[2]), int(DataList[3])), style=StyleCells)
            print(AU_DataGen.random.randint(int(DataList[2]), int(DataList[3])))
    else:
        pass

    if 'Email' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.col(i - 5).width = int(30 * 260)
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.email(), style=StyleCells)
            print(AU_DataGen.email())
    else:
        pass

    if 'Safe_Mail' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.col(i - 5).width = int(30 * 260)
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.safe_email(), style=StyleCells)
            print(AU_DataGen.safe_email())
    else:
        pass

    if 'Country' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.col(i - 5).width = int(24 * 260)
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.country(), style=StyleCells)
            print(AU_DataGen.country())
    else:
        pass

    if 'City' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.city(), style=StyleCells)
            print(AU_DataGen.city())
    else:
        pass

    if 'Address' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.col(i - 5).width = int(40 * 260)
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.address(), style=StyleCells)
            print(AU_DataGen.address())
    else:
        pass

    if 'Zipcode' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.zipcode(), style=StyleCells)
            print(AU_DataGen.zipcode())
    else:
        pass

    if 'CustomString' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.bothify(text=DataList[2]), style=StyleCells)
            print(AU_DataGen.bothify(text=DataList[2]))
    else:
        pass

    if 'SecondaryAddress' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.secondary_address(), style=StyleCells)
            print(AU_DataGen.secondary_address())
    else:
        pass

    if 'State' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.state(), style=StyleCells)
            print(AU_DataGen.state())
    else:
        pass

    if 'Street' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.street_name(), style=StyleCells)
            print(AU_DataGen.street_name())
    else:
        pass

    if 'CountryCode' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.country_code(representation="alpha-{}".format(int(DataList[2]))), style=StyleCells)
            print(AU_DataGen.country_code(representation="alpha-{}".format(int(DataList[2]))))
    else:
        pass

    if 'LicencePlate' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.license_plate(), style=StyleCells)
            print(AU_DataGen.license_plate())
    else:
        pass

    if 'BBAN' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.bban(), style=StyleCells)
            print(AU_DataGen.bban())
    else:
        pass

    if 'IBAN' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.iban(), style=StyleCells)
            print(AU_DataGen.iban())
    else:
        pass

    if 'EAN' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.ean(length=int(DataList[2])), style=StyleCells)
            print(AU_DataGen.ean(length=int(DataList[2])))
    else:
        pass

    if 'Colour' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.color_name(), style=StyleCells)
            print(AU_DataGen.color_name())
    else:
        pass

    if 'Hex_Colour' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.hex_color(), style=StyleCells)
            print(AU_DataGen.hex_color())
    else:
        pass

    if 'CreditCard' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.credit_card_number(card_type=None), style=StyleCells)
            print(AU_DataGen.credit_card_number(card_type=None))
    else:
        pass

    if 'CreditCard_Provider' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.credit_card_provider(card_type=None), style=StyleCells)
            print(AU_DataGen.credit_card_provider(card_type=None))
    else:
        pass

    if 'Currency' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.currency_name(), style=StyleCells)
            print(AU_DataGen.currency_name())
    else:
        pass

    if 'CurrencyCode' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.currency_code(), style=StyleCells)
            print(AU_DataGen.currency_code())
    else:
        pass

    if 'CryptoCurrency' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_name(), style=StyleCells)
            print(AU_DataGen.cryptocurrency_name())
    else:
        pass

    if 'CryptoCurrency_Code' in DataList[1]:
        for times in range(NumberOfTimes):
            ws.write(0, i - 5, HeaderNames[i - 2], style=StyleHeader)
            ws.write(times + 1, i - 5, AU_DataGen.cryptocurrency_code(), style=StyleCells)
            print(AU_DataGen.cryptocurrency_code())
    else:
        pass


wb.save("C:\\Users\\rhebbar\\Documents\\ACT\\Excel\\AutoDataGenerator.xls")

