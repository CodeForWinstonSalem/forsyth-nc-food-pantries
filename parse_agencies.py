"""NC Food Pantry Sheet Parser
Usage:
    parse_agencies.py <filename> [<sheet>...]
    parse_agencies.py (-j | --json) <filename> [<sheet>...]
    parse_agencies.py (-h | --help)

Options:
    -j --json  Export as json
    -h --help  Display this message

This will parse out the following information from an xls file used for the
NC Food Pantry Schedule, all other infomation will be thrown out:
-Agency Name
-Street Address, City, Zipcode
-Latitude and Longitude
-Agency Type
-Monday through Sunday Schedule
-Frequency
-Telephone Number

The output is a separate csv for each sheet.

Specify sheet names to only parse those, otherwise all will be parsed.

The type of agency will default to 'pantry' if not specified in the xls.

This script isn't perfect, some double checking should be done on the output.
Notable errors are when the address is a place name and not an address (see Rockingham sheet)
"""
import sys
import re
import xlrd
from time import sleep
from geopy.geocoders import GoogleV3
from docopt import docopt


class Hours(object):
    def __init__(self, monday=None, tuesday=None, wednesday=None, thursday=None,
                 friday=None, saturday=None, sunday=None):
        self.monday = monday
        self.tuesday = tuesday
        self.wednesday = wednesday
        self.thursday = thursday
        self.friday = friday
        self.saturday = saturday
        self.sunday = sunday


class Agency(object):
    def __init__(self, name, street, city, zipcode, latitude, longitude, hours,
                 frequency, telephone, agency_type='pantry'):
        self.name = name
        self.latitude = latitude
        self.longitude = longitude
        self.street = street
        self.city = city
        self.zipcode = zipcode
        self.hours = hours
        self.frequency = frequency
        self.telephone = telephone
        self.type = agency_type

    def csv(self):
         return '"{name}","{street}","{city}",{zipcode},{latitude},{longitude},{agency_type},{mon},{tue},{wed},{thu},{fri},{sat},{sun},"{frequency}",{telephone}\n'.format(name=self.name, street=self.street, zipcode=self.zipcode, agency_type=self.type, mon=self.hours.monday, tue=self.hours.tuesday, wed=self.hours.wednesday, thu=self.hours.thursday, fri=self.hours.friday, sat=self.hours.saturday, sun=self.hours.sunday, latitude=self.latitude, longitude=self.longitude, frequency=self.frequency, telephone=self.telephone, city=self.city)

    def json(self):
        template = """{{\n"name": "{name}",\n"street": "{street}",\n"city": "{city}",\n"zip": {zipcode},\n"latitude": {latitude},\n"longitude": {longitude},\n"type": "{agency_type}",\n"monday": "{mon}",\n"tuesday": "{tue}",\n"wednesday": "{wed}",\n"thursday": "{thu}",\n"friday": "{fri}",\n"saturday": "{sat}",\n"sunday": "{sun}",\n"frequency": "{frequency}".,\n"phone": "{telephone}"\n}}"""
        return template.format(name=self.name, street=self.street, zipcode=self.zipcode, agency_type=self.type, mon=self.hours.monday, tue=self.hours.tuesday, wed=self.hours.wednesday, thu=self.hours.thursday, fri=self.hours.friday, sat=self.hours.saturday, sun=self.hours.sunday, latitude=self.latitude, longitude=self.longitude, frequency=self.frequency, telephone=self.telephone, city=self.city)


def find_column(sheet, pattern):
    """
    params
        sheet (xlrd.Sheet) - sheet to search
        pattern (re.RegexObject) - pattern to use for search
    returns:
        (col_num, next_row)
        None if column not found
    """
    for ri in range(sheet.nrows):
        for ci in range(sheet.ncols):
            if pattern.search(str(sheet.cell_value(rowx=ri, colx=ci))):
                return (ci, ri+1)
    return None, None


def parse_sheet(sheet):
    """
    params:
        sheet (xlrd.Sheet) - sheet to parse_sheet
    returns:
        [Agency...]
    """
    geolocator = GoogleV3()
    agencies = []
    has_type_col = False
    first_data_row = 0
    # Find Name Column
    name_col, row = find_column(sheet, re.compile('agency.name', flags=re.IGNORECASE))
    if row > first_data_row:
        first_data_row = row
    # Find Address Column
    address_col, row = find_column(sheet, re.compile('address', flags=re.IGNORECASE))
    if address_col and row > first_data_row:
        first_data_row = row
    # Find Zipcode Column
    zipcode_col, row = find_column(sheet, re.compile('zip', flags=re.IGNORECASE))
    if zipcode_col and row > first_data_row:
        first_data_row = row
    # Find Frequency Column
    frequency_col, row = find_column(sheet, re.compile('frequency', flags=re.IGNORECASE))
    if row > first_data_row:
        first_data_row = row
    # Find Telephone Column
    telephone_col, row = find_column(sheet, re.compile('telephone|contact', flags=re.IGNORECASE))
    if row > first_data_row:
        first_data_row = row
    # Find Type Column
    type_col, row = find_column(sheet, re.compile('type', flags=re.IGNORECASE))
    if type_col:
        has_type_col = True
        if row > first_data_row:
            first_data_row = row
    # Find Hours Columns
    mon_col, row = find_column(sheet, re.compile('monday', flags=re.IGNORECASE))
    tue_col, row = find_column(sheet, re.compile('tuesday', flags=re.IGNORECASE))
    wed_col, row = find_column(sheet, re.compile('wednesday', flags=re.IGNORECASE))
    thu_col, row = find_column(sheet, re.compile('thursday', flags=re.IGNORECASE))
    fri_col, row = find_column(sheet, re.compile('friday', flags=re.IGNORECASE))
    sat_col, row = find_column(sheet, re.compile('saturday', flags=re.IGNORECASE))
    sun_col, row = find_column(sheet, re.compile('sunday', flags=re.IGNORECASE))
    if row > first_data_row:
        first_data_row = row
    if not has_type_col:
        re_agency_type = re.compile('onsite|soup.kitchen', flags=re.IGNORECASE)
        agency_type = 'pantry'
    for ri in range(first_data_row, sheet.nrows):
        name = sheet.cell_value(rowx=ri, colx=name_col)
        if name == '' or re.search('agency.name', name, flags=re.IGNORECASE):
            continue
        if not has_type_col and re_agency_type.search(name):
            agency_type = 'onsite'
            continue
        if 'UPDATED' in name:
            name = name[:name.index('UPDATED')]
        name = name.strip()
        frequency = sheet.cell_value(rowx=ri, colx=frequency_col)
        if frequency == '':
            continue
        telephone = str(sheet.cell_value(rowx=ri, colx=telephone_col))
        mon_hours = sheet.cell(rowx=ri, colx=mon_col)
        if mon_hours.ctype == 3:
            mon_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            mon_hours = mon_hours.value
        tue_hours = sheet.cell(rowx=ri, colx=tue_col)
        if tue_hours.ctype == 3:
            tue_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            tue_hours = tue_hours.value
        wed_hours = sheet.cell(rowx=ri, colx=wed_col)
        if wed_hours.ctype == 3:
            wed_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            wed_hours = wed_hours.value
        thu_hours = sheet.cell(rowx=ri, colx=thu_col)
        if thu_hours.ctype == 3:
            thu_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            thu_hours = thu_hours.value
        fri_hours = sheet.cell(rowx=ri, colx=fri_col)
        if fri_hours.ctype == 3:
            fri_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            fri_hours = fri_hours.value
        sat_hours = sheet.cell(rowx=ri, colx=sat_col)
        if sat_hours.ctype == 3:
            sat_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            sat_hours = sat_hours.value
        sun_hours = sheet.cell(rowx=ri, colx=sun_col)
        if sun_hours.ctype == 3:
            sun_hours = xlrd.xldate.xldate_as_datetime(sheet.cell(colx=9, rowx=70).value, sheet.book.datemode).strftime("%H:%M")
        else:
            sun_hours = sun_hours.value
        hours = Hours(mon_hours, tue_hours, wed_hours, thu_hours, fri_hours,
                      sat_hours, sun_hours)
        if has_type_col:
            agency_type = sheet.cell_value(rowx=ri, colx=type_col)
        if zipcode_col:
            zipcode = sheet.cell_value(rowx=ri, colx=zipcode_col)
            if type(zipcode) is float:
                zipcode = str(int(zipcode))
        else:
            zipcode = ''
        if address_col:
            address = sheet.cell_value(rowx=ri, colx=address_col)
            sleep(0.2)
            location = "{}, NC {}".format(address,zipcode)
            location = geolocator.geocode(location)
            if location:
                split_address = location.address.split(',')
                if len(split_address) == 4:
                    street = split_address[0]
                    city = split_address[1]
                    zipcode = split_address[2][-5:]
                latitude = location.latitude
                longitude = location.longitude
            else:
                street = address
                city = address
                latitude = ''
                longitude = ''
        else:
            street = ''
            city = ''
            latitude = ''
            longitude = ''
        agencies.append(Agency(name, street, city, zipcode, latitude, longitude, hours,
                        frequency, telephone, agency_type))
    return agencies


if __name__=="__main__":
    arguments = docopt(__doc__)
    book = xlrd.open_workbook(arguments['<filename>'])
    extension = 'csv'
    if arguments['--json']:
        extension = 'json'
    if len(arguments['<sheet>'])>0:
        names = arguments['<sheet>']
    else:
        names = book.sheet_names()
    for sname in names:
        sheet = book.sheet_by_name(sname)
        print("Reading {} Sheet".format(sheet.name))
        agencies = parse_sheet(sheet)
        print("{} agencies found".format(len(agencies)))
        with open('{}Agencies.{}'.format(sheet.name, extension), 'w+') as f:
            if extension == 'csv':
                f.write('name,street,city,zip,latitude,longitude,type,monday,tuesday,wednesday,thursday,friday,saturday,sunday,frequency,phone\n')
            elif extension == 'json':
                f.write('var pantries =\n  [\n')
            for agency in agencies:
                if extension == 'csv':
                    f.write(agency.csv())
                elif extension == 'json':
                    f.write(agency.json())
                    f.write(',\n')
