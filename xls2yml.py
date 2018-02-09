import xlrd
import yaml
import re

#
def unicode_representer(dumper, data):
    return dumper.represent_scalar(u'!unicode', u'%s' % data)
yaml.add_representer(unicode, unicode_representer)
yaml.add_implicit_resolver(u'!unicode', re.compile('^.*$'))


# Open input Excel file including device information
try:
    book = xlrd.open_workbook('sample_data.xlsx')
except:
    print "ERROR : Can't open EXCEL file!!"

# Open output YAML file.
try:
    f = open("sample_out.yml", "w")
except:
    print "ERROR : Can't write output file!!"

# Read a sheet from input excel file.
sheet_1 = book.sheet_by_index(0)

# Initializing parameter.
label = []
f.write('device:\n')


for row in range(sheet_1.nrows):
    #label = ['']*100
    value = {}

    if row == 0:
        for col in range(sheet_1.ncols):
            label.append(sheet_1.cell(row,col).value)
            #label(col) = sheet_1.cell(row,col).value
        #print('----------------------------')

    else:
        for col in range(sheet_1.ncols):
            value[label[col]] = sheet_1.cell(row,col).value

        f.write(yaml.dump(value, default_flow_style=False, allow_unicode=True))
        f.write('\n')

        print value

f.close()
