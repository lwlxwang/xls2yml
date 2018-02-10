# -*- coding: utf-8 -*-
#  Copyright (c) 2018 Yuji Higashi. Code released under the Apache license.
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import xlrd
import yaml
import re
import sys

# Process for unicode character
def unicode_representer(dumper, data):
    return dumper.represent_scalar(u'!unicode', u'%s' % data)
yaml.add_representer(unicode, unicode_representer)
yaml.add_implicit_resolver(u'!unicode', re.compile('^.*$'))

# Get in/out files from standard input.
argvs = sys.argv
argc  = len(argvs)

if argc != 3:
    print '  Usage: # python %s [input file] [output file]' % argvs[0]
    print '        [input file]  : Excel file'
    print '        [output file] : YAML file'
    quit()
else:
    input_file  = argvs[1]
    output_file = argvs[2]


# Open input Excel file including device information
try:
    book = xlrd.open_workbook(input_file)
except:
    print "ERROR : Can't open EXCEL file as input data file!!"

# Open output YAML file.
try:
    f = open(output_file, "w")
except:
    print "ERROR : Can't write output file!!"

# Read a sheet from input excel file.
yaml_data = book.sheet_by_index(0)
yaml_type = book.sheet_by_index(0).name

# Initializing global parameter.
label  = []
device = {}


# Start to read excel file by each row.
for row in range(yaml_data.nrows):
    # Initializing local parameter.
    value = {}

    # Make label list.
    if row == 0:
        for col in range(yaml_data.ncols):
            label.append(yaml_data.cell(row,col).value)

    # Read device data.
    else:
        for col in range(yaml_data.ncols):
            if label[col] == "name":
                device_name = yaml_data.cell(row,col).value
            else:
                value[label[col]] = yaml_data.cell(row,col).value

        device[device_name] = value

# Write device information in YAML data format.
out = {}
out[yaml_type] = device
f.write('#\n#  This YAML file has been made by xls2yml.py\n#\n\n')
f.write(yaml.dump(out, default_flow_style=False, allow_unicode=True))
f.write('\n')

# Closing output YAML file.
f.close()
