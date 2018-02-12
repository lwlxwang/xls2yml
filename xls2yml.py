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

#
def write_spliter(file):
    file.write('#-------------------------------------------------\n')

#####################################################
#   Collect parameters as a initial step
#####################################################
# Get in/out files from standard input.
argvs = sys.argv
argc  = len(argvs)

if argc != 2:
    print '  Usage: # python %s [input file]' % argvs[0]
    print '        [input file]  : Excel file'
    quit()
else:
    input_file  = argvs[1]


#####################################################
#   Open in/out file
#####################################################
# Open input Excel file including device information.
try:
    book = xlrd.open_workbook(input_file)
    file_name = re.split('[/.]', input_file)
except:
    print "ERROR : Can't open EXCEL file as input data file!!"

# Open output YAML file.
output_file = "./output/" + file_name[-2] + ".yaml"
try:
    f = open(output_file, "w")
    f.write('# This YAML file has been made by xls2yml.py\n')
    write_spliter(f)
except:
    print "ERROR : Can't write output file!!"


#####################################################
#    Read input file and convert to YAML format
#####################################################
# Initializing file level parameter.
out = {}

# Read each sheets included in input excel file.
for sheet_name in book.sheet_names():           ##### EACH SHEET

    # Read description from Note page
    if sheet_name == "Note":
        yaml_data = book.sheet_by_name(sheet_name)
        f.write("# " + yaml_data.cell(0, 0).value + "\n")
        write_spliter(f)

    # Read a sheet from input excel file.
    else:
        yaml_data = book.sheet_by_name(sheet_name)

        # Initializing sheet level parameter.
        label  = []
        sheet = {}

        # Start to read excel file by each row.
        for row in range(yaml_data.nrows):      ##### EACH ROW
            # Initializing line level parameter.
            line = {}

            # Make label list.
            if row == 0 :
                for col in range(yaml_data.ncols):      ##### EACH COLUM
                    label.append(yaml_data.cell(row,col).value)
                    #print yaml_data.cell(row,col).value

            # Read each line data.
            else:
                for col in range(yaml_data.ncols):      ##### EACH COLUM
                    if label[col] == "name":
                        name = yaml_data.cell(row,col).value
                    else:
                        line[label[col]] = yaml_data.cell(row,col).value

                sheet[name] = line

        # Write device information in YAML data format.
        out[sheet_name] = sheet

#####################################################
#    Write the YAML file
#####################################################
f.write(yaml.dump(out, default_flow_style=False, allow_unicode=True))
f.write('\n')

# Closing output YAML file.
f.close()
