__author__ = 'brucepannaman'

import zipfile
import csv
import os
import datetime
import time
import shutil

template_name = 'Letter_Template.docx'
customization_list = 'Contact_List.csv'

# Get your template file and search for the number of parameter to be included in it
templateDocx = zipfile.ZipFile(template_name)

# Look through and validate you customisation list

Customisation_list = open(customization_list, "rb")
reader = csv.reader(Customisation_list)

folder_name = time.strftime('%Y-%m-%d %H:%M:%S')
os.mkdir(folder_name)

letter_iterator = 0

for row in reader:
    replacement_dict = {}
    field_iterator = 1

    with open(templateDocx.extract("word/document.xml", "")) as tempXmlFile:
        tempXmlStr = tempXmlFile.read()
        num_variables = str.count(tempXmlStr, '{parameter')



    for field in row:
        replacement_dict["param"] = field
        field_iterator = field_iterator + 1

    if len(replacement_dict) > num_variables:
        print "Holy shit man, you have too many variables in your customisation sheet"

    if len(replacement_dict) < num_variables:
        print "Holy shit man, you have too few variables in your customisation sheet"

    else:
        print "All Good in the hood, creating your letter"

        # Creates the new letter replacing the headers

        # Open a new word document

        newDocx = zipfile.ZipFile("%s/letter%s.docx" % (folder_name, letter_iterator), "a")

        for key in replacement_dict.keys():
            print key
            print replacement_dict.get(key)
            tempXmlStr = tempXmlStr.replace(str(key), str(replacement_dict.get(key)))

        with open("temp.xml", "w+") as tempXmlFile:
            tempXmlFile.write(tempXmlStr)

        for file in templateDocx.filelist:
            if not file.filename == "word/document.xml":
                newDocx.writestr(file.filename, templateDocx.read(file))

        newDocx.write("temp.xml", "word/document.xml")

        letter_iterator = letter_iterator + 1



        newDocx.close()

shutil.rmtree('word')
os.remove('temp.xml')
