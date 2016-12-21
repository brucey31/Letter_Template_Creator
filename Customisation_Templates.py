__author__ = 'brucepannaman'

import zipfile
import csv
import os
import subprocess
import time
import shutil


template_name = 'Focus GE Back.docx'
customization_list = 'Contact_list_python.csv'

# Get your template file and search for the number of parameter to be included in it
templateDocx = zipfile.ZipFile(template_name)

# Look through and validate you customisation list

Customisation_list = open(customization_list, "rb")
reader = csv.reader(Customisation_list)

folder_name = time.strftime('%Y-%m-%d-%H:%M:%S')
os.mkdir(folder_name)

letter_iterator = 0

for row in reader:
    replacement_dict = []

    with open(templateDocx.extract("word/document.xml", "")) as tempXmlFile:
        tempXmlStr = tempXmlFile.read()
        num_variables = str.count(tempXmlStr, 'param')

    for field in row:
        replacement_dict.append(field)

# Validation of the customization list against the variables
    if len(replacement_dict) > num_variables:
        print "Holy shit man, there are not enough paramater in your template for the ones ons your customisation sheet"

    if len(replacement_dict) < num_variables:
        print "Holy shit man, you have too few variables in your customisation sheet"

    else:
        print "All Good in the hood, creating your letter with variables %s" % replacement_dict

# Creates the new letter replacing the headers
        newDocx = zipfile.ZipFile("%s/letter%s.docx" % (folder_name, letter_iterator), "a")

# Replaces each param with the field in the customization dictionary
        for key in replacement_dict:
            tempXmlStr = tempXmlStr.replace("param", str(key), 1)

# Create a tmp file to put the new xml contents in
        with open("temp.xml", "w+") as tempXmlFile:
            tempXmlFile.write(tempXmlStr)

# Write these new xmls to the new file in question
        for file in templateDocx.filelist:
            if not file.filename == "word/document.xml":
                newDocx.writestr(file.filename, templateDocx.read(file))

# Compress the new word file together
        newDocx.write("temp.xml", "word/document.xml")


# Send it to the printer
        print "Sending to printer letter with variables %s" % replacement_dict
        # subprocess.check_call(["lp", "-d", "Hannah_s_Printer", "%s/letter%s.pdf" % (folder_name, letter_iterator)])
        subprocess.call(["launch", "-p", "%s/letter%s.docx" % (folder_name, letter_iterator)])

        letter_iterator = letter_iterator + 1
        newDocx.close()

        time.sleep(3)

shutil.rmtree('word')
os.remove('temp.xml')
