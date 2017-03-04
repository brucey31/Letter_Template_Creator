__author__ = 'brucepannaman'

import zipfile
import csv
import os
import subprocess
import time
import shutil


template_folder = 'FOCUSGE_TEMPLATES'
customization_list = 'FOCUSGE_CONTACTS.csv'

# Look through and validate you customisation list
Customisation_list = open(customization_list, "rb")
reader = csv.reader(Customisation_list)

folder_name = time.strftime('%Y-%m-%d-%H:%M:%S')
os.mkdir(folder_name)

letter_iterator = 0

for row in reader:
    template_name = template_folder +"/"+ str(row[0])+".docx"
    print "Using Template %s for this letter" % template_name

    templateDocx = zipfile.ZipFile(template_name)
    replacement_dict = []

    with open(templateDocx.extract("word/document.xml", "")) as tempXmlFile:
        tempXmlStr = tempXmlFile.read()
        num_variables = str.count(tempXmlStr, 'param')

    field_iterator = 0
    for field in row:
        if field_iterator > 0:
            replacement_dict.append(field)

        field_iterator += 1

# Validation of the customization list against the variables
    if len(replacement_dict) > num_variables:
        print "Holy shit man, there are not enough parameters in your template for the ones on your customisation sheet"
        continue

    if len(replacement_dict) < num_variables:
        print "Holy shit man, you have too few variables in your customisation sheet"
        continue

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
        newDocx.close()

# Convert to PDF
        print "Converting " + "/%s/letter%s.docx to PDF" % (folder_name, letter_iterator)

        subprocess.Popen(["soffice", "--convert-to", "pdf", "/Users/Bruce/Documents/Customisation_Templates/%s/letter%s.docx" % (folder_name, letter_iterator)], cwd="/Users/Bruce/Documents/Customisation_Templates/%s/" % folder_name)
        time.sleep(7)

# Send it to the printer
#         print "Sending to printer letter with variables %s" % replacement_dict
        # subprocess.check_call(["lp", "-d", "Hannah_s_Printer", "%s/letter%s.pdf" % (folder_name, letter_iterator)])
        # subprocess.call(["launch", "-p", "%s/letter%s.docx" % (folder_name, letter_iterator)])

        letter_iterator = letter_iterator + 1
        time.sleep(3)

shutil.rmtree('word')
os.remove('temp.xml')

deletelist = [f for f in os.listdir(folder_name) if f.endswith(".docx")]
for f in deletelist:
    os.remove("%s/%s" % (folder_name, f))



