__author__ = 'brucepannaman'

import zipfile
import csv
import os
import subprocess
import time
import shutil
import requests
from requests.auth import HTTPBasicAuth
import json



template_name = 'Focus GE Back.docx'
customization_list = 'Contact_list_python.csv'

zamzar_api_key = 'c0d9af9932bd3c368f535b59cd5667abfc9c7930'
endpoint = "https://sandbox.zamzar.com/v1/jobs"

# Get your template file and search for the number of parameter to be included in it
templateDocx = zipfile.ZipFile(template_name)

# Look through and validate you customisation list

Customisation_list = open(customization_list, "rb")
reader = csv.reader(Customisation_list)

folder_name = time.strftime('%Y-%m-%d-%H:%M:%S')
os.mkdir(folder_name)

letter_iterator = 1

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

# Convert to PDF
#         shutil.move("%s/letter%s.docx" % (folder_name, letter_iterator), "/Applications/LibreOffice.app/Contents/MacOS/letter%.docx" % letter_iterator)
#         os.chdir("/Applications/LibreOffice.app/Contents/MacOS")
#         os.rename('/Applications/LibreOffice.app/Contents/MacOS/letter%socx' % letter_iterator, "/Applications/LibreOffice.app/Contents/MacOS/letter%s.docx" % letter_iterator)
#         subprocess.call(["./soffice", "--convert-to", "pdf", "letter%s.docx" % letter_iterator])
#         subprocess.call(["rm", "%s/letter%s.docx" % (folder_name, letter_iterator)])
#         shutil.move("letter%s.pdf" % letter_iterator, "~/Documents/Customisation_Templates/%s/letter%s.pdf" % (folder_name, letter_iterator))
#         subprocess.call(["cd", "/Documents/Customisation_Templates"])

# Convert to PDF2

        # Submit Job
        source_file = "%s/letter%s.docx" % (folder_name, letter_iterator)
        target_format = "pdf"

        file_content = {'source_file': open(source_file, 'rb')}
        data_content = {'target_format': target_format}
        res = requests.post(endpoint, data=data_content, files=file_content, auth=HTTPBasicAuth(zamzar_api_key, ''))
        print res.json()
        file_id = res.json()['id']
        print file_id
        # job_no = response['id'][0]
        # print job_no

        # Check if the job is done
        endpoint = "https://sandbox.zamzar.com/v1/jobs/{}".format(file_id)
        response = requests.get(endpoint, auth=HTTPBasicAuth(zamzar_api_key, ''))
        status = response.json()['status']

        cannot get job_id working get this file and reset it!!!!!

        file_id = response.json()['target_files']['id']
        print "Checking Status of Conversion of %s" % file_id
        print status

        while status != 'successful':
            print "File Not ready yet \nWaiting 5 Secs"
            time.sleep(5)
            response = requests.get(endpoint, auth=HTTPBasicAuth(zamzar_api_key, ''))
            status = response.json()['status']
            print status


        # Download the Finished file
        endpoint = "https://sandbox.zamzar.com/v1/files/{}/content".format(file_id)
        response = requests.get(endpoint, stream=True, auth=HTTPBasicAuth(zamzar_api_key, ''))

        try:
            with open("%s/letter%s.pdf" % (folder_name, letter_iterator), 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        f.flush()

                print "File downloaded"

        except IOError:
            print "Error"



# Send it to the printer Word
#         print "Sending to printer letter with variables %s" % replacement_dict
#         # subprocess.check_call(["lp", "-d", "Hannah_s_Printer", "%s/letter%s.pdf" % (folder_name, letter_iterator)])
#         subprocess.call(["launch", "-p", "%s/letter%s.docx" % (folder_name, letter_iterator)])

        letter_iterator = letter_iterator + 1
        newDocx.close()

        time.sleep(3)

shutil.rmtree('word')
os.remove('temp.xml')
