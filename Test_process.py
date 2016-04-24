import zipfile

replaceText = {"param" : "Ollie", "param" : "Ollie"}
templateDocx = zipfile.ZipFile("Letter_Template.docx")
newDocx = zipfile.ZipFile("NewDocument.docx", "a")

with open(templateDocx.extract("word/document.xml", "temp/")) as tempXmlFile:
    tempXmlStr = tempXmlFile.read()

for key in replaceText.keys():
    print key
    print replaceText.get(key)
    tempXmlStr = tempXmlStr.replace(str(key), str(replaceText.get(key)))

with open("temp/temp.xml", "w+") as tempXmlFile:
    tempXmlFile.write(tempXmlStr)

for file in templateDocx.filelist:
    if not file.filename == "word/document.xml":
        newDocx.writestr(file.filename, templateDocx.read(file))

newDocx.write("temp/temp.xml", "word/document.xml")

templateDocx.close()
newDocx.close()