import subprocess
import os
import re

print "Optimising Document Units for CNC Operations"
with open("/home/pi/Desktop/letter24_2.svg", 'w+') as write_file:
    with open('/home/pi/Desktop/letter24.svg', 'r') as svg_file:
        regex1 = 'height=".*"'
        regex2 = 'width=".*"'
        for line in svg_file:
            line = re.sub(regex1, 'height="297mm"', line)
            line = re.sub(regex2, 'width="210mm"', line)
            write_file.write(line)
    svg_file.close()
write_file.close()

print 'Changing Objects to Paths'
subprocess.call(["inkscape", "letter24_2.svg", "--export-text-to-path", "--export-plain-svg", "letter24_3.svg"],
                cwd="/home/pi/Desktop/")

print 'Ungrouping all'
subprocess.call(["inkscape", "-f","letter24_3.svg", "-g", "--verb", "EditSelectAll", "--verb", "SelectionUnGroup",
                 "--verb", "SelectionUnGroup","--verb", "SelectionUnGroup","--verb", "SelectionUnGroup","--verb",
                 "SelectionUnGroup","--verb", "SelectionUnGroup",
                 "--verb", "FileSave", "--verb", "FileQuit"],
                cwd="/home/pi/Desktop/")

print "Removing Last Line"
subprocess.call(["inkscape", "-f",  "letter24_3.svg", "-g",  "--verb", "EditSelectAll", "--verb",
                 "net.wasbo.filter.reopenSingleLineFont.noprefs", "--verb", "FileSave", "--verb", "FileQuit"],
                cwd="/home/pi/Desktop/")


# os.remove('/home/pi/Desktop/letter24.svg')
os.remove('/home/pi/Desktop/letter24_2.svg')

print "Sending file to Axidraw"
subprocess.call(["inkscape", "-f",  "letter24_3.svg", "-g",  "--verb", "command.evilmadscientist.axidraw.rev110b.noprefs"],
                cwd="/home/pi/Desktop/")

print "Finished Finally"

