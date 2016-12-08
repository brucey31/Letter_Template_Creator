import subprocess
import os

with open("/home/pi/Desktop/letter24_2.svg", 'w') as write_file:
    with open('/home/pi/Desktop/letter24.svg', 'r') as svg_file:
        for line in svg_file:
            line2 = line.replace('height="1045.275565px"', 'height="297mm"')
            line3 = line2.replace('width="744.09447px"', 'width="210mm"')
            write_file.write(line3)
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
subprocess.call(["inkscape", "-f",  "letter24_3.svg", "-g",  "--verb", "command.evilmadscientist.axidraw.rev110b.noprefs"
                 ,"--verb", "FileSave", "--verb", "FileQuit"],
                cwd="/home/pi/Desktop/")

print "Finished Finally"

