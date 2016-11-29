import subprocess
import os

print 'Changing Objects to Paths'
subprocess.call(["inkscape", "letter24.svg", "--export-text-to-path", "--export-plain-svg", "letter24_2.svg"],
                cwd="/home/pi/Desktop/")

print 'Ungrouping all'
subprocess.call(["inkscape", "-f","letter24_2.svg", "-g", "--verb", "EditSelectAll", "--verb", "SelectionUnGroup",
                 "--verb", "SelectionUnGroup","--verb", "SelectionUnGroup","--verb", "SelectionUnGroup","--verb",
                 "SelectionUnGroup","--verb", "SelectionUnGroup",
                 "--verb", "FileSave", "--verb", "FileQuit"],
                cwd="/home/pi/Desktop/")

print "Removing Last Line"
subprocess.call(["inkscape", "-f",  "letter24_2.svg","-g",  "--verb", "EditSelectAll", "--verb", "net.wasbo.filter.reopenSingleLineFont.noprefs",
                 "--verb", "FileSave", "--verb", "FileQuit"],
                cwd="/home/pi/Desktop/")

print "Finished Finally"

