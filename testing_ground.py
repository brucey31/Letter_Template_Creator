import subprocess
import os
import re


def prepare_and_send_to_machine(folder, filo):
    filo_to_use = filo.replace(".svg", "")
    print "Optimising Document Units for CNC Operations for %s/%s" % (folder, filo)
    try:
        with open("%s/%s_2.svg" % (folder, filo_to_use), 'w+') as write_file:
            with open("%s/%s" % (folder, filo), 'r') as svg_file:
                regex1 = 'height=".*"'
                regex2 = 'width=".*"'
                for line in svg_file:
                    line = re.sub(regex1, 'height="297mm"', line)
                    line = re.sub(regex2, 'width="210mm"', line)
                    write_file.write(line)
            svg_file.close()
        write_file.close()
    except Exception as e:
        print "Problems changing the units of the svg"
        print e

    print 'Changing Objects to Paths'
    try:
        subprocess.call(["inkscape", "%s_2.svg" % filo_to_use, "--export-text-to-path", "--export-plain-svg", "%s_3.svg" % filo_to_use],
                        cwd="%s" % folder)
    except Exception as e:
        print "Failed to change objects to path within the svg"
        print e

    print 'Ungrouping all'
    try:
        subprocess.call(["inkscape", "-f", "%s_3.svg" % filo_to_use, "-g", "--verb", "EditSelectAll", "--verb", "SelectionUnGroup",
                         "--verb", "SelectionUnGroup", "--verb", "SelectionUnGroup", "--verb", "SelectionUnGroup", "--verb",
                         "SelectionUnGroup", "--verb", "SelectionUnGroup",
                         "--verb", "FileSave", "--verb", "FileQuit"],
                        cwd="%s" % folder)
    except Exception as e:
        print "Failed to ungroup the text into individual letters"
        print e

    print "Removing Last Line"
    try:
        subprocess.call(["inkscape", "-f", "%s_3.svg" % filo_to_use, "-g", "--verb", "EditSelectAll", "--verb",
                         "net.wasbo.filter.reopenSingleLineFont.noprefs", "--verb", "FileSave", "--verb", "FileQuit"],
                        cwd="%s" % folder)
    except Exception as e:
        print "Failed to remove last line from single line font"
        print e

    # os.remove('/home/pi/Desktop/letter24.svg')
    os.remove('%s/%s_2.svg' % (folder, filo_to_use))

    print "Sending file to Axidraw"
    try:
        subprocess.call(
            ["inkscape", "-f", "%s_3.svg" % filo_to_use, "-g", "--verb", "command.evilmadscientist.axidraw.rev110b.noprefs",
             "--verb", "FileSave", "--verb", "FileQuit"],
            cwd="%s" % folder)
    except Exception as e:
        print "Failed to send to Axidraw"
        print e

    print "Finished Finally"
