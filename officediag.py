#!/usr/bin/env python
import subprocess
import os
import shutil

# Office application strings
word = "Microsoft\ Word"
excel = "Microsoft\ Excel"
powerpoint = "Microsoft\ PowerPoint"
onenote = "Microsoft\ OneNote"
outlook = "Microsoft\ Outlook"

# MTOD
print("Python Office for Mac Logs gatherer v0.2"
      "The logs will be generated in logs folder below the one you are running this from."
      "\nIMPORTANT. Each application (Word, Excel, PowerPoint, OneNote and Outlook"
      "must have been ran at least once, even if not successfully.")


# creates the log folder
def logfolder():
    print("Checking for logs folder..")
    pwd = os.getcwd()
    logs = pwd + "/logs"
    if not os.path.exists(logs):
        os.makedirs(logs)

logfolder()


# some system variables. for TO-DO exception environ which goes to sys.txt
loggeduser = os.getlogin()
environ = os.environ
strenviron = str(os.environ)


# creates a unique txt file each log
logfileword = open("logs/MSword.txt", "w")
logfileexcel = open("logs/MSexcel.txt", "w")
logfilepowerpoint = open("logs/MSppt.txt", "w")
logfileonenote = open("logs/MSonenote.txt", "w")
logfileoutlook = open("logs/MSoutlook.txt", "w")
logfilegeneral = open("logs/OS.txt", "w")
logfilemsft = open("logs/msft.txt", "w")
logfilesyspro = open("logs/syspro.txt", "w")


# creates sys.txt apart, as it works differently
def logsys():
    print("Gathering environment variables..")
    sysfile = open("logs/sys.txt", 'w')
    sysfile.write(strenviron)
    sysfile.close()


# the logs themselves, syslog grep + app go to their correspondent txt file
def logoffice():
    print("Exporting Office applications logs.. This could take a moment.")
    logword = subprocess.check_call(['syslog | grep ' + word], shell="True", stdout=logfileword)
    logexcel = subprocess.check_call(['syslog | grep ' + excel], shell="True", stdout=logfileexcel)
    logpowerpoint = subprocess.check_call(['syslog | grep ' + powerpoint], shell="True", stdout=logfilepowerpoint)
    logonenote = subprocess.check_call(['syslog | grep ' + onenote], shell="True", stdout=logfileonenote)
    logoutlook = subprocess.check_call(['syslog | grep ' + outlook], shell="True", stdout=logfileoutlook)
    loggeneral = subprocess.check_call(['syslog'], shell="True", stdout=logfilegeneral)
    loggmsft = subprocess.check_call(['system_profiler | grep Microsoft'], shell="True", stdout=logfilemsft)


# the full system profile output
def logsyspro():
    print("Exporting system configuration.. This could take a moment.")
    logsyspro = subprocess.check_call(['system_profiler'], shell="True", stdout=logfilesyspro)


# announces the end of the process, zips the results, and opens the finder showing them
def results():
    pwd = os.getcwd()
    logs = pwd + "/logs"
    shutil.make_archive("OfficeLogs", 'zip', logs)
    print("All the information has been collected, the logs folder will now open,"
          "and its content have been zipped into OfficeLogs.zip.")
    subprocess.call(["open", pwd])


# execute functions
logoffice()
logsys()
logsyspro()
results()


# keep updated at https://github.com/audricd/PythonOfficeforMacLogs
