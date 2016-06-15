import subprocess
import os
#logs.py logs variables and functions
word = "Microsoft\ Word"
excel = "Microsoft\ Excel"
powerpoint = "Microsoft\ PowerPoint"
onenote = "Microsoft\ OneNote"
outlook = "Microsoft\ Outlook"

def logfolder():
    pwd = os.getcwd()
    logs = pwd + "/logs"
    if not os.path.exists(logs):
        os.makedirs(logs)

logfolder()

logfileword = open("logs/MSword.txt", "w")
logfileexcel = open("logs/MSexcel.txt", "w")
logfilepowerpoint = open("logs/MSppt.txt", "w")
logfileonenote = open("logs/MSonenote.txt", "w")
logfileoutlook = open("logs/MSoutlook.txt", "w")
logfilegeneral = open("logs/OS.txt", "w")

print("Python Office for Mac Logs gatherer v0.1.1"
      "The logs will be generated in logs folder below the one you are running this from."
      "\nIMPORTANT. Each application (Word, Excel, PowerPoint, OneNote and Outlook"
      "must have been ran at least once, even if not successfully.")

logword = subprocess.check_call(['syslog | grep ' + word], shell="True", stdout=logfileword)
logexcel = subprocess.check_call(['syslog | grep ' + excel], shell="True", stdout=logfileexcel)
logpowerpoint = subprocess.check_call(['syslog | grep ' + powerpoint], shell="True", stdout=logfilepowerpoint)
logonenote = subprocess.check_call(['syslog | grep ' + onenote], shell="True", stdout=logfileonenote)
logoutlook = subprocess.check_call(['syslog | grep ' + outlook], shell="True", stdout=logfileoutlook)
loggeneral= subprocess.check_call(['syslog'], shell="True", stdout=logfilegeneral)

#  keep updated at https://github.com/audricd/PythonOfficeforMacLogs
