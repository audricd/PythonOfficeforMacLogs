import subprocess

#logs.py logs variables and functions
word = "Microsoft\ Word"
excel = "Microsoft\ Excel"
powerpoint = "Microsoft\ PowerPoint"
onenote = "Microsoft\ OneNote"
outlook = "Microsoft\ Outlook"

logfileword = open("MSword.txt", "w")
logfileexcel = open("MSexcel.txt", "w")
logfilepowerpoint = open("MSppt.txt", "w")
logfileonenote = open("MSonenote.txt", "w")
logfileoutlook = open("MSoutlook.txt", "w")

print("Office for Mac Logs gatherer. The logs will be generated in the same folder you are running this from.")

logword = subprocess.check_call(['syslog | grep ' + word], shell="True", stdout=logfileword)
logexcel = subprocess.check_call(['syslog | grep ' + excel], shell="True", stdout=logfileexcel)
logpowerpoint = subprocess.check_call(['syslog | grep ' + powerpoint], shell="True", stdout=logfilepowerpoint)
logonenote = subprocess.check_call(['syslog | grep ' + onenote], shell="True", stdout=logfileonenote)
logoutlook = subprocess.check_call(['syslog | grep ' + outlook], shell="True", stdout=logfileoutlook)
