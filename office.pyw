# import required modules
from tkinter import *
import subprocess
import datetime
import getpass #Used to fetch username only


# Set functions
def startWord():
    subprocess.Popen('C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE')
def startPowerPoint():
    subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
def startExcel():
    subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
def startOutlook():
    subprocess.Popen('C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE')
def startAccess():
    subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\MSACCESS.EXE")
def startOneDrive():
    subprocess.Popen("C:\\Program Files (x86)\\Microsoft OneDrive\\OneDrive.exe")


# Pre-script
window = Tk()
window.overrideredirect(False)
window.geometry("550x370")
window.title("Office")
window.minsize(450, 360)
window.maxsize(900, 500)


# Color code for back and foreground colors in different color mode
lightBg="#FAF9F8"
lightFg="#252423"
darkBg="#1B1A19"
darkFg="#E9E8E8"


# Fetch current system username for greet string later
currentUser=getpass.getuser()


# Fetch current time to determine the color mode to be rendered
currentTime=datetime.datetime.now()
currentTime.hour


# Deetrmine whether it's morning, afternoon or evening to print greets based on that
if currentTime.hour < 12:
    greetString=("Good morning, {}".format(currentUser))
elif 12 <= currentTime.hour < 18:
    greetString=("Good afternoon, {}".format(currentUser))
else:
    greetString=("Good evening, {}".format(currentUser))


# Create a function called darkMode to be executed for dark mode
def darkMode():
    windowColorMode=darkBg
    window.configure(bg=windowColorMode)
    firstLine=Label(master=window, text=greetString, font=("Arial Bold", 18), bg=darkBg, fg=darkFg)
    firstLine.place(relx = 0.5, rely = 0.1, anchor = "center")
    secondLine=Label(window, text="Office apps avaliable:", font=("Arial Bold", 13), bg=darkBg, fg=darkFg)
    secondLine.place(relx = 0.3, rely = 0.23, anchor = "center")
    thirdLine=Label(window, text="More apps:", font=("Arial Bold", 11), bg=darkBg, fg=darkFg)
    thirdLine.place(relx = 0.73, rely = 0.23, anchor = "center")


 # Create a function called lightMode to be executed for light mode
def lightMode():
    windowColorMode=lightBg
    window.configure(bg=windowColorMode)
    firstLine=Label(window, text=greetString, font=("Arial Bold", 18), bg=lightBg, fg=lightFg)
    firstLine.place(relx = 0.5, rely = 0.1, anchor = "center")
    secondLine=Label(window, text="Office apps avaliable:", font=("Arial Bold", 13), bg=lightBg, fg=lightFg)
    secondLine.place(relx = 0.3, rely = 0.23, anchor = "center")
    thirdLine=Label(window, text="More apps:", font=("Arial Bold", 11), bg=lightBg, fg=lightFg)
    thirdLine.place(relx = 0.73, rely = 0.23, anchor = "center")


# Determine which render mode to use
if currentTime.hour < 8:
    darkMode()
elif currentTime.hour < 19:
    lightMode()
else:
    darkMode()


# Deploy all the buttons on the screen
    # Word button
wordButton=Button(window, text="Word", font=("Arial Bold", 13), width=14, height=2, bg="#1651AA", fg="white", command=startWord)
wordButton.place(relx = 0.3, rely = 0.36, anchor = "center")
wordButton["border"]="0"


    # PowerPoint button
powerPointButton=Button(window, text=" PowerPoint ", font=("Arial Bold", 13), width=14, height=2, bg="#B13719", fg="white", command=startPowerPoint)
powerPointButton.place(relx = 0.3, rely = 0.5, anchor = "center")
powerPointButton["border"]="0"


    # Excel button
excelButton=Button(window, text="      Excel      ", font=("Arial Bold", 13), width=14, height=2, bg="#0F703B", fg="white", command=startExcel)
excelButton.place(relx = 0.3, rely = 0.64, anchor = "center")
excelButton["border"]="0"


    # Outlook button
outlookButton=Button(window, text="      Outlook      ", font=("Arial Bold", 13), width=10, height=1, bg="#106EBE", fg="white", command=startOutlook)
outlookButton.place(relx = 0.73, rely = 0.33, anchor = "center")
outlookButton["border"]="0"


    # OneDrive button
oneDriveButton=Button(window, text="      OneDrive      ", font=("Arial Bold", 13), width=10, height=1, bg="#0078D4", fg="white", command=startOneDrive)
oneDriveButton.place(relx = 0.73, rely = 0.425, anchor = "center")
oneDriveButton["border"]="0"


    # Access button
accessButton=Button(window, text="      Access      ", font=("Arial Bold", 13), width=10, height=1, bg="#A4373A", fg="white", command=startAccess)
accessButton.place(relx = 0.73, rely = 0.52, anchor = "center")
accessButton["border"]="0"


# Loop the script till interrupt occurs
window.mainloop()