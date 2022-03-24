import tkinter as tk
from tkinter import filedialog
import win32gui
import win32com.client
import os
import time

slideCount = 0

root = tk.Tk()

root.title('User Interface WIP')

root.filename = filedialog.askopenfilename(
    initialdir="C:/Users/lnegwer/Desktop/PowerPointDB",
    title="Choose PowerPoint to upload",
    filetypes=(("PowerPoint presentations", "*.pptx"), ("All files", "*.*"))
)

#Get PowerPoint App
PowerPointApp = win32com.client.Dispatch("PowerPoint.Application")
PowerPointApp.Visible = True

#Open selected presentation
Presentation = PowerPointApp.Presentations.Open(root.filename)

#SlideCount
for ppSlide in Presentation.Slides:
    slideCount = slideCount + 1

#LastModifiedDate
TimeSinceEpoch = os.path.getmtime(root.filename)
#Convertion to readable time
ModTime = time.strftime('%d-%m-%Y %H:%M:%S', time.localtime(TimeSinceEpoch))

Presentation_Main_Info = {
    'Filename': Presentation.Name,
    'SlideCount': slideCount,
    'LastModifiedDate': ModTime,
    'Tags': ['','',''],
    'LinkToFile': ''
}

for key, value in Presentation_Main_Info.items():
    print(key, ' : ', value)

slideCount = 1
wordCount = 0
shapeCount = 0
tableCount = 0
formatSize = 0

for ppSlide in Presentation.Slides:
    print("-------------------------------")
    print("Slide: " + str(slideCount))
    print("-------------------------------")

    for ppShape in ppSlide.Shapes:
        try:
            if len(ppShape.TextFrame.TextRange.Text) != 0:
                print(Presentation.PageSetup.SlideSize)
                print(Presentation.PageSetup.SlideWidth)
                print(Presentation.PageSetup.SlideHeight)
        except:
            print("Fehler")


    for ppShape in ppSlide.Shapes:
        try:
            print(ppShape.Name + "--------------------" )
            print(ppShape.TextFrame.TextRange.Text)
            print(ppShape.Type)
            print(ppShape.Left)
            print(ppShape.Top)
            print(ppShape.Width)
            print(ppShape.Height)
        except:
            print("Fehler")
    slideCount = slideCount + 1

Presentation.Close(True)

PowerPointApp.Quit()

del PowerPointApp





root.mainloop()