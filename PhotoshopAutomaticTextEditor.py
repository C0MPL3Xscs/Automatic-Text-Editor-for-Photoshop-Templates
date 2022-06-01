import win32com.client
import os
import tkinter as tk


## ---------- This code was made by C0MPL3X ---------- ##
## ---------- https://github.com/C0MPL3Xscs ---------- ##
##---------- https://twitter.com/C0MPL3Xscs ---------- ##

##--------------EDIT HERE: -----------------------##
Title = ""                          ##Enter here the template title 

Text1 = ""                          ##Enter text here [ex. "ENTER HERE YOUR NICKNAME:"]

source = r"Source"                  ## ENTER HERE THE .psd FILE PATH [ex. C:\\Users\\Admin\\DeskTop\\Test.psd] (DO NOT REMOVE THE "r")

destination = "Destination"         ## ENTER HERE WHERE YOU WANT TO SAVE THE FINAL .PNG (Enter at the end the name of the new file [EX. New.png]) [ex. C:\\Users\\Admin\\DeskTop\\New.png]

layername = ""                      ## ENTER HERE THE NAME OF THE TEXT LAYER YOU WANT TO EDIT [Layer cant be inside a group on the photoshop project]
##------------------------------------------------##

## DO NOT CHANGE THE CODE BELOW OR IT MIGHT STOP WORKING!!!

##-------------------------------------------- APP UI --------------------------
root= tk.Tk()
canvas1 = tk.Canvas(root, width = 400, height = 300,  relief = 'raised')
canvas1.pack()

label1 = tk.Label(root, text=Title) 
label1.config(font=('helvetica', 9))
canvas1.create_window(200, 45, window=label1)

label2 = tk.Label(root, text=Text1) 
label2.config(font=('helvetica', 10))
canvas1.create_window(200, 100, window=label2)

label3 = tk.Label(root, text='C0MPL3X©')
label3.config(font=('helvetica', 14))
canvas1.create_window(200, 25, window=label3)

entry1 = tk.Entry (root) 
canvas1.create_window(200, 140, window=entry1)
##-------------------------------------------- PHOTOSHOP SCRIPT ----------------------------
def ChangeName():

    new_name = entry1.get()

    psApp = win32com.client.Dispatch("Photoshop.Application")

    psApp.Open(source)

    doc = psApp.Application.ActiveDocument

    layer_facts = doc.ArtLayers[layername]
    text_of_layer = layer_facts.TextItem
    text_of_layer.contents = new_name

    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13  
    options.PNG8 = False  

    pngfile = destination
    doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)

    while True:
        try:
            psApp.Application.ActiveDocument.Close(2)
        except:
            break
    psApp.Quit()

    
button1 = tk.Button(text='ENTER', command=ChangeName, bg='orange', fg='white', font=('helvetica', 9, 'bold'))
canvas1.create_window(200, 180, window=button1)

root.mainloop()

