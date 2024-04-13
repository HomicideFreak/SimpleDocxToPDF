import customtkinter
import threading
import os
import comtypes.client

customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("500x500")
root.title("Simple Docx to PDF Converter")
root.resizable(False, False)
iconPath = os.getcwd()
root.iconbitmap(iconPath+"\\Resources\\icon.ico")

inFilePath = None
outFilePath = None
inFolderFiles = []

def conv(pathIn, pathOut, entry):
    word = comtypes.client.CreateObject("Word.Application")
    entryValue = entry.get()
    if os.path.isfile(pathIn):
        doc = word.Documents.Open(os.path.abspath(pathIn))
        fileName = os.path.splitext(os.path.basename(pathIn))[0]
        if outFilePath != None and outFilePath != "":
            if entryValue == "" or entryValue == None:
                print("")
                print(os.path.abspath(pathOut  + "/" + fileName + ".pdf"))
                print("")
                doc.SaveAs(os.path.abspath(pathOut  + "/" + fileName + ".pdf"), FileFormat=17)
            else:
                print("")
                print(os.path.abspath(pathOut  + "/" + entryValue + ".pdf"))
                print("")
                doc.SaveAs(os.path.abspath(pathOut  + "/" + entryValue + ".pdf"), FileFormat=17)
        else:
            pathOut = os.path.dirname(os.path.abspath(pathIn))
            if entryValue == "" or entryValue == None:
                print("")
                print(os.path.abspath(pathOut + "/" + fileName + ".pdf"))
                print("")
                doc.SaveAs(os.path.abspath(pathOut + "/" + fileName + ".pdf"), FileFormat=17)
            else:
                print("")
                print(os.path.abspath(pathOut + "/" + entryValue + ".pdf"))
                print("")
                doc.SaveAs(os.path.abspath(pathOut + "/" + entryValue + ".pdf"), FileFormat=17)
        doc.Close()
    else:
        if outFilePath != None and outFilePath != "":
            for file in inFolderFiles:
                doc = word.Documents.Open(os.path.abspath(file))
                fileName = os.path.splitext(os.path.basename(file))[0]
                doc.SaveAs(os.path.abspath(pathOut  + "/" + fileName + ".pdf"), FileFormat=17)
                doc.Close()
        else:
            pathOut = os.path.abspath(pathIn)
            print(pathOut)
            for file in inFolderFiles:
                doc = word.Documents.Open(os.path.abspath(file))
                fileName = os.path.splitext(os.path.basename(file))[0]
                doc.SaveAs(os.path.abspath(pathOut  + "/" + fileName + ".pdf"), FileFormat=17)
                doc.Close()
    word.Quit()
        
def openFileDialInFile(label1):
    global inFilePath
    inFilePath = customtkinter.filedialog.askopenfilename()
    label1.configure(text = "In File Path: " + inFilePath)
    
def openFileDialOutFile(label2):
    global outFilePath, noNameOutFilePath
    outFilePath = customtkinter.filedialog.askdirectory()
    label2.configure(text = "Out File Path: " + outFilePath)
    
def openFileDialInFileFolder(label1):
    global inFilePath
    inFilePath = customtkinter.filedialog.askdirectory()
    label1.configure(text = "In File Folder Path: " + inFilePath)
    for file_name in os.listdir(inFilePath):
        if os.path.isfile(os.path.join(inFilePath, file_name)):
            inFolderFiles.append(os.path.join(inFilePath, file_name))
        

frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Simple Docx To PDF Converter", font=("Arial", 24))
label.pack(pady=12, padx=10)

entry1 = customtkinter.CTkEntry(master=frame, placeholder_text="Enter Filename (Leave For Unchanged)", width=230)
entry1.pack(pady=12, padx=10)

button_file_in = customtkinter.CTkButton(master=frame, text="Choose In File Path", command=lambda: openFileDialInFile(label1))
button_file_in.pack(pady=12, padx=10)

button_file_in = customtkinter.CTkButton(master=frame, text="Choose File Folder", command=lambda: openFileDialInFileFolder(label1))
button_file_in.pack(pady=12, padx=10)

label1_Text = "In File Path: " + str(inFilePath)
label1 = customtkinter.CTkLabel(master=frame, text=label1_Text)
label1.pack(pady=12, padx=10)

button_file_out = customtkinter.CTkButton(master=frame, text="Choose Out File Path", command=lambda: openFileDialOutFile(label2))
button_file_out.pack(pady=12, padx=10)

label2_Text = "Out File Path: Default (Same as Input)"
label2 = customtkinter.CTkLabel(master=frame, text=label2_Text)
label2.pack(pady=12, padx=10)

button = customtkinter.CTkButton(master=frame, text="Convert", font=("Arial", 24), width=150, height=50, command=lambda: threading.Thread(target=conv, args=(inFilePath, outFilePath, entry1), daemon=True).start())
button.pack(pady=12, padx=10)

root.mainloop()