import os
import time
import win32gui
import win32com.client

boolVal = True

file_path = input("Insert path to your desired file: ")

file_name = os.path.basename(file_path)

def isFileOpen(file_path):
    try:
        os.rename(file_path, file_path)
        return False
    except OSError as e:
        return True

while boolVal:
    if isFileOpen(file_path) == True:
        time.sleep(3)
        hndl = win32gui.FindWindow(None, file_name + " - Excel")
        
        if(hndl == win32gui.GetForegroundWindow()):
            win32gui.ShowWindow(hndl, 3)
            pass
        else:
            try:   
                win32gui.ShowWindow(hndl, 3) #Alternative: 9
                win32gui.BringWindowToTop(hndl)
                win32gui.SetForegroundWindow(hndl) 
            except win32gui.error as e:
                print("Fehler:" + str(e))
                
    else:
        
        sheetList = []
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        
        ExcelWorkbook = ExcelApp.Workbooks.Open(file_path)
        
        for x in ExcelWorkbook.Sheets:
            sheetList.append(x.Name)
        
        print("Following sheets have been found: ")
        print(*sheetList, sep= ", ")
        target_sheet = input("Which sheet should be focused? (Type the exact name): ")
        
        ExcelWorkbook.Worksheets(target_sheet).Activate()
        
        ExcelApp.Visible = True
        
        time.sleep(5)
        
    time.sleep(5)

