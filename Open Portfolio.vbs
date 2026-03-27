Set WshShell = CreateObject("WScript.Shell")

' Run the price refresh script and wait for it to finish
pythonPath = "C:\Users\Tommy_w7c1d3j\AppData\Local\Programs\Python\Python312\python.exe"
scriptPath = "C:\Users\Tommy_w7c1d3j\OneDrive\Desktop\Claude\Finance\refresh_prices.py"

WshShell.Run """" & pythonPath & """ """ & scriptPath & """", 0, True

' Now open the Excel file
excelPath = "C:\Users\Tommy_w7c1d3j\OneDrive\Desktop\Claude\Finance\Finance Project.xlsx"
WshShell.Run """" & excelPath & """", 1, False
