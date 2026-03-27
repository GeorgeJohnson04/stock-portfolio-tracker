@echo off
schtasks /create /tn "Finance Portfolio - Hourly Price Refresh" /tr "\"C:\Users\Tommy_w7c1d3j\AppData\Local\Programs\Python\Python312\python.exe\" \"C:\Users\Tommy_w7c1d3j\OneDrive\Desktop\Claude\Finance\refresh_prices.py\"" /sc hourly /mo 1 /f
schtasks /create /tn "Finance Portfolio - Login Price Refresh" /tr "\"C:\Users\Tommy_w7c1d3j\AppData\Local\Programs\Python\Python312\python.exe\" \"C:\Users\Tommy_w7c1d3j\OneDrive\Desktop\Claude\Finance\refresh_prices.py\"" /sc onlogon /f
echo.
echo Scheduled tasks created successfully!
echo - Hourly refresh: runs every 1 hour
echo - Login refresh: runs when you log in to Windows
pause
