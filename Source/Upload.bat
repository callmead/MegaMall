@ echo off
Call Home
echo Ready to Upload Data!
echo **********************************************  
echo.
pause
echo Uploading Data...
echo.
mysql -u root -psamsung MM_POS < D:\Backup\Backup.sql
echo.
echo **********************************************
echo Data uploaded Sucessfully, Thank you.