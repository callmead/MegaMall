@ Echo Off
Call Home
echo Processing Request Please Wait...
echo **********************************************  
echo.
mysqldump -u root -psamsung MM_POS > D:\Backup\Backup.sql
echo.
echo Database Backup File [D:\Backup\Backup.sql] Created...
echo.
echo **********************************************
echo Data Downloaded Sucessfully, Thank you.