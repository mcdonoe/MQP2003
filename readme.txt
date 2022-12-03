[File:]   Readme.txt
[Author:] Matthew Barrett
[Date:]   3/10/03


[Overview:] This CD includes the files required to implement the Emissions Database system.


[Contents:]
The "IIS Emissions" folder contains all the active server pages source code.

The "Manuals" folder contains the user manuals for basic users, system operators and
system administrators. It also contains a file called Deployment and Maintenance which
explains how the IT departments should implement the Emissions Database System.

The "Student Report.doc" file is the report written by the students explaining the project,
the user manuals, system design details and source code. 

The "SQL Server Import" folder contains an Excel spreadsheet containing the five tables
required to support the Emissions Database System.  

This readme file


[Note to IT Department:]
The Deployment and Maintenance Guide explains in detail how to install and configure IIS, 
SQL Server 2000 and how to integrate the project with your servers.  
The basic steps required by the IT department, as explained in the Deployment and Maintenance
Guide are:

1. Copy .asp pages to new directory on server. (e.g. C:\Emissions Web")
2. Create a virtual directory in Windows Internet Service Manager pointing to the directory.
3. Create new Database in SQL Server (e.g. EmissionData)
4. Import "import.xls" file into the new database.
5. Change the connection string in "global.asi" to reflect SQL Server name and uid/password.
6. Open IE 6.0 and navigate to "http://<servername>/<virtualdirectoryname>"
7. Login using the name and password "administrator".
8. Change the administrator password.

[Correction:]
The two images in the "\IIS Emissions\Images" directory should be copied to the location
"C:\inetpub\WWWRoot\Images".  Failure to do this will prevent the CAE and WPI logos from appearing. 
