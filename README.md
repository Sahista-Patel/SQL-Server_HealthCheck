# SQL-Server_HealthCheck
This script will create one HTML file and send over E-Mail as an attachment. The HTML file contains below details,<br>
    •	New User DB creation report – Weekly<br>
    •	Database backup report – Monthly<br>
    •	Disk space report – Monthly<br>
    •	Sysadmin access report – Monthly<br>
    •	DBMS availability – Monthly<br>
    •	DBMS memory – Monthly<br>
    •	Database size report – Monthly<br>
    •	DBMS version – Monthly<br>
    •	Maximum number of concurrent sessions – Monthly<br>
    •	OS Specific<br>
 
The single glance of the SQL Server health check. It includes more then ten types os status with well formatted HTML file. Easy graphical interface catchy and handy information for all servers mentioned in Server List.
    
## Prerequisites

Windows OS - Powershell<br>
SqlServer Module need to be installed if not than type below command in powershell prompt.<br>
Install-Module -Name SqlServer

## Note

SqlServer Module need to be installed if not than type below command in powershell prompt.<br>
Install-Module -Name SqlServer

## Use

Open Powershell
"C:\DBA_HealthCheck.ps1"


# Input
Server List - txt file with the name of the machines/servers which to examine.<br>
Please set varibles like server list path, output file path, E-Mail id and password as and when guided by comment through code.

## Example O/P

![alt text](https://github.com/Sahista-Patel/SQL-Server_HealthCheck/blob/Powershell/healthcheck_1.PNG)<br>

![alt text](https://github.com/Sahista-Patel/SQL-Server_HealthCheck/blob/Powershell/healthcheck_2.PNG)<br>

![alt text](https://github.com/Sahista-Patel/SQL-Server_HealthCheck/blob/Powershell/healthcheck_3.PNG)

## License

Copyright 2020 Harsh & Sahista

## Contribution

* [Harsh Parecha] (https://github.com/TheLastJediCoder)
* [Sahista Patel] (https://github.com/Sahista-Patel)<br>
We love contributions, please comment to contribute!

## Code of Conduct

Contributors have adopted the Covenant as its Code of Conduct. Please understand copyright and what actions will not be abided.
