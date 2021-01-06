<#
.SYNOPSIS
    This script will create one HTML file and send over E-Mail as an attachment.
    The HTML file contains below details,
    •	New User DB creation report – Weekly
    •	Database backup report – Monthly
    •	Disk space report – Monthly
    •	Sysadmin access report – Monthly
    •	DBMS availability – Monthly
    •	DBMS memory – Monthly
    •	Database size report – Monthly
    •	DBMS version – Monthly
    •	Maximum number of concurrent sessions – Monthly
    •	OS Specific
    
.DESCRIPTION
    The single glance of the SQL Server health check.
    It includes more then ten types os status with well formatted HTML file.
    Easy graphical interface catchy and handy information for all servers mentioned in Server List.
    
.INPUTS
    Server List - txt file with the name of the machines/servers which to examine.
    Please set varibles like server list path, output file path, E-Mail id and password as and when guided by comment through code.

.EXAMPLE
    .\DBA_HealthCheck.ps1
    This will execute the script and gives HTML file and email with the details as an Attachment.

.NOTES
    PUBLIC
    SqlServer Module need to be installed if not than type below command in powershell prompt.
    Install-Module -Name SqlServer

.AUTHOR & OWNER
    Harsh Parecha
    Sahista Patel
#>


Import-Module SqlServer

$ServerList = "C:\example.txt"
$attach.Dispose()
#Set Email From
$EmailFrom = “example@outlook.com”
#Set Email To
$EmailTo = “example@outlook.com, example@outlook.com"
#Set Email Subject
$Subject = “DBA Health Check”
#Set SMTP Server Details
$SMTPServer = “smtp.outlook.com”
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential(“example@outlook.com”, “password”);

$HTML = "C:\example.txt"
$count = 0
$DB_User= @()
$Report = @()
$dbUser_id = 0
$date = Get-Date
$obj=Get-Content -Path $ServerList
$id = 0
$Serial_count = 1

$Acount = 0

$NewDBUser = 0
$BackupNottaken = 0
$OfflineDBs = 0
$CPULoad = 0
$MemoryLoad = 0

$Report += '

<!doctype html>
    <html lang="en">
        <head>
            <!-- Required meta tags -->
            <meta charset="utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

            <!-- Bootstrap CSS -->
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" integrity="sha384-JcKb8q3iqJ61gNV9KGb8thSsNjpSL0n8PARn9HuZOnIxN0hoP+VmmDGMN5t9UJ0Z" crossorigin="anonymous">

            <title>Report</title>
            </head>
        <body>
            <div style="position: fixed; top: 0; width: 100%;float: left; top: 0;display: block; z-index: 100;">
            <nav class="navbar navbar-expand-lg navbar-dark bg-dark" >
            <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavAltMarkup" aria-controls="navbarNavAltMarkup" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
            </button>
            <div class="collapse navbar-collapse" id="navbarNavAltMarkup">
            <div class="navbar-nav">
            <a id="link1" class="nav-link active" href="#section1" onclick="hide(section1); active(this.id)">Home <span class="sr-only">(current)</span></a>
            <a id="link2" class="nav-link" href="#section2" onclick="hide(section2); active(this.id)">DB User</a>
            <a id="link3" class="nav-link" href="#section3" onclick="hide(section3); active(this.id)">Backup</a>
            <a id="link4" class="nav-link" href="#section4" onclick="hide(section4); active(this.id)">SysAdmin Access</a>
            <a id="link5" class="nav-link" href="#section5" onclick="hide(section5); active(this.id)">Availability</a>
            <a id="link6" class="nav-link" href="#section6" onclick="hide(section6); active(this.id)">DB Size</a>
            <a id="link7" class="nav-link" href="#section7" onclick="hide(section7); active(this.id)">DBMS Memory</a>
            <a id="link8" class="nav-link" href="#section8" onclick="hide(section8); active(this.id)">DBMS Version</a>
            <a id="link9" class="nav-link" href="#section9" onclick="hide(section9); active(this.id)">Concurrent Sessions</a>
            <a id="link10" class="nav-link" href="#section10" onclick="hide(section10); active(this.id)">OS Specific</a>$date
            </div>
            </div>
            <p align="right"  style="margin: 5px; color:white;float: right;">Time:- '+$date+'</p>
            </nav>
            </div>
            <nav class="navbar navbar-light bg-dark" style="height: 50px; position: fixed; bottom: 0; width: 100%;float: left;display: block; z-index: 100;">
            <p align="right"  style="margin: 5px; color:white;float: right;"><a href="mailto:IN.AHS.Automation@atos.net">AHS Automation</a></p>
            <p id="footer" align="right"  style="margin: 5px;float: right; color:white;">Created By: Harsh & Sahista.</p>
            </nav>'








#--------------------------Section Start 2 (New User)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section2" style="padding-top: 60px;padding-bottom: 60px; margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query 'Select 
                                              name,
                                              type_desc,
                                              is_disabled,
                                              create_date,
                                              modify_date,
                                              default_database_name
                                                from sys.server_principals 
                                                where datediff(d,create_date,GETDATE()) < 57' -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="7">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">User Name</th>
<th scope="col">Type</th>
<th scope="col">Status</th>
<th scope="col">Created Date</th>
<th scope="col">Modified Date</th>
<th scope="col">Default Database</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){
                $NewDBUser += 1

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.name +"</td>
<td>"+ $line.type_desc +"</td>
<td>"+ $line.is_disabled +"</td>
<td>"+ $line.create_date +"</td>
<td>"+ $line.modify_date +"</td>
<td>"+ $line.default_database_name +"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 2 (New User)---------------------------------------------------------------------------

#--------------------------Section Start 3 (Backup)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section3" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>

<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query "WITH backupsetSummary
          AS ( SELECT   bs.database_name ,
                        bs.type bstype ,
                        MAX(backup_finish_date) MAXbackup_finish_date
               FROM     msdb.dbo.backupset bs
               GROUP BY bs.database_name ,
                        bs.type
             ),
        MainBigSet
          AS ( SELECT   
                        @@SERVERNAME servername,
                        db.name ,
                        db.state_desc ,
                        db.recovery_model_desc ,
                        bs.type ,
                        convert(decimal(10,2),bs.backup_size/1024.00/1024) backup_sizeinMB,
                        bs.backup_start_date,
                        bs.backup_finish_date,
                        physical_device_name,
                        DATEDIFF(MINUTE, bs.backup_start_date, bs.backup_finish_date) AS DurationMins
                        FROM     master.sys.databases db
                        LEFT OUTER JOIN backupsetSummary bss ON bss.database_name = db.name
                        LEFT OUTER JOIN msdb.dbo.backupset bs ON bs.database_name = db.name
                                                              AND bss.bstype = bs.type
                                                              AND bss.MAXbackup_finish_date = bs.backup_finish_date
                        JOIN msdb.dbo.backupmediafamily m ON bs.media_set_id = m.media_set_id
                        where  db.database_id>4
             )
         
SELECT
    name,
    recovery_model_desc,
    Last_Backup      = MAX(a.backup_finish_date),  
    Last_Full_Backup_start_Date = MAX(CASE WHEN A.type='D' 
                                        THEN a.backup_start_date ELSE NULL END),
    Last_Full_Backup_end_date = MAX(CASE WHEN A.type='D' 
                                        THEN a.backup_finish_date ELSE NULL END),
    Last_Full_BackupSize_MB=  MAX(CASE WHEN A.type='D' THEN backup_sizeinMB  ELSE NULL END),
    Full_DurationSeocnds = MAX(CASE WHEN A.type='D' 
                                        THEN DATEDIFF(SECOND, a.backup_start_date, a.backup_finish_date) ELSE NULL END),
    Last_Full_Backup_path = MAX(CASE WHEN A.type='D' 
                                        THEN a.physical_Device_name ELSE NULL END),
    Last_Diff_Backup_start_Date = MAX(CASE WHEN A.type='I' 
                                        THEN a.backup_start_date ELSE NULL END),
    Last_Diff_Backup_end_date = MAX(CASE WHEN A.type='I' 
                                         THEN a.backup_finish_date ELSE NULL END),
    Last_Diff_BackupSize_MB=  MAX(CASE WHEN A.type='I' THEN backup_sizeinMB  ELSE NULL END),
    Diff_DurationSeocnds = MAX(CASE WHEN A.type='I' 
                                        THEN DATEDIFF(SECOND, a.backup_start_date, a.backup_finish_date) ELSE NULL END),
    Last_Log_Backup_start_Date = MAX(CASE WHEN A.type='L' 
                                        THEN a.backup_start_date ELSE NULL END),
    Last_Log_Backup_end_date = MAX(CASE WHEN A.type='L' 
                                         THEN a.backup_finish_date ELSE NULL END),
    Last_Log_BackupSize_MB=  MAX(CASE WHEN A.type='L' THEN backup_sizeinMB  ELSE NULL END),
    Log_DurationSeocnds = MAX(CASE WHEN A.type='L' 
                                        THEN DATEDIFF(SECOND, a.backup_start_date, a.backup_finish_date) ELSE NULL END),
    Last_Log_Backup_path = MAX(CASE WHEN A.type='L' 
                                        THEN a.physical_Device_name ELSE NULL END),
    [Days_Since_Last_Backup] = DATEDIFF(d,(max(a.backup_finish_Date)),GETDATE())
FROM
    MainBigSet a
group by 
     servername,
     name,
     state_desc,
     recovery_model_desc
--  order by name,backup_start_date desc;" -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="14">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">DB Name</th>
<th scope="col">Recovery Model</th>
<th scope="col">Last Backup</th>
<th scope="col">Last Full Backup Start Date</th>
<th scope="col">Last Full Backup End Date</th>
<th scope="col">Last Differential Backup Start Date</th>
<th scope="col">Last Differential Backup End Date</th>
<th scope="col">Last Log Backup Start Date</th>
<th scope="col">Last Log Backup End Date</th>
<th scope="col">Days Since Last Backup</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){
if($line.Days_Since_Last_Backup -gt 7){
                $BackupNottaken++
}

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.name +"</td>
<td>"+ $line.recovery_model_desc +"</td>
<td>"+ $line.Last_Backup +"</td>
<td>"+ $line.Last_Full_Backup_start_Date +"</td>
<td>"+ $line.Last_Full_Backup_end_date +"</td>
<td>"+ $line.Last_Diff_Backup_start_Date +"</td>
<td>"+ $line.Last_Diff_Backup_end_date +"</td>
<td>"+ $line.Last_Log_Backup_start_Date +"</td>
<td>"+ $line.Last_Log_Backup_end_date +"</td>
<td>"+ $line.Days_Since_Last_Backup +"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 3 (Backup)---------------------------------------------------------------------------

#--------------------------Section Start 4 (SysAdmin))---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section4" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query "SELECT   name,type_desc,is_disabled
                                               FROM     master.sys.server_principals 
                                               WHERE    IS_SRVROLEMEMBER ('sysadmin',name) = 1
                                               ORDER BY name;" -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="4">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Name</th>
<th scope="col">Type</th>
<th scope="col">Status</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.name +"</td>
<td>"+ $line.type_desc +"</td>
<td>"+ $line.is_disabled +"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 4 (SysAdmin)---------------------------------------------------------------------------

#--------------------------Section Start 5(DB Avaibility)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section5" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query "SELECT name, state_desc, recovery_model_desc 
                                                FROM sys.databases ;" -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="4">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Name</th>
<th scope="col">Status</th>
<th scope="col">Recovery Model</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){
                if($line.state_desc -ne "ONLINE"){
                    $OfflineDBs++
                }

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.name +"</td>
<td>"+ $line.state_desc +"</td>
<td>"+ $line.recovery_model_desc +"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 5 (DB Avaibility)---------------------------------------------------------------------------

#--------------------------Section Start 6 (DB Size)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section6" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query 'SELECT DB_NAME(database_id) AS Database_Name,  
                                                SUM(size/128.0) AS Size_MB
                                                FROM sys.master_files
                                                WHERE database_id > 6 AND type IN (0,1)
                                                GROUP BY database_id' -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="3">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Database Name</th>
<th scope="col">Size MB</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.Database_Name +"</td>
<td>"+ $line.Size_MB+"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 6 (DB Size)---------------------------------------------------------------------------

#--------------------------Section Start 7 (DBMS Memory)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section7" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query 'SELECT (dosm.total_physical_memory_kb/128.0) As Total_MB, 
                                                (dosm.available_physical_memory_kb/128.0) AS Available,
                                              ((dosm.available_physical_memory_kb*100)/dosm.total_physical_memory_kb) As Percentage_Free,
                                                dosm.system_memory_state_desc
                                                FROM sys.dm_os_sys_memory dosm;' -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="5">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Total Size MB</th>
<th scope="col">Available Size MB</th>
<th scope="col">Percentage Free</th>
<th scope="col">Status</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){
$percentNotFree = [Math]::Round(100 - $line.Percentage_Free, 2);
if($percentNotFree -le 33){
         $color = "success"
}
elseif($percentNotFree -le 66){
         $color = "warning"
}
else{
         $color = "danger"
}

                $Report += "
<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.Total_MB +"</td>
<td>"+ $line.Available+"</td>
<td>"+ $line.Percentage_Free+"</td>
<td>
<div class='progress'>
    <div class='progress-bar progress-bar-striped progress-bar-animated bg-"+ $color +"' role='progressbar' style='width: "+$percentNotFree+"%' aria-valuenow='100' aria-valuemin='0' aria-valuemax='100'></div>
</div>
</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 7 (DBMS Memory)---------------------------------------------------------------------------

#--------------------------Section Start 8 (DBMS Version)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section8" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query "
                                                SELECT  
                                                SERVERPROPERTY('MachineName') AS ComputerName,
                                                SERVERPROPERTY('ServerName') AS InstanceName,  
                                                SERVERPROPERTY('Edition') AS Edition,
                                                SERVERPROPERTY('ProductVersion') AS ProductVersion,  
                                                SERVERPROPERTY('ProductLevel') AS ProductLevel;" -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="4">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Edition</th>
<th scope="col">Product Version</th>
<th scope="col">Product Level</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.Edition+"</td>
<td>"+ $line.ProductVersion+"</td>
<td>"+ $line.ProductLevel+"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 8 (DBMS Version)---------------------------------------------------------------------------

#--------------------------Section Start 9 (Concurrent Sessions)---------------------------------------------------------------------------
$Acount++
$Report += '

<div id="section9" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
<div id="accordion'+$Acount+'">'
        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

        $Report += '

<div class="card">
<div class="card-header" id="heading'+$count+'">

<h5 class="mb-0" style="display:inline-block; float:left;">
<button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
'+$_+'
</button>

</h5>
<!--
<span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
<span style="display:inline-block; float:right; padding:5px;">     </span>
<span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
-->

</div>

<div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion'+$Acount+'">
<div class="card-body">'
        
            $Inst_list = $_ | Foreach-Object {Get-ChildItem -Path "SQLSERVER:\SQL\$_"} 

            $count++

            

            Foreach ($Inst_list_item in $Inst_list){
                
                $Result = Invoke-Sqlcmd -Query "SELECT DB_NAME(dbid) AS DBName,
                                                COUNT(dbid)   AS NumberOfConnections,
                                                loginame      AS LoginName,
                                                nt_domain     AS NT_Domain,
                                                nt_username   AS NT_UserName,
                                                hostname      AS HostName
                                                FROM   sys.sysprocesses
                                                WHERE  dbid > 0
                                                GROUP  BY dbid,
                                                          hostname,
                                                          loginame,
                                                          nt_domain,
                                                          nt_username
                                                ORDER  BY NumberOfConnections DESC;" -ServerInstance $Inst_list_item.Name

                if($Result -ne $null){

                $Report += '

<table class="table table-hover">
<thead>
<tr>
<th scope="col" colspan="6">'+ $Inst_list_item.Name +'</th>
</tr>
<tr>
<th scope="col">#</th>
<th scope="col">Database Name</th>
<th scope="col">Number Of Connections</th>
<th scope="col">Login Name</th>
<th scope="col">NT_Domain</th>
<th scope="col">NT_UserName</th>
</tr>
</thead>
<tbody>'

                ForEach($line in $Result){

                $Report += "

<tr>
<td>"+ ($Serial_count++) +"</td>
<td>"+ $line.DBName+"</td>
<td>"+ $line.NumberOfConnections+"</td>
<td>"+ $line.LoginName+"</td>
<td>"+ $line.NT_Domain+"</td>
<td>"+ $line.NT_UserName+"</td>
</tr>"

                }
                $Serial_count = 1

                $Report += '

</tbody>
</table>'
               

                }
            }

             $Report +=
                                '
                                
</div>
</div>
</div>'
        }
    }
    catch{

    }
}

$Report += '
          
        
</div>
</div>'

#--------------------------Section End 9 (Concurrent Sessions)---------------------------------------------------------------------------

#--------------------------Section End---------------------------------------------------------------------------

$pichart = 0
$piscript = 0

#--------------------------Section 10 Start (OS Specific)---------------------------------------------------------------------------
$Acount++
$Report += '

            <div id="section10" style="padding-top: 60px;padding-bottom: 60px;margin:10px; display:none;">
    
                <div id="accordion">'

$Report += '

                <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
                <script type="text/javascript">
                    google.charts.load("current", {packages:["corechart"]});
                    google.charts.setOnLoadCallback(drawChart);
                    function drawChart() {'

[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){

       $AVGProc = Get-WmiObject -computername "$_" win32_processor | Measure-object -property LoadPercentage -Average | Select Average 
       if($AVGProc.Average -gt 0){
        $CPULoad++
       }
        $ComputerMemory =  Get-WmiObject -Class WIN32_OperatingSystem -computerName "$_"
        $Total_Memory = ((($ComputerMemory.TotalVisibleMemorySize - $ComputerMemory.FreePhysicalMemory)*100)/ $ComputerMemory.TotalVisibleMemorySize)
        $RoundMemory = [math]::Round($Total_Memory, 2)
        if($RoundMemory -gt 0){
            $MemoryLoad++
        }

        $piscript++

$Report += '

                        var mdata'+$piscript+' = google.visualization.arrayToDataTable([
                            ["Memory", "Usage"]'

#Object For Memory
$properties=@(
    @{Name="Name"; Expression = {$_.name}},
    @{Name="Memory"; Expression = {[Math]::Round(($_.workingSetPrivate / 1mb),2)}}
)
$m = Get-WmiObject -class Win32_PerfFormattedData_PerfProc_Process -filter "Name != '_Total'" -ComputerName "$_" | Sort-Object -Property workingSetPrivate -Descending |
    Select-Object $properties -First 6

foreach($mm in $m){
    

$Report += '
    ,["'+$mm.Name+'",     '+$mm.Memory+']

'

}


$Report +='
                            
                        ]);

                        var moptions'+$piscript+' = {
                            title: "Top Memory Usage",
                            is3D: true,
                        };

                        var pdata'+$piscript+' = google.visualization.arrayToDataTable([
                            ["CPU", "Usage"]'

#Object For CPU
$properties=@(
    @{Name="Name"; Expression = {$_.name}},
    @{Name="CPU"; Expression = {$_.PercentProcessorTime/10}}   
)
$p = Get-WmiObject -class Win32_PerfFormattedData_PerfProc_Process -filter "Name != '_Total'" -ComputerName "$_" | Sort-Object -Property PercentProcessorTime -Descending |
    Select-Object $properties -First 6

foreach($pp in $p){
    

$Report += '
    ,["'+$pp.Name+'",     '+$pp.CPU+']

'

}

$Report +='

                            
                        ]);

                        var poptions'+$piscript+' = {
                            title: "Top Process Usage",
                            is3D: true,
                        };

                        var mchart'+$piscript+' = new google.visualization.PieChart(document.getElementById("Memory'+$piscript+'"));
                        var pchart'+$piscript+' = new google.visualization.PieChart(document.getElementById("Process'+$piscript+'"));
                        
                        mchart'+$piscript+'.draw(mdata'+$piscript+', moptions'+$piscript+');
                        pchart'+$piscript+'.draw(pdata'+$piscript+', poptions'+$piscript+');
                        '

        }
    }
    catch{

    }
}

$Report += '
                    }
                </script>

'



        

        
[System.IO.File]::ReadLines($ServerList) | ForEach-Object {
    
    try{
        $ol = Get-WmiObject -Class Win32_Service -ComputerName "$_"
        
        if ($ol -ne $null){
        $pichart++

        $Report += '

                    <div class="card">
                        <div class="card-header" id="heading'+$count+'">
                            <h5 class="mb-0" style="display:inline-block; float:left;">
                                <button class="btn btn-link" data-toggle="collapse" data-target="#collapse'+$count+'" aria-expanded="true" aria-controls="collapse'+$count+'">
                                   '+ $_+'
                                </button>
                            </h5>
                            <!--
                            <span class="badge badge-danger" style="display:inline-block; float:right; padding:5px;">Danger: 10</span>
                            <span style="display:inline-block; float:right; padding:5px;">     </span>
                            <span class="badge badge-warning" style="display:inline-block; float:right; padding:5px;">Warning 15</span>
                            -->
                        </div>

                        <div id="collapse'+$count+'" class="collapse" aria-labelledby="heading'+$count+'" data-parent="#accordion">
                            <div class="card-body">'
        
            

            $count++

            $disks = Get-WmiObject -ComputerName $_ -Class Win32_LogicalDisk -Filter "DriveType < 6";

            
                $Report += '

                                <div class="card-deck" >
                                    <div id="Memory'+$pichart+'" style="width: 400px;"></div>
                                    <div id="Process'+$pichart+'" style="width: 400px;"></div>
                                </div>
                                <table class="table table-hover">
                                    <thead>
                                        
                                        <tr>
                                            <th scope="col">#</th>
                                            <th scope="col">Drive</th>
                                            <th scope="col">Occupied/Total Space</th>
                                            <th scope="col">Status</th>
                                        </tr>
                                    </thead>
                                    <tbody>'
                
                
                
                foreach($disk in $disks){
                
                $deviceID = $disk.DeviceID;
                [float]$size = $disk.Size;
                [float]$freespace = $disk.FreeSpace;

                if($size -ne 0){

                $percentFree = [Math]::Round(($freespace / $size) * 100, 2);
                $percentNotFree = 100 - $percentFree
                $sizeGB = [Math]::Round($size / 1073741824, 2);
                $freeSpaceGB = [Math]::Round($freespace / 1073741824, 2);
                $color = $null
                $NotfreeSpaceGB = [Math]::Round($sizeGB - $freeSpaceGB, 2);

                if($percentNotFree -le 33){
                    $color = "success"
                }
                elseif($percentNotFree -le 66){
                    $color = "warning"
                }
                else{
                    $color = "danger"
                }

                $Report += "

                                        <tr>
                                            <td>"+ ($Serial_count++) +"</td>
                                            <td>" +$deviceID+ "</td>
                                            <td>"+ $NotfreeSpaceGB +" GB / "+$sizeGB+" GB</td>
                                            <td>
                                            <div class='progress'>
                                              <div class='progress-bar progress-bar-striped progress-bar-animated bg-"+ $color +"' role='progressbar' style='width: "+$percentNotFree+"%' aria-valuenow='100' aria-valuemin='0' aria-valuemax='100'></div>
                                            </div>
                                            
                                            </td>
                                        </tr>"

                    }
                }
                $Serial_count = 1

                $Report += '

                                    </tbody>
                                </table>'
               

                

             $Report +='
                                
                            </div>
                        </div>
                    </div>'
        }
    }
    catch{

    }
}

$Report += '
          
                </div>
            </div>'

#--------------------------Section 10 End (OS Specific)---------------------------------------------------------------------------

#--------------------------Section Start 1---------------------------------------------------------------------------
$Report += '

<div id="section1" style="padding-top: 60px;margin:10px;">
<div class="card-deck">
                <div class="card text-white bg-secondary mb-3" style="max-width: 18rem;">
                <div class="card-header">New Users</div>
                <div class="card-body">
                    <h5 class="card-title">'+$NewDBUser+'<i class="fa fa-user-plus" style="float : right; font-size:26px;"></i></h5>
                    <p class="card-text">Within 57 Days</p>
                </div>
                </div>
                 <div class="card text-white bg-success mb-3" style="max-width: 18rem;">
                <div class="card-header">Pending Backup</div>
                <div class="card-body">
                    <h5 class="card-title">'+$BackupNottaken+'<i class="fas fa-hdd" style="float : right; font-size:26px;"></i></h5>
                    <p class="card-text">Since Last 7 Days</p>
                </div>
                </div>
                 <div class="card text-white bg-primary mb-3" style="max-width: 18rem;">
                <div class="card-header">Offline DBs</div>
                <div class="card-body">
                    <h5 class="card-title">'+$OfflineDBs+'<i class="fa fa-database" style="float : right; font-size:26px;"></i></h5>
                    <p class="card-text">Current In Offline State</p>
                </div>
                </div>
                 <div class="card text-white bg-danger mb-3" style="max-width: 18rem;">
                <div class="card-header">Memory Alerts</div>
                <div class="card-body">
                    <h5 class="card-title">'+$MemoryLoad+'<i class="fas fa-memory" style="float : right; font-size:26px;"></i></h5>
                    <p class="card-text">Beyond Threshold</p>
                </div>
                </div>
                 <div class="card text-white bg-warning mb-3" style="max-width: 18rem;">
                <div class="card-header">CPU Alerts</div>
                <div class="card-body">
                    <h5 class="card-title">'+$CPULoad+'<i class="fas fa-server" style="float : right; font-size:26px;"></i></h5>
                    <p class="card-text">Beyond Threshold</p>
                </div>
                </div>
</div>
</div>'

#--------------------------Section End 1---------------------------------------------------------------------------


$Report +='

<!-- Optional JavaScript -->
<!-- jQuery first, then Popper.js, then Bootstrap JS -->
<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js" integrity="sha384-9/reFTGAW83EW2RDu2S0VKaIzap3H66lZH81PoYlFhbGU+6BZp6G7niu735Sk7lN" crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js" integrity="sha384-B4gt1jrGC7Jh4AgTPSdUtOBvfO8shuf57BaghqFfPlYxofvL8/KUEfYiJOMMV+rV" crossorigin="anonymous"></script>
<script>

var b = "Cre"
b += "ate"
b += "d B"
b += "y: H"

function hide(val) {

section1.style.display = "none";
section2.style.display = "none";
section3.style.display = "none";
section4.style.display = "none";
section5.style.display = "none";
section6.style.display = "none";
section7.style.display = "none"; 
section8.style.display = "none";              
section9.style.display = "none";
section10.style.display = "none";
val.style.display = "block";

}
var a = document.getElementById("footer");

b += "ars"
b += "h & S"
b += "ahi"
b += "sta."

if(a){
    document.getElementById("footer").innerHTML = b;
  }
else{
  alert("Change found in code.");
  document.body.innerHTML = "Please do not make any changes in code.";
  }
function active(id) {

document.getElementById("link1").classList.remove("active");
document.getElementById("link2").classList.remove("active");
document.getElementById("link3").classList.remove("active");
document.getElementById("link4").classList.remove("active");
document.getElementById("link5").classList.remove("active");
document.getElementById("link6").classList.remove("active");
document.getElementById("link7").classList.remove("active");
document.getElementById("link8").classList.remove("active");
document.getElementById("link9").classList.remove("active");
document.getElementById("link10").classList.remove("active");

document.getElementById(id).classList.add("active");

}

</script>
   
   
 <script src="https://kit.fontawesome.com/a076d05399.js"></script>
 
</body>
</html>'

Set-Content $HTML $Report 

sleep(5)

$attach = new-object Net.Mail.Attachment($HTML) 

$Body = "Please download attached Report"
$SMTPClient.EnableSsl = $true
# Create the message
$mail = New-Object System.Net.Mail.Mailmessage $EmailFrom, $EmailTo, $Subject, $Body
$mail.Attachments.Add($attach) 
$mail.IsBodyHTML=$true
$SMTPClient.Send($mail)
$attach.Dispose()
