<#
Requirements:
1) Databases on SQL Server must be set offline.  Only databases set offline will be copied and detached.
2) FCIV on the root of C: at C:\FCIV\FCIV.exe.  Download from https://support.microsoft.com/en-us/kb/841290
3) Admin access to SQL Server on SQL Server
4) Write access to the Archive location
#>
$env:PSModulePath = "C:\users\user\My Documents\WindowsPowerShell\Modules;C:\Windows\system32\WindowsPowerShell\v1.0\Modules\;C:\Program Files (x86)\Microsoft SQL Server\110\Tools\PowerShell\Modules"
import-module sqlps
set-location c:

function formatSize($size){
    $dataFileSizeInKB=($size)*8
    
    switch($dataFileSizeInKB){

        {$_ -lt 1024}{$fileSizeFormatted = $dataFileSizeInKB.ToString() + "KB"}
        {($_ -lt 1048576) -and ($_ -ge 1024)}{$fileSizeFormatted = [math]::round((($dataFileSizeInKB)/1024),2).ToString() + "MB"}
        {($_ -lt 1073741824) -and ($_ -ge 1048576)}{$fileSizeFormatted = [math]::round(((($dataFileSizeInKB)/1024)/1024),2).ToString() + "GB"}
      
    }
    return $fileSizeFormatted
}

$SQLServer="SQLServer"
$copyLocation="\\ArchiveServer\Databases Off $SQLServer"
$logFile="$copyLocation\database.csv"
$databaseCount=(Invoke-Sqlcmd -query "select count(*) AS Count from sys.databases as system where system.state_desc = 'offline'" -serverinstance "$SQLServer").Count

write-host "Information about this run:" -ForegroundColor Cyan
write-host "    SQL Server: " -NoNewline -ForegroundColor Cyan; write-host "$SQLServer" -ForegroundColor Yellow
write-host "    Archive Location: " -NoNewline -ForegroundColor Cyan; write-host "$copyLocation" -ForegroundColor Yellow
write-host "    CSV Log File: " -NoNewline -ForegroundColor Cyan; write-host "$logFile" -ForegroundColor Yellow

if($databaseCount -gt "0"){

    $databases=(Invoke-Sqlcmd -query "select system.name AS Name,master.physical_name AS Location,system.state_desc AS State,master.size, master.database_id AS 'Database ID' from sys.master_files AS master left outer join sys.databases as system on master.database_id = system.database_id where system.state_desc = 'offline' order by name" -serverinstance "$SQLServer")

    $sumOfSizes=($databases.size | measure-object -sum).sum

    $fileSizeFormatted=formatSize $sumOfSizes

    write-host "    Total Offline Databases: " -NoNewline -ForegroundColor Cyan; write-host "$databasecount" -ForegroundColor Yellow
    write-host "    Total size of files to be copied: " -NoNewline -ForegroundColor Cyan; write-host "$fileSizeFormatted`n" -ForegroundColor Yellow

    write-host "Saving the database information to CSV Log File`n" -ForegroundColor Cyan
    $databases | export-csv -NoTypeInformation -Path "$logFile" -Append
  
    write-host "Beginning to copy Offline Database data and log files to the Archive Location.`n" -ForegroundColor Cyan
     
    foreach($item in $databases){
        $fileName=split-path $item.location -leaf
        $databaseName=$item.name
        $dataFileToCopy="\\$SQLServer.domain.com\"+ ($item.Location -replace ":","$")
        $fileSizeFormatted=formatSize ($item.size)

        write-host "Size of $fileName is $fileSizeFormatted"

        if($dataFileToCopy -like "*.ldf"){
            $toCopyLocation="$copyLocation\Logs\"
        }
        else{
            $toCopyLocation="$copyLocation\Data\"
        }
        write-host "Copying $fileName from $SQLServer to Archive Location`n" -foregroundcolor yellow
        
        switch(test-path "$dataFileToCopy" -pathtype container){
            $true{$type="folder"}
            $false{$type="file"}
        }

        if($type -eq "file"){
            (xcopy /Y "$dataFileToCopy" "$toCopyLocation") | out-null
        }
        elseif($type -eq "folder"){
            (xcopy /E /H /I "$dataFileToCopy" "$toCopyLocation\$fileName") | out-null
        }
    
        if($? -eq $true){
            write-host "    The following $type was successfully copied from $SQLServer to the Archive Location:" -foregroundcolor green
            write-host "        $dataFileToCopy" -foregroundcolor green

            if($(Invoke-Sqlcmd -query "select name from sys.databases where name = '$databaseName'" -serverinstance "$SQLServer") -ne $null){
                write-host "`Detaching database $databaseName on SQL Server $SQLServer`n" -foregroundcolor Cyan            
                invoke-sqlcmd -query "exec sp_detach_db '$databaseName', 'true'" -serverinstance "$SQLServer"
                write-host "`n    Database $databaseName detatched from SQL Server $SQLServer`n`n" -foregroundcolor green
            }
            else{
                write-host "`nDatabase $databaseName is not attached to $SQLServer, no action taken.`n`n" -foregroundcolor DarkGreen
            }
        }
        else{
            write-host "Something went wrong, the following $type was not copied from $SQLServer to the Archive Location:" -ForegroundColor Red
            write-host "$dataFileToCopy`n`n" -ForegroundColor Red
        }


    }

    write-host "Copying and Detatching databases complete.`n" -foregroundcolor Green
    write-host "Beginning comparison of original and copied data." -foregroundcolor Cyan

    $FilesLocated=$databases.Location

    foreach($file in $FilesLocated){
        $fileName=split-path $file -leaf

        $originalLocation="\\$SQLServer.domain.com\"+$file -replace ":","$"
        $originalLocationReplace=$originalLocation -replace "\\","\\"
        $originalLocationReplace=$originalLocationReplace -replace "\$","\$"

        if($fileName -like "*.ldf"){
            $newLocation="$copyLocation\Logs\"+$fileName
        }
        else{
            $newLocation="$copyLocation\Data\"+$fileName
        }

        $newLocationReplace=$newLocation -replace "\\","\\"

        $OriginalFileExist=(test-path $originalLocation)
        $NewFileExist=(test-path $newLocation)

        if($OriginalFileExist -eq $true -and $NewFileExist -eq $true){

            write-host "    Hashing Originial file, $originalLocation" -foregroundcolor Cyan
            $hash1=$(C:\fciv\fciv.exe -md5 $originalLocation)
            $hash1=($hash1 | select-string "$originalLocationReplace") -replace " $originalLocationReplace",""
            write-host "        Hashing of the Original file complete.  Hash: " -nonewline -foregroundcolor green; write-host "$hash1" -foregroundcolor yellow

            write-host "    Hashing Copied file, $newLocation" -foregroundcolor Cyan
            $hash2=$(C:\fciv\fciv.exe -md5 $newLocation)
            $hash2=($hash2 | select-string "$newLocationReplace") -replace " $newLocationReplace",""
            write-host "        Hashing of the Copied file complete.  Hash: " -nonewline -foregroundcolor green; write-host "$hash2" -foregroundcolor yellow

            if($hash1 -eq $hash2){
        
                write-host "The original and copied files, $fileName, have the same hash. Deleting the original file from $SQLServer.`n" -foregroundcolor Green
                (get-childitem "$originalLocation" | remove-item -confirm:$false -force) | out-null

                if($? -eq $true){
                    write-host "Deletion of $fileName was successful.`n" -foregroundcolor Green
                }
                else{
                    write-host "Deletion of $fileName was not successful.  Please check into this`n`n." -foregroundcolor Red
                }
            }
            else{
                write-host "$fileName is not the same in both locations.  Leaving the original file on $SQLServer." -foregroundcolor Red
                write-host "Please check that the files were copied properly.  If needed perform a new copy of the files manually." -foregroundcolor Red
            }
        }
        elseif($OriginalFileExist -eq $false -and $NewFileExist -eq $true){
            write-host "`n$fileName does not exist on $SQLServer.  Cannot perform comparison and deletion." -foregroundcolor Red
        }
        elseif($OriginalFileExist -eq $true -and $NewFileExist -eq $false){
            write-host "`n$fileName does not exist in the Archive Location.  Cannot perform comparison and deletion." -foregroundcolor Red
            write-host "Please verify the file was copied to the Archive Location.  Comparison and deletion will need to be handled manually." -foregroundcolor Red
        }
        else{
            write-host "`n$fileName does not exist on the SQL Server or in the Archive Location.  Cannot perform comparison and deletion." -foregroundcolor Red
        }
    }


}
else{
    write-host "No databases on $SQLServer have been set offline." -ForegroundColor Yellow
}