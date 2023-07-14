$AccessRecords = New-Object System.Collections.Generic.List[System.Object]

$Destination = Read-Host 'Enter destination path: '
New-Item -Path $Destination'\OSTs' -ItemType Directory
$Destination = Get-Item $Destination'\OSTs'

Write-Output "1`tDESTINATION FOLDER`t$Destination`n" > OstMove.log

$TemplateAcl = Get-Acl $Destination

## COMMENT IF YOU ALREADY RUN SCRIPT AND HAVE EXISTING DESTINATION FOLDER WITH CUSTOMIZED ACCESS RULES
# disable inheritance
$TemplateAcl.SetAccessRuleProtection($true,$false)
# set basic rules for Administrators and SYSTEM
$TemplateAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators","FullControl","Allow")))
$TemplateAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM","FullControl","Allow")))
$TemplateAcl | Set-Acl $Destination.FullName

Write-Output "2`tACLs FOR THE DESTINATION FOLDER" >> OstMove.log
Write-Output $TemplateAcl.AccessToString >> OstMove.log

$UsersFolders = Get-ChildItem C:\Users\ | Select-Object Name, FullName | Where-Object {($_.FullName -ne 'C:\Users\Public')} # -and ($_.FullName -ne 'C:\Users\Administrator')}

Write-Output "`n3`tUSERS FOLDERS LIST" >> OstMove.log
Write-Output $UsersFolders.FullName >> OstMove.log

Foreach ($UserFolder in $UsersFolders) {
    
    Write-Output "`n4`tUSER:"($UserFolder.Name)"" >> OstMove.log

    $Osts = $False
    $Osts = ($UserFolder | Select-Object -ExpandProperty FullName) + '\AppData\Local\Microsoft\Outlook\'
    $Osts = Get-ChildItem $Osts -ErrorAction 'SilentlyContinue' | Where-Object {$_.Name -match '.ost'} #| Select-Object -ExpandProperty FullName

    if ($Osts) {

        Write-Output "`n4`tOSTS EXIST:"($Osts.FullName)"" >> OstMove.log

        Foreach ($Ost in $Osts) {
            
            $AccessRecord = New-Object PSObject -Property @{
                FolderFullPath = $FolderFullPath
                Owner = $Owner
                DestinationFile = $DestinationFile
                UsersAccessRule = $UsersAccessRule
            }

            $AccessRecord.FolderFullPath = $Ost.Directory.FullName + '\'
            $AccessRecord.Owner = $UserFolder.Name
            $AccessRecord.UsersAccessRule = (Get-Acl $Ost.FullName).Access | Where-Object {($_.IdentityReference -ne 'NT AUTHORITY\SYSTEM') -and ($_.IdentityReference -ne 'BUILTIN\Administrators')}
        

            if (!(Get-ChildItem $Destination | Where {$_.Name -eq $Ost.Name}) -or ($Ost.Length -gt (Get-ChildItem $Destination | Where-Object {$_.Name -eq $Ost.Name}).Length)) {
                # if (Get-ChildItem $Destination | Where {$_.Name -eq $Ost.Name}) {
                Remove-Item (Get-ChildItem $Destination | Where {$_.Name -eq $Ost.Name}).FullName -ErrorAction SilentlyContinue
                # }
                Copy-Item -Path $Ost.FullName -Destination $Destination
                if ((Get-ChildItem $Destination).Name -match $Ost.Name) {
                    Write-Output "`n4`tOST SUCCESSFULLY COPIED, MAKING A LINK" >> OstMove.log
                    Remove-Item -Path $Ost.FullName
                    New-Item -Path $Ost.FullName -ItemType SymbolicLink -Value (Get-ChildItem $Destination |Where {$_.Name -eq $Ost.Name}).FullName
                    Write-Output (Get-Item $Ost.FullName | Select Mode, Length, FullName | fl) >> OstMove.log
                } else {
                    Write-Output "`n4`tOST COULD NOT BE COPIED FOR SOME REASON, PLEASE CHECK IT MANUALLY" >> OstMove.log
                }
            } else {
                Write-Output "`n4`tOST ALREADY EXISTS, MAKING A LINK" >> OstMove.log
                Remove-Item -Path $Ost.FullName
                New-Item -Path $Ost.FullName -ItemType SymbolicLink -Value (Get-ChildItem $Destination |Where {$_.Name -eq $Ost.Name}).FullName
                Write-Output (Get-Item $Ost.FullName | Select Mode, Length, FullName | fl) >> OstMove.log
            }

            $AccessRecord.DestinationFile = Get-ChildItem $Destination | Where-Object {$_.Name -eq $Ost.Name}
            $AccessRecords.Add($AccessRecord)

            Write-Output "`n4`tADDED ACCESS RECORD"($AccessRecord | fl)""($AccessRecord.UsersAccessRule)"" >> OstMove.log
        }
    } else {
        Write-Output "`n4`tOST DOES NOT EXIST FOR THIS USER" >> OstMove.log
    }
}

$Files = $AccessRecords.DestinationFile | Get-Unique

Write-Output "`n5`tGETTING UNIQUE FILES LIST FROM ACCESS RECORDS" >> OstMove.log
Write-Output $Files >> OstMove.log

foreach ($file in $Files) {
    Write-Output "`n6`tADDING ACCESS RULES FOR THE FILE:"($File.VersionInfo.FileName)"" >> OstMove.log
    $TemplateAcl = Get-Acl $Destination
    $FileAcl = $TemplateAcl
    $FileRules = $AccessRecords | Where-Object {$_.DestinationFile -match $File.Name}
    Write-Output "`n6`tADDING RULES FROM THE FILES:"($FileRules.FolderFullPath)"" >> OstMove.log
    foreach ($FileRule in $FileRules.UsersAccessRule) {
        $FileAcl.AddAccessRule($FileRule)
    }
    $FileAcl | Set-Acl $file.FullName
    Write-Output "`n6`tADDED RULES:"($FileAcl.AccessToString)"" >> OstMove.log
}

Write-Output "`n7`tALL TASKS COMPLETED, PLEASE SEE THE ERRORS LIST BELOW:" >> OstMove.log
Write-Output $Error >> OstMove.log
