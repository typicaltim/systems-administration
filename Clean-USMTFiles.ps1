$deploymentSeverName = ""
$usmtFolderPath      = ""
$usmtSfdFolderPath   = ""
$logMovedToSfd       = ""
$logDeleted          = ""

# EMAIL ALERT CONFIGURATION

    # EMAIL FIELDS
    $to                         = @("xyz@xyz.com","xyz@xyz.com")
    $from                       = "xyz@xyz.com"

    # SERVER INFORMATION
    $server                     = "SERVER"
    $port                       = 25

Function Email-Report {
    Param (
        [Parameter(Mandatory=$false)]   [array]$emailTo           = $to,
        [Parameter(Mandatory=$false)]   [string]$emailFrom        = $from,
        [Parameter(Mandatory=$false)]   [string]$emailSubject     = $subject,
        [Parameter(Mandatory=$true)]    [string]$emailBody,
        [Parameter(Mandatory=$false)]   [array]$emailAttachment   = $attachments,
        [Parameter(Mandatory=$false)]   [string]$mailServer       = $server,
        [Parameter(Mandatory=$false)]   [string]$mailServerPort   = $port
    )
        
    $emailObject = New-Object System.Net.Mail.Mailmessage($emailFrom, $emailTo, $emailSubject, $emailBody)

        ForEach ($emailAddress in $emailTo){$emailObject.To.Add("$emailAddress")}
        ForEach ($attachment in $emailAttachment){$emailObject.Attachments.Add($attachment)}
        $emailObject.From        = $emailFrom
        $emailObject.Subject     = $emailSubject
        $emailObject.Body        = $emailBody
        $emailObject.IsBodyHtml  = $false
        
    $smtpObject = New-Object System.Net.Mail.SMTPClient($mailServer,$mailServerPort)
    $smtpObject.Send($emailObject)
    $emailObject.Dispose()
    $smtpObject.Dispose()
}

$reportDeleted=@()

ForEach ($file in (Get-ChildItem $usmtSfdFolderPath)) {

    $itemproperties = $null
    $itemproperties = (Get-ItemProperty $usmtSfdFolderPath\$file | Select-Object *)

    If ($itemproperties.CreationTime -lt (Get-Date).AddDays(-30)) {

        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name FilePath -Value $itemproperties.FullName
        $obj | Add-Member -MemberType NoteProperty -Name CreationTime -Value $itemproperties.CreationTime
        $reportDeleted += $obj

        Remove-Item -Path $usmtSfdFolderPath\$file -Destination $usmtSfdFolderPath
        
    }

}

$reportMovedToSfd=@()

ForEach ($file in (Get-ChildItem $usmtFolderPath)) {

    $itemproperties = $null
    $itemproperties = (Get-ItemProperty $usmtFolderPath\$file | Select-Object *)

    If ($itemproperties.CreationTime -lt (Get-Date).AddDays(-14)) {

        $obj = New-Object -TypeName PSObject
        $obj | Add-Member -MemberType NoteProperty -Name FilePath -Value $itemproperties.FullName
        $obj | Add-Member -MemberType NoteProperty -Name CreationTime -Value $itemproperties.CreationTime
        $reportMovedToSfd += $obj

        Move-Item -Path $usmtFolderPath\$file -Destination $usmtSfdFolderPath
        
    }

}

$reportMovedToSfd | Out-File -FilePath $logMovedToSfd
$reportDeleted | Out-File -FilePath $logDeleted

# Send the notification email that the replication is finished
Email-Report -emailSubject    "$DeploymentSeverName : Migration Files Cleaned"`
             -emailBody       "Please review the attached log files. If a migration file located in the MovedToSfd-log file is still needed, manually move it from $usmtSfdFolderPath to $usmtFolderPath . Files in the _ScheduledForDeletion folder will be deleted after their creation date is older than 30 days. Items in the Deleted-log file have been deleted from the _ScheduledForDeletion folder. Migration files with creation dates older than 14 days are automatically moved to the _ScheduledforDeletion folder. This script runs every Monday @ 8AM as a scheduled task on COVCHDEPLOY. The Script file is located at \\covchfs01\software\scripts\prodcution\deployment-server\Clean-UsmtFiles.ps1"`
             -emailAttachment @($logMovedToSfd,$logDeleted)
