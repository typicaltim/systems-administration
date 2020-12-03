$ComputerList = @(Get-ADComputer -Filter * -SearchBase $SearchBase -SearchScope 2 | select Name)

ForEach ($Computer in $ComputerList) {

    Invoke-Command -ComputerName $Computer.Name -ErrorAction SilentlyContinue -ScriptBlock {

        Get-WmiObject Win32_Volume -Filter "DriveType='3'" | ForEach {
                New-Object PSObject -Property @{
                    Name = $_.Name
                    Label = $_.Label
                    FreeSpace_GB = ([Math]::Round($_.FreeSpace /1GB,2))
                    TotalSize_GB = ([Math]::Round($_.Capacity /1GB,2))
                    UsedSpace_GB = ([Math]::Round($_.Capacity /1GB,2)) - ([Math]::Round($_.FreeSpace /1GB,2))
                }
            } | Where {$_.Name -eq "C:\"} | Select FreeSpace_GB
    }
}
