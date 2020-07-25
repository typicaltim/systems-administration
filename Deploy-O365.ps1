Workflow Deploy-O365 {

    # Target List
    $TargetMachinesList = @(Get-ADComputer -Filter * -SearchBase 'OU=myorgou,DC=private,DC=mydomain,DC=com' -SearchScope 2 | select Name).Name
    
    # Source Resources Configuration
    $OfficeRemovalResourcesSource = "\\server\share\Office365-Rollout\tools\Remove-PreviousOfficeInstalls\*"
    $Office2016ResourcesSource    = "\\server\share\office2016DVD\*"
    $Office365ResourcesSource     = "\\server\share\o365files\*"
    $CiscoViewmailResourcesSource = "\\server\share\Cisco Viewmail\*"
    $MicrosoftEdgeResourcesSource = "\\server\share\Microsoft Edge\*"

    # Session Variables
    $CurrentHostNumber          = 0
    $CompletedList              = @()
    $FailedList                 = @()
    $ActiveLoginList            = @()
    $HostnameConflictList       = @()
    $RemoteCommandsFailureList  = @()
    $BadInstallList             = @()
    $UnresponsiveList           = @()

    # Begin Parallel Processing
    ForEach -Parallel -ThrottleLimit 5 ($Computer in $TargetMachinesList) {
        
        # Write Progress Notification
        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer STARTING PROCESSING"

        # Increment Host Number by 1
        $WORKFLOW:CurrentHostNumber++

        # Verify that the host is online
        If (Test-Connection -ComputerName $Computer -Quiet) {

            # Check if WinRM is running
            If ([bool](Test-WSMan -ComputerName $Computer -Authentication Default -ErrorAction SilentlyContinue)) {

                # Continue if the reported hostname matches the target computer name
                If ($(Invoke-Command -ComputerName $Computer -ScriptBlock {hostname}).ToUpper() -eq "$Computer") {

                    # Check if O365 is already installed so we don't waste time
                    If ($(Invoke-Command -ComputerName $Computer -ScriptBlock {Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "Office 16 Click*"}}) -ne $null){

                        $WORKFLOW:CompletedList += "$Computer"

                    }

                        # If O365 was not present install it
                        Else {

                            # If there are any users logged in skip the machine
                            If ((quser /SERVER:$Computer 2> $null) -NE $null) {

                                $WORKFLOW:FailedList += $Computer
                                $WORKFLOW:ActiveLoginList += $Computer

                            }
                                # Continue if no users were logged in
                                Else {

                                    # Temp Directory Configuration
                                    $LocalCShare                  = "\\" + "$Computer" + "\C$"
                                    $TempDirectoryName            = "temp_O365-Rollout"
                                    $OfficeRemovalResourcesName   = "OfficeRemovalTools"
                                    $Office2016ResourcesName      = "Office2016"
                                    $Office365ResourcesName       = "Office365"
                                    $CiscoViewmailResourcesName   = "Viewmail12dot5"
                                    $MicrosoftEdgeResourcesName   = "MicrosoftEdge"
                                    $TempDirectory                = "$LocalCShare" + "\" + "$TempDirectoryName"
                                    $OfficeRemovalResources       = "$TempDirectory" + "\" + "$OfficeRemovalResourcesName"
                                    $Office2016Resources          = "$TempDirectory" + "\" + "$Office2016ResourcesName"
                                    $Office365Resources           = "$TempDirectory" + "\" + "$Office365ResourcesName"
                                    $CiscoViewmailResources       = "$TempDirectory" + "\" + "$CiscoViewmailResourcesName"
                                    $MicrosoftEdgeResources       = "$TempDirectory" + "\" + "$MicrosoftEdgeResourcesName"
                    
                                    # Run Inline Commands (Non-workflow Cmdlets)
                                    InlineScript {
                                
                                        # Execute Prep Commands (Requiring Authentication)

                                        # Create Temp Directories
                                        New-Item -Path "$USING:LocalCShare"   -Name "$USING:TempDirectoryName"          -ItemType "Directory" -Force | Out-Null
                                        New-Item -Path "$USING:TempDirectory" -Name "$USING:OfficeRemovalResourcesName" -ItemType "Directory" -Force | Out-Null
                                        New-Item -Path "$USING:TempDirectory" -Name "$USING:Office2016ResourcesName"    -ItemType "Directory" -Force | Out-Null
                                        New-Item -Path "$USING:TempDirectory" -Name "$USING:Office365ResourcesName"     -ItemType "Directory" -Force | Out-Null
                                        New-Item -Path "$USING:TempDirectory" -Name "$USING:CiscoViewmailResourcesName" -ItemType "Directory" -Force | Out-Null
                                        New-Item -Path "$USING:TempDirectory" -Name "$USING:MicrosoftEdgeResourcesName" -ItemType "Directory" -Force | Out-Null
                        
                                        # Copy Resource Files to Temp Directories
                                        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $USING:Computer COPYING RESOURCE FILES"

                                        # Copy Office removal script resources to the remote host
                                        Copy-Item -Recurse "$USING:OfficeRemovalResourcesSource" -Destination "$USING:OfficeRemovalResources" -Force

                                        # Check if Office 2016 is installed so we can copy the uninstall files only if needed
                                        $Office2016 = Get-WmiObject -Class Win32_Product | Where-Object Name -like "Microsoft Office*2016"
                                        If ($Office2016 -ne $null){
                                            Copy-Item -Recurse "$USING:Office2016ResourcesSource" -Destination "$USING:Office2016Resources" -Force
                                            $Office2016 = $null
                                        }

                                        # Copy the Office 365 resources to the remote host
                                        Copy-Item -Recurse "$USING:Office365ResourcesSource"      -Destination "$USING:Office365Resources"     -Force

                                        # Copy the Cisco Viewmail resources to the remote host
                                        Copy-Item -Recurse "$USING:CiscoViewmailResourcesSource"  -Destination "$USING:CiscoViewmailResources" -Force

                                        # Copy the Microsoft Edge resources to the remote host
                                        Copy-Item -Recurse "$USING:MicrosoftEdgeResourcesSource"  -Destination "$USING:MicrosoftEdgeResources" -Force

                                        # Execute Remote Commands
                                        Invoke-Command -ComputerName $USING:Computer -ScriptBlock 2> $null {

                                            # Set Variables

                                            # Hostname Variable
                                            $Hostname = hostname
                                            $Hostname = $Hostname.ToUpper()

                                            # Temp Directory Configuration
                                            $LocalCShare                  = "C:"
                                            $TempDirectoryName            = "temp_O365-Rollout"
                                            $OfficeRemovalResourcesName   = "OfficeRemovalTools"
                                            $Office2016ResourcesName      = "Office2016"
                                            $Office365ResourcesName       = "Office365"
                                            $CiscoViewmailResourcesName   = "Viewmail12dot5"
                                            $MicrosoftEdgeResourcesName   = "MicrosoftEdge"
                                            $TempDirectory                = "$LocalCShare" + "\" + "$TempDirectoryName"
                                            $OfficeRemovalResources       = "$TempDirectory" + "\" + "$OfficeRemovalResourcesName"
                                            $Office2016Resources          = "$TempDirectory" + "\" + "$Office2016ResourcesName"
                                            $Office365Resources           = "$TempDirectory" + "\" + "$Office365ResourcesName"
                                            $CiscoViewmailResources       = "$TempDirectory" + "\" + "$CiscoViewmailResourcesName"
                                            $MicrosoftEdgeResources       = "$TempDirectory" + "\" + "$MicrosoftEdgeResourcesName"

                                            # Remove old office versions
                                            Set-ExecutionPolicy -ExecutionPolicy Bypass
                                            cd $OfficeRemovalResources
                                            . .\Remove-PreviousOfficeInstalls.ps1
                                            Remove-PreviousOfficeInstalls

                                            # Remove Cisco Viewmail
                                            $CiscoViewmail = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Viewmail*"}
                                            If ($CiscoViewmail -ne $null){

                                                msiexec.exe /q /X $CiscoViewmail.IdentifyingNumber

                                            }
                                            $CiscoViewmail = $null

                                            # Check if Office 2016 is installed
                                            $Office2016 = Get-WmiObject -Class Win32_Product | Where-Object Name -like "Microsoft Office*2016"
                                            If ($Office2016 -ne $null){

                                                Set-Location $Office2016Resources
                                                Start-Process .\setup.exe -ArgumentList "/uninstall ProPlus","/config .\uninstall.xml" -NoNewWindow -Wait

                                            }

                                            # Install office 365
                                            "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer INSTALLING O365"
                                            Set-Location $Office365Resources
                                            Start-Process .\setup.exe -ArgumentList "/configure .\Configurations\Standard.xml" -NoNewWindow -Wait

                                            # Verify that O365 was installed before attempting to install Viewmail
                                            "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer VERIFYING INSTALL"
                                            If ($(Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "Office 16 Click*"}) -ne $null){
                                        
                                                # Install ViewMail
                                                "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer TRYING VIEWMAIL"
                                                "$CiscoViewmailResources"
                                                Set-Location $CiscoViewmailResources
                                                Start-Process .\setup.exe -ArgumentList "/I","/q" -NoNewWindow -Wait
                                            }
                                        }
                                    }

                                    # Clean up the Temp Folders
                                    Remove-Item -Recurse -Force $TempDirectory

                                    # Verify that O365 was installed before marking complete
                                    "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer VERIFYING INSTALL AGAIN"
                                    If ($(Invoke-Command -ComputerName $Computer -ScriptBlock {Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "Office 16 Click*"}}) -ne $null){

                                        # Update Machine Status to COMPLETE
                                        $WORKFLOW:CompletedList += $Computer
                                    }
                                    
                                        # If O365 was not present, that means something went wrong. Log it.
                                        Else {
                                            # Update Machine Status to FAILURE
                                            $WORKFLOW:FailedList += $Computer
                                            $WORKFLOW:BadInstallList += $Computer
                                        }
                                }  
                        }
                }
                    # The reported hostname is different than expected
                    Else {
                        $WORKFLOW:FailedList += $Computer
                        $WORKFLOW:HostnameConflictList += $Computer
                    }
            }
                # WinRM Test Failed
                Else {
                    $WORKFLOW:FailedList += $Computer
                    $WORKFLOW:RemoteCommandsFailureList += $Computer
                }
        }

        # If the host is not online, do these steps
        Else {
            $WORKFLOW:FailedList += $Computer
            $WORKFLOW:UnresponsiveList += $Computer
        }

    # After each device write the lists to the console.
    "=== COMPLETED LIST ==="
    $CompletedList
    "=== FAILED LIST ==="
    $FailedList
    "=== OFFLINE LIST ==="
    $UnresponsiveList
    "=== ACTIVE LOGIN LIST ==="
    $ActiveLoginList
    "=== HOSTNAME CONFLICT LIST ==="
    $HostnameConflictList
    "=== REMOTE COMMAND FAILURE LIST ==="
    $RemoteCommandsFailureList
    "=== BAD INSTALL LIST ==="
    $BadInstallList

    # Output script status after each host is finished processing.
    "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') | PROCESSING : ($CurrentHostNumber/$($TargetMachinesList.Count)) | COMPLETED : ($($CompletedList.Count)/$($CurrentHostNumber)) | FAILED : ($($FailedList.Count)/$($CurrentHostNumber)) | OFFLINE : ($($UnresponsiveList.Count)/$($FailedList.Count)) | ACTIVE LOGIN : ($($ActiveLoginList.Count)/$($FailedList.Count)) | HOSTNAME CONFLICT : ($($HostnameConflictList.Count)/$($FailedList.Count)) | REMOTE COMMANDS : ($($RemoteCommandsFailureList.Count)/$($FailedList.Count)) | INSTALL CONFLICT : ($($BadInstallList.Count)/$($FailedList.Count))"
    
    }

    # POST-BATCH COMMANDS HERE
    "=== FINAL COMPLETED LIST ==="
    $CompletedList
    "=== FINAL FAILED LIST ==="
    $FailedList
    "=== FINAL OFFLINE LIST ==="
    $UnresponsiveList
    "=== FINAL ACTIVE LOGIN LIST ==="
    $ActiveLoginList
    "=== FINAL HOSTNAME CONFLICT LIST ==="
    $HostnameConflictList
    "=== FINAL REMOTE COMMAND FAILURE LIST ==="
    $RemoteCommandsFailureList
    "=== FINAL BAD INSTALL LIST ==="
    $BadInstallList
}
