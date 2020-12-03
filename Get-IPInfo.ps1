Workflow Get-IPV4Configs {
    # Config
        # Big Groups
        $Desktops         = @(Get-ADComputer -Filter * -SearchBase '$SearchBase' -SearchScope 2 | select Name)
        #$Laptops         = @(Get-ADComputer -Filter * -SearchBase '$SearchBase' -SearchScope 2 | select Name)
        #$Tablets         = @(Get-ADComputer -Filter * -SearchBase '$SearchBase' -SearchScope 2 | select Name)
        # Target List
        $TargetMachinesList = $Desktops + $Tablets + $Laptops
        $TargetMachinesList = $TargetMachinesList.Name

    # Begin Parallel Processing
        ForEach -Parallel -ThrottleLimit 5 ($Computer in $TargetMachinesList) {
        
            # Write Progress Notification
            "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer STARTING PROCESSING"

            # Increment Host Number by 1
            $WORKFLOW:CurrentHostNumber++

            # Verify that the host is online
            If (Test-Connection -ComputerName $Computer -Quiet) {

                "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer HOST ONLINE"

                # Check if WinRM is running
                If ([bool](Test-WSMan -ComputerName $Computer -Authentication Default -ErrorAction SilentlyContinue)) {

                    "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer WINRM RUNNING"

                    # Continue if the reported hostname matches the target computer name
                    If ($(Invoke-Command -ComputerName $Computer -ScriptBlock {hostname}).ToUpper() -eq "$Computer") {

                        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer HOSTNAME MATCHES"

                        InlineScript {

                            # Run it @_@
                            Invoke-Command -ComputerName $USING:Computer -ScriptBlock {

                                "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $(hostname) RUNNING REMOTE COMMANDS"
                            
                                $Hostname  = hostname ; $Hostname = $Hostname.ToUpper()
                                $IPAddress = Get-NetIPAddress -InterfaceAlias "Ethernet" -AddressState Preferred -AddressFamily IPv4 | Select-Object IPAddress ; $IPAddress = $IPAddress.IPAddress
                                $DNSConfig = Get-DnsClientServerAddress -InterfaceAlias "Ethernet" -AddressFamily "IPv4" | Select-Object ServerAddresses ; $DNSConfig = $DNSConfig.ServerAddresses

                                Remove-Item -Path C:\ipconfig-slash-all-results.txt -Force
                                Remove-Item -Path C:\IPV4-Config-Dump.txt -Force

                                Add-Content -Path "C:\IPV4-Config-Dump.txt" -Value "$Hostname, $IPAddress, $DNSConfig"

                            }

                        }

                        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer COMPILING RESULTS"
                        Add-Content -Path \\server\share\GLOBALIPV4CONFIG.txt -Value (Get-Content -Path \\$Computer\c$\IPV4-Config-Dump.txt)

                        $WORKFLOW:CompletedList += $Computer

                    }
                        # The reported hostname is different than expected
                        Else {
                            "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer HOSTNAME DIFFERENT"
                            $WORKFLOW:FailedList += $Computer
                            $WORKFLOW:HostnameConflictList += $Computer
                        }
                }
                    # WinRM Test Failed
                    Else {
                        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer WINRM NOT RUNNING"
                        $WORKFLOW:FailedList += $Computer
                        $WORKFLOW:RemoteCommandsFailureList += $Computer
                    }
            }

            # If the host is not online, do these steps
            Else {
                "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') $Computer HOST OFFLINE"
                $WORKFLOW:FailedList += $Computer
                $WORKFLOW:UnresponsiveList += $Computer
            }

        # Output script status after each host is finished processing.
        "$(Get-Date -Format 'yyyy-MM-dd-HH:mm') PROCESSING : ($CurrentHostNumber/$($TargetMachinesList.Count)) | COMPLETED : ($($CompletedList.Count)/$($CurrentHostNumber)) | FAILED : ($($FailedList.Count)/$($CurrentHostNumber)) | OFFLINE : ($($UnresponsiveList.Count)/$($FailedList.Count)) | ACTIVE LOGIN : ($($ActiveLoginList.Count)/$($FailedList.Count)) | HOSTNAME CONFLICT : ($($HostnameConflictList.Count)/$($FailedList.Count)) | REMOTE COMMANDS : ($($RemoteCommandsFailureList.Count)/$($FailedList.Count)) | INSTALL CONFLICT : ($($BadInstallList.Count)/$($FailedList.Count))"
    
        }
}
