Import-Module "C:\Program Files\Microsoft Deployment Toolkit\bin\MicrosoftDeploymentToolkit.psd1"

New-PSDrive -Name "DS001"`
            -PSProvider MDTProvider `
            -Root "\\SERVER\d$\mdt\deployment-shares\production"

New-PSDrive -Name "DS002"`
            -PSProvider MDTProvider `
            -Root "\\SERVER\d$\mdt\deployment-shares\development"

Update-MDTDeploymentShare -path "DS001:"`
                            -Force `
                            -Verbose

Update-MDTDeploymentShare -path "DS002:"`
                            -Force `
                            -Verbose
