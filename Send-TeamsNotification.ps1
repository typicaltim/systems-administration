# Webhook Config
$WebhookURI   = <#WEBHOOK URI#>
$JsonTemplate = <#TEMPLATE FILE#>
$JsonCard     = Get-Content $JsonTemplate | Out-String | ConvertFrom-Json

# Gather
$myText = "If any issues are detected with this deployment, please reply to this notification thread and @mention Tim so that the task sequence can be improved."
$JsonCard.sections | % {$_.text = $myText}
$JsonCard.sections.facts[0].value = hostname
$JsonCard.sections.facts[1].value = $args[0]
$JsonCard.sections.facts[2].value = "$(Get-Date -Format "dddd yyyy.MM.dd HH:mm")"
$JsonCard.sections.facts[3].value = "\\\\SERVER\productionDS$\Logs\$($JsonCard.sections.facts[0].value)\"

# Convert
$JsonCard = $JsonCard | ConvertTo-Json -Depth 5

# Send
Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $JsonCard -Uri $WebhookURI
