Set-Location $PSScriptRoot

$files = Get-ChildItem “*.msg”

foreach ($file in $files){
    start-process -FilePath $file.fullName -Verb Print
}
