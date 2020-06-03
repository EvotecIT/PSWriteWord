Import-Module $PSScriptRoot\..\PSWriteWord.psd1 -Force

$T1 = [PSCustomObject] @{
    Test   = 1
    Test2  = 7
    Ole    = 'bole'
    Trolle = 'A'
    Alle   = 'sd'
}

$T2 = [ordered] @{
    Test   = 1
    Test2  = 7
    Ole    = 'bole'
    Trolle = 'A'
    Alle   = 'sd'
}


$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables9.docx"

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text "Before" -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $T1 -Design ColorfulGrid -Supress $true
Add-WordText -WordDocument $WordDocument -Text "After" -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $T1 -Transpose -Design ColorfulGrid -Supress $true

Add-WordText -WordDocument $WordDocument -Text "Before T2" -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $T2 -Design ColorfulGrid -Supress $true
Add-WordText -WordDocument $WordDocument -Text "After T2" -Supress $true
Add-WordTable -WordDocument $WordDocument -DataTable $T2 -Transpose -Design ColorfulGrid -Supress $true

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument