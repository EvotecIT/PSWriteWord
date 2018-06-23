<#
    Work in progress. Doesn't work.
#>

Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-PageOrientation1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21
$p = Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $false

Set-WordPageSettings -WordDocument $WordDocument
Save-WordDocument $WordDocument