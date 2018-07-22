Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables2.docx"

Clear-Host
$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text "This is a text, after which we add section break, followed by table" -FontSize 20 -Supress $true

Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true
$Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime
Add-WordTable -WordDocument $WordDocument -DataTable $Object1 -Design 'ColorfulList' -Supress $true #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Then we do another pagebreak, and add another table" -FontSize 20 -Supress $true
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true
$Object2 = Get-PSDrive
Add-WordTable -WordDocument $WordDocument -DataTable $Object2 -Design "LightShading" -Supress $true #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Then we do another pagebreak, and add another table" -FontSize 20 -Supress $true
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true
$Object3 = $Object1 | Select-Object ProcessName, Id, StartTime
Add-WordTable -WordDocument $WordDocument -DataTable $Object3 -Design 'ColorfulList' -Supress $true #-Verbose

Save-WordDocument $WordDocument -Supress $true

### Start Word with file
Invoke-Item $FilePath