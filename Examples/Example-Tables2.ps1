Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables2.docx"

Clear-Host
$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text "This is a text, after which we add section break, followed by table" -FontSize 20

Add-WordSection -WordDocument $WordDocument -PageBreak
$Object1 = Get-Process #| Select-Object ProcessName, Site, StartTime
Add-WordTable -WordDocument $WordDocument -Table $Object1 -Design 'ColorfulList' #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Then we do another pagebreak, and add another table" -FontSize 20
Add-WordSection -WordDocument $WordDocument -PageBreak
$Object2 = Get-PSDrive
Add-WordTable -WordDocument $WordDocument -Table $Object2 -Design "LightShading" #-Verbose

Add-WordText -WordDocument $WordDocument -Text "Then we do another pagebreak, and add another table" -FontSize 20
Add-WordSection -WordDocument $WordDocument -PageBreak
$Object3 = $Object1 | Select-Object ProcessName, Site, StartTime
Add-WordTable -WordDocument $WordDocument -Table $Object3 -Design 'ColorfulList' #-Verbose


Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath