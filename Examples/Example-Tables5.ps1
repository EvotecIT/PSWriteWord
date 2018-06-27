Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables5.docx"

$myitems = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$WordDocument = New-WordDocument $FilePath

Add-WordTable -WordDocument $WordDocument -Table $myitems -Design 'ColorfulList' #-Verbose

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath