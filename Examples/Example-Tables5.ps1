Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables5.docx"

$myitems = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)

$WordDocument = New-WordDocument $FilePath

Add-WordTable -WordDocument $WordDocument -DataTable $myitems -Design ColorfulList #-Verbose

Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myitems1 -Design ColorfulList #-Verbose

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath