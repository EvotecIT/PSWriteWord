Import-Module .\PSWriteWord.psd1 -Force

$FilePath = "$PSScriptRoot\Output\PSWriteWord-Example-Tables5.docx"

$myitems = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)

$WordDocument = New-WordDocument $FilePath
$WordDocument.PackagePart.Package.PackageProperties.Title = "Test"

Add-WordTable -WordDocument $WordDocument -DataTable $myitems -Design ColorfulList -Supress $True

Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myitems1 -Design ColorfulList -Supress $True #-Verbose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True

### Start Word with file
Invoke-Item $FilePath