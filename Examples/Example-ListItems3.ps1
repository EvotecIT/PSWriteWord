
Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems3.docx"

$myitems = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $myitems -Supress $false -Verbose

#Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true

#Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
#Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $true

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
#Invoke-Item $FilePath