Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems1.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')
$OverrideLevels = @(0, 1, 2, 1, 3)
$OverrideLevelsPartially = @(0, 3)

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $True -Verbose

Add-WordParagraph -WordDocument $WordDocument -Supress $True # Empty Line

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list - levels defined' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $True -ListLevels $OverrideLevels -Verbose

Add-WordParagraph -WordDocument $WordDocument -Supress $True # Empty Line

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list - levels defined partially' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $True -ListLevels $OverrideLevelsPartially -Verbose

Add-WordSection -WordDocument $WordDocument -PageBreak

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
$List = Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $false
Set-WordList -List $List -FontSize 8 -FontFamily 'Tahoma' -Color Orange -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath