Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems5.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
#Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $True -Verbose

$List = $null
for ($i = 0; $i -le 8; $i++) {
    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel $i -ListItemType Bulleted -ListValue "Test $i"
    $Paragraph = Get-WordListItemParagraph -List $List -LastItem

    Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text ' Added text' -Color Blue
    #$Paragraph
    #break
}
Add-WordListItem -WordDocument $WordDocument -List $List

$List = $null
for ($i = 0; $i -le 8; $i++) {
    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel $i -ListItemType Numbered -ListValue "Test $i"
}
Add-WordListItem -WordDocument $WordDocument -List $List

$List3 = New-WordListItem -WordDocument $WordDocument -List $null -ListLevel 0 -ListItemType Numbered -ListValue 'Test 1'
$List3 = New-WordListItem -WordDocument $WordDocument -List $List3 -ListLevel 1 -ListItemType Numbered -ListValue 'Test 1'
$List3 = New-WordListItem -WordDocument $WordDocument -List $List3 -ListLevel 2 -ListItemType Bulleted -ListValue 'Test 1'

$List2 = New-WordListItem -WordDocument $WordDocument -List $null -ListLevel 0 -ListItemType Numbered -ListValue 'Test 2'
$List2 = New-WordListItem -WordDocument $WordDocument -List $List2 -ListLevel 1 -ListItemType Numbered -ListValue 'Test 2'
$List2 = New-WordListItem -WordDocument $WordDocument -List $List2 -ListLevel 2 -ListItemType Numbered -ListValue 'Test 2'

#$WordDocument.AddList

#$List = $WordDocument.AddList($ListValue, $ListLevel, $ListItemType)
#$List = $WordDocument.AddListItem($List, $ListValue, $ListLevel)

Add-WordListItem -WordDocument $WordDocument -List $List2

Add-WordListItem -WordDocument $WordDocument -List $List3
#Add-WordSection -WordDocument $WordDocument -PageBreak

#Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
#Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath