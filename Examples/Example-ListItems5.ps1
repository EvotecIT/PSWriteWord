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
    Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text ' Added text' -Color Blue -AppendToExistingParagraph -Supress $True
}
Add-WordListItem -WordDocument $WordDocument -List $List -Supress $True

$List = $null
for ($i = 0; $i -le 8; $i++) {
    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel $i -ListItemType Numbered -ListValue "Test $i"
    $Paragraph = Get-WordListItemParagraph -List $List -LastItem
    Set-WordText -Paragraph $Paragraph -Color Red -FontSize 10 -FontFamily Tahoma -Supress $True
}
Add-WordListItem -WordDocument $WordDocument -List $List -Supress $True


$ListColors = $null
for ($i = 0; $i -le 8; $i++) {
    $ListColors = New-WordListItem -WordDocument $WordDocument -List $ListColors -ListLevel $i -ListItemType Bulleted -ListValue "Testing Colors $i"
    #$Paragraph = Get-WordListItemParagraph -List $List -LastItem
    #Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text ' Added text' -Color Blue -AppendToExistingParagraph -Supress $True
}
$ParagraphFromList0 = Get-WordListItemParagraph -List $ListColors -Item 1 #-Verbose
$ParagraphFromList1 = Get-WordListItemParagraph -List $ListColors -Item 3 #-Verbose
$ParagraphFromList2 = Get-WordListItemParagraph -List $ListColors -Item 5 #-Verbose
$ParagraphFromList3 = Get-WordListItemParagraph -List $ListColors -Item 7 #-Verbose
Set-WordText -Paragraph $ParagraphFromList1 -Color Red -FontFamily Calibri -Supress $True
Set-WordText -Paragraph $ParagraphFromList2 -Text 'This will add text' -Color Blue -FontFamily Calibri -Supress $True -Verbose
Set-WordText -Paragraph $ParagraphFromList3 -Color Orange -FontFamily Calibri -Supress $True -Verbose
Set-WordText -Paragraph $ParagraphFromList0 -Text 'Replacing this text with color Orange' -FontSize 8 -ClearText -Color Orange -Verbose
Add-WordListItem -WordDocument $WordDocument -List $ListColors -Supress $True

$List = $null
for ($i = 0; $i -le 8; $i++) {
    $List = New-WordListItem -WordDocument $WordDocument -List $List -ListLevel $i -ListItemType Bulleted -ListValue "Test $i"
    #$Paragraph = Get-WordListItemParagraph -List $List -LastItem
    #Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph -Text ' Added text' -Color Blue -AppendToExistingParagraph -Supress $True
}
Add-WordListItem -WordDocument $WordDocument -List $List -Supress $True

$List2 = New-WordListItem -WordDocument $WordDocument -List $null -ListLevel 0 -ListItemType Numbered -ListValue 'Test 2'
$List2 = New-WordListItem -WordDocument $WordDocument -List $List2 -ListLevel 1 -ListItemType Numbered -ListValue 'Test 2'
$List2 = New-WordListItem -WordDocument $WordDocument -List $List2 -ListLevel 2 -ListItemType Numbered -ListValue 'Test 2'

$List3 = New-WordListItem -WordDocument $WordDocument -List $null -ListLevel 0 -ListItemType Numbered -ListValue 'Test 1'
$List3 = New-WordListItem -WordDocument $WordDocument -List $List3 -ListLevel 1 -ListItemType Numbered -ListValue 'Test 1'
$List3 = New-WordListItem -WordDocument $WordDocument -List $List3 -ListLevel 2 -ListItemType Bulleted -ListValue 'Test 1'

#$WordDocument.AddList

#$List = $WordDocument.AddList($ListValue, $ListLevel, $ListItemType)
#$List = $WordDocument.AddListItem($List, $ListValue, $ListLevel)

Add-WordListItem -WordDocument $WordDocument -List $List2 -Supress $True
Add-WordListItem -WordDocument $WordDocument -List $List3 -Supress $True
#Add-WordSection -WordDocument $WordDocument -PageBreak

#Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
#Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true
Invoke-Item $FilePath