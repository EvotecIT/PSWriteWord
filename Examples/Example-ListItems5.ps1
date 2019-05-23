Import-Module .\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems5.docx"

$WordDocument = New-WordDocument $FilePath


<#
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
}
$ParagraphFromList0 = Get-WordListItemParagraph -List $ListColors -Item 1 #-Verbose
$ParagraphFromList1 = Get-WordListItemParagraph -List $ListColors -Item 3 #-Verbose
$ParagraphFromList2 = Get-WordListItemParagraph -List $ListColors -Item 5 #-Verbose
$ParagraphFromList3 = Get-WordListItemParagraph -List $ListColors -Item 7 #-Verbose
Set-WordText -Paragraph $ParagraphFromList1 -Color Red -FontFamily Calibri -Supress $True
Set-WordText -Paragraph $ParagraphFromList2 -Text 'This will add text' -Append -Color Blue -FontFamily Calibri -Supress $True -Verbose
Set-WordText -Paragraph $ParagraphFromList3 -Color Orange -FontFamily Calibri -Supress $True -Verbose
Set-WordText -Paragraph $ParagraphFromList0 -Text 'Replacing this text with color Orange' -FontSize 8 -Color Orange -Verbose
Add-WordListItem -WordDocument $WordDocument -List $ListColors -Supress $True

#>


Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True


New-WordList -WordDocument $WordDocument -Type Bulleted {
    New-WordListItem -ListLevel 0 -ListValue 'Test 1'
    New-WordListItem -ListLevel 0 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 0 -ListValue 'Test 1'
}


Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True


New-WordList -WordDocument $WordDocument -Type Numbered {
    New-WordListItem -ListLevel 0 -ListValue 'Test 1' -StartNumber 3
    New-WordListItem -ListLevel 0 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 1 -ListValue 'Test 1'
    New-WordListItem -ListLevel 0 -ListValue 'Test 1'
}

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument