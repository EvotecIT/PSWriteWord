Import-Module PSWriteWord -Force
Clear-Host
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent6.docx"
$ListOfHeaders = @('This is 1st section', 'This is 2nd section', 'This is 3rd section', 'This is 4th section', 'This is 5th section')

$ListsToConvertToheaders = @()

$WordDocument = New-WordDocument -FilePath $FilePath
Add-WordToc -WordDocument $WordDocument -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1 -Supress $True

#Write-Color "lists count before ", $WordDocument.Lists.Count -Color White, Yellow
#Write-Color "lists elements before ", $WordDocument.Lists[0].Items.Count -Color White, Yellow

foreach ($Section in $ListOfHeaders) {
    $List = New-WordListItem -WordDocument $WordDocument -List $null -Text $Section -ListItemType Numbered -ContinueNumbering $true
    $List = Add-WordListItem -WordDocument $WordDocument -List $List
    $Paragraph = Add-WordText -WordDocument $WordDocument -Text "Random Text in section $section" -Supress $true
    $ListsToConvertToheaders += $List
}

foreach ($List in $ListsToConvertToheaders) {
    $Headings = Convert-ListToHeadings -WordDocument $WordDocument -List $List
}

$List = New-WordListItem -WordDocument $WordDocument -List $null -Text 'Adding one more thing' -ListItemType Numbered -ContinueNumbering $true -ListLevel 1
$List = Add-WordListItem -WordDocument $WordDocument -List $List
Convert-ListToHeadings -WordDocument $WordDocument -List $List -Supress $True

#Write-Color "lists count after ", $WordDocument.Lists.Count -Color White, Yellow
#Write-Color "lists elements after ", $WordDocument.Lists[0].Items.Count -Color White, Yellow

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True
### Start Word with file
Invoke-Item $FilePath