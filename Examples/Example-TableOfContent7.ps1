Import-Module PSWriteWord #-Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent7.docx"
$ListOfHeaders = @('This is 1st section', 'This is 2nd section', 'This is 3rd section', 'This is 4th section', 'This is 5th section')

$WordDocument = New-WordDocument -FilePath $FilePath
$WordDocument | Add-WordToc -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1 -Supress $true

foreach ($Section in $ListOfHeaders) {
    $Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
    $Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange
}
$Paragraph = $WordDocument | Add-WordTocItem -Text 'Adding another one' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
$Paragraph = $WordDocument | Add-WordText -Text 'This is my test - outside of loop. Added after TOC Item.' -Color Red

$WordDocument | Save-WordDocument -Language 'en-US' -Supress $True
### Start Word with file
Invoke-Item $FilePath