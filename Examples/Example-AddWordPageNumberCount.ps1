Import-Module ..\PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddWordHeaderFooter.docx"

### define new document
$WordDocument = New-WordDocument $FilePath -Verbose

$Footer = Add-WordFooter -WordDocument $WordDocument -DifferentFirstPage $true -DifferentOddAndEvenPages $false
$Header = Add-WordHeader -WordDocument $WordDocument
### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
$Paragraph = Add-WordPageBreak -WordDocument $WordDocument -Verbose
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $false
$Paragraph | Add-WordPageBreak -InsertWhere BeforeSelf -Supress $True -Verbose

Add-WordPageCount -Header $Header -PageNumberFormat normal -TextBefore 'Page Nr ' -TextMiddle ' of ' -TextAfter '' -Alignment center -Supress $True
Add-WordPageCount -Footer $Footer -Type First -PageNumberFormat normal -Option PageNumberOnly -Supress $True
Add-WordPageCount -Footer $Footer -Type Odd -PageNumberFormat normal -Option Both -Alignment  right -TextMiddle ' of ' -Supress $True

# this is an alias to Add-WordPageCount
$FooterParagraphsWithChanges = Add-WordPageNumber -Footer $Footer -Type All -PageNumberFormat roman -Option PageNumberOnly -Alignment right -TextBefore 'Page Number ' -Supress $false
#$FooterParagraphsWithChanges[0]

# this basically takes only 1 paragraph on the first footer (odd, even footers have their own paragraphs)
# and adds page count only to first footer along with text
# it also centers it on the first page (leaves as is on the rest)
Add-WordPageCount -Paragraph $FooterParagraphsWithChanges[0] -PageNumberFormat roman -Alignment center -Option PageCountOnly -TextBefore ' of '


### Save document
Save-WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument