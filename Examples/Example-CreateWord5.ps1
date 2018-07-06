Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord5.docx"

### define new document
$WordDocument = New-WordDocument $FilePath
### add 3 paragraphs, using pipeline for $worddocument
$WordDocument | Add-WordText -Text 'This is a text' -FontSize 10
$WordDocument | Add-WordText -Text 'This is a text font size 21' -FontSize 21
$WordDocument | Add-WordText -Text 'This is a text font size 15' -FontSize 15

$Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $false
$Paragraph = Add-WordTabStopPosition -Paragraph $Paragraph -TabStopPositionLeader dot -HorizontalPosition 216 -Alignment center -Supress $false
$Paragraph = Add-WordTabStopPosition -Paragraph $Paragraph -TabStopPositionLeader dot -HorizontalPosition 432 -Alignment right -Supress $false
$Paragraph = Add-WordText -Paragraph $Paragraph -Text "Tab stop position on Left`t Middle `t and Right" -FontSize 15 -Supress $false

### adds green color to paragraph above
Set-WordText -Paragraph $Paragraph -Color Green -FontSize 30 -Verbose

### adds another empty paragraph
$Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $false

### adds 3 texts to to an empty paragraph from above
Add-WordText -Paragraph $Paragraph -Color Green -FontSize 30 -bold $null, $null, $true -Text 'Font size 30', ' not font size 30', ' not font size 30 but bold'

### Save document
$WordDocument |Save-WordDocument -Language 'en-US'

### Start Word with file
Invoke-Item $FilePath