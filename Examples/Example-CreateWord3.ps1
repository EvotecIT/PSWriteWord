Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord3.docx"

$WordDocument = New-WordDocument $FilePath
#$WordDocument.InsertParagraph("This is a text").FontSize("10").FontColor() | Out-Null
#$WordDocument.InsertParagraph("Like me like i do").FontSize("21") | Out-Null
#$WordDocument.InsertParagraph("Process").FontSize("15") | Out-Null
Add-WordText -WordDocument $WordDocument -Text 'This is text that has font size of 15', ' and this is font size of 10 ', ' while this will be 12.' -FontSize 15, 10, 12 -Color White, Yellow, White
Save-WordDocument $WordDocument