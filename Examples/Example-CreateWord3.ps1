Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord3.docx"

$WordDocument = New-WordDocument $FilePath
$WordDocument.InsertParagraph("This is a text").FontSize("10") | Out-Null
$WordDocument.InsertParagraph("Like me like i do").FontSize("21") | Out-Null
$p = $WordDocument.InsertParagraph("Process").FontSize("15")
$p.StyleName = [HeadingType]::Heading3
Save-WordDocument $WordDocument