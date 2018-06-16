Import-Module PSWriteWord -Force

$FilePath = "C:\Users\pklys\Desktop\File2.docx"

$WordDocument = New-WordDocument $FilePath
$WordDocument.InsertParagraph("This is a text").FontSize("10") | Out-Null
$WordDocument.InsertParagraph("Like me like i do").FontSize("21") | Out-Null
$WordDocument.InsertParagraph("Process").FontSize("15") | Out-Null
Save-WordDocument $WordDocument