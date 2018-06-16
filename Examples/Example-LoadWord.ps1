Import-Module PSWriteWord -Force

$FilePath = "C:\Users\pklys\Desktop\File1.docx"

$WordDocument = Get-WordDocument -FilePath $FilePath
$WordDocument.Paragraphs
