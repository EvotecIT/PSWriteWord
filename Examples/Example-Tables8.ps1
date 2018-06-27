Import-Module PSWriteWord -Force

$FilePathDocumentWithTable = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables3.docx"
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables8.docx"

$WordDocumentTemplate = Get-WordDocument -FilePath $FilePathDocumentWithTable

$Tables = Get-WordTable -WordDocument $WordDocumentTemplate -ListTables
