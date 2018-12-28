Import-Module PSWriteWord -Force

$FilePath1 = "$($Env:USERPROFILE)\Desktop\EmptyDocument.docx"
$FilePath2 = "$($Env:USERPROFILE)\Desktop\PSWriteWord-Example-PageOrientation1.docx"
$FileOutput = "$($Env:USERPROFILE)\Desktop\Output.docx"

Merge-WordDocument -FilePath1 $FilePath1 -FilePath2 $FilePath2 -FileOutput $FileOutput -OpenDocument -Supress $true