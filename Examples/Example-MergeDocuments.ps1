Import-Module PSWriteWord -Force

$FilePath1 = "$($Env:USERPROFILE)\Desktop\File1.docx"
$FilePath2 = "$($Env:USERPROFILE)\Desktop\File2.docx"
$FileOutput = "$($Env:USERPROFILE)\Desktop\Output.docx"

Merge-WordDocument -FilePath1 $FilePath1 -FilePath2 $FilePath2 -FileOutput $FileOutput -Supress