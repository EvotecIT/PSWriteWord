Import-Module PSWriteWord -Force

$FilePath = "C:\Users\pklys\Desktop\File-TableOfContent.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

$toc = $WordDocument.InsertTableOfContents("Table of content", 1)
