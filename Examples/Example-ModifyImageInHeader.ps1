Import-Module PSWriteWord

$FilePath = "$PSScriptRoot\Templates\WordTemplate-ImageInHeader.docx"
$FilePathNew = "$PSScriptRoot\Output\WordTemplate-ImageInHeaderReplaced.docx"

$ReplaceImage = 'C:\Support\GitHub\PSWriteWord\Examples\Images\Logo-FakeCompany.png'

$WordDocument = Get-WordDocument -FilePath $FilePath

$HeaderFirst = Get-WordHeader -WordDocument $WordDocument -Type Odd
$HeaderFirst.Paragraphs[0]

Set-WordPicture -WordDocument $WordDocument -Paragraph $HeaderFirst.Paragraphs[0] -ImagePath $ReplaceImage

Save-WordDocument -WordDocument $WordDocument -OpenDocument -FilePath $FilePathNew