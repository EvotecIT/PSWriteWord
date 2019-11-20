Import-Module PSWriteWord -Force

$FilePath = "$PSScriptRoot\Templates\CV-Template.docx"
$FilePathOutput = "$PSScriptRoot\Output\CV-ReplacedText.docx"

$WordDocument = Get-WordDocument -FilePath $FilePath

foreach ($Paragraph in $WordDocument.Paragraphs) {
    $Paragraph.ReplaceText('image','picture')
}

Save-WordDocument -WordDocument $WordDocument -FilePath $FilePathOutput -OpenDocument -Supress $true