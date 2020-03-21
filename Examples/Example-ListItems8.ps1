
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems8.docx"

### define new document
$WordDocument = New-WordDocument $FilePath -Verbose

New-WordList -WordDocument $WordDocument {
    New-WordListItem -Level 0 -Text 'Test'
    New-WordListItem -Level 1 -Text 'Test1'
    New-WordListItem -Level 0 -Text 'Test2'
}

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument