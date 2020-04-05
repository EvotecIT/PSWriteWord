Import-Module $PSScriptRoot\..\PSWriteWord.psd1 -Force

$Object1 = [pscustomobject]@{
    Test1 = 1
    Test2 = 2
    Test3 = 3
}

# PSWriteWord
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables15.docx"
$WordDocument = New-WordDocument $FilePath
Add-WordTable -WordDocument $WordDocument -DataTable $Object1 -Design 'ColorfulList' -Supress $true -OverwriteTitle 'Test' -AutoFit Window -Transpose
Save-WordDocument $WordDocument -Supress $true #-OpenDocument

# Documentimo
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables15-Documentimo.docx"
Documentimo {
    DocumentimoTable -DataTable $Object1 -Design 'ColorfulList' -OverwriteTitle 'Test' -AutoFit Window -Transpose
} -FilePath $FilePath -Open