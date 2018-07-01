$obj = New-Object System.Object
$obj | Add-Member -type NoteProperty -name Name -Value "Ryan_PC"
$obj | Add-Member -type NoteProperty -name Manufacturer -Value "Dell"
$obj | Add-Member -type NoteProperty -name ProcessorSpeed -Value "3 Ghz"
$obj | Add-Member -type NoteProperty -name Memory -Value "6 GB"

$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -name Memory -Value "4 GB"

$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -name Memory -Value "2.5 GB"

$myArray = @($obj, $myobject2, $myObject3)
### prepare data end

Import-Module PSWriteWord #-Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables6.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 50, 5, 25, 20
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 50, 5, 25, 20 -AutoFit Window
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 10, 25, 50, 15
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 10, 25, 50, 15 -AutoFit Contents
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $false -ColummnWidth 5, 5, 60, 30 -Verbose
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $false -ColummnWidth 5, 5, 60, 30 -AutoFit Window -Verbose
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 5, 5, 60, 30 -Verbose
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColummnWidth 5, 5, 60, 30 -AutoFit Window -Verbose
Add-WordParagraph -WordDocument $WordDocument


Save-WordDocument $WordDocument -Language 'en-US'
Invoke-Item $FilePath