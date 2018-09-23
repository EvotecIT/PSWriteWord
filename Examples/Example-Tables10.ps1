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
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables10.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 50, 5, 25, 20 -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 50, 5, 25, 20 -AutoFit Window -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 10, 25, 50, 15 -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 10, 25, 50, 15 -AutoFit Contents -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $false -ColumnWidth 5, 5, 60, 30 -Verbose -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $false -ColumnWidth 5, 5, 60, 30 -AutoFit Window -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 5, 5, 60, 30 -Verbose -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulGrid -Percentage $true -ColumnWidth 5, 5, 60, 30 -AutoFit Window -Verbose -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True


Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument