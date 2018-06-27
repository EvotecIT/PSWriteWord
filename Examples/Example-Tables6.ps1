### prepare data
$hash = [ordered] @{}
$hash.add("HQ-1", "5.54.546")
$hash.add("EUR-1", "6.0.0.1")
$hash.add("HQ-2", "5.6")
$hash.add("EUR-2", "6.1.5")
$hash.add("EUR-3", "6.2")

$hash1 = @{}
$hash1.add("HQ-1", "5.54.546")
$hash1.add("EUR-1", "6.0.0.1")
$hash1.add("HQ-2", "5.6")
$hash1.add("EUR-2", "6.1.5")
$hash1.add("EUR-3", "6.2")

$hash2 = @{}
$hash2.add("HQ-1", "5.54.546")

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
Add-WordTable -WordDocument $WordDocument -DataTable $hash -Design ColorfulList #-Verbose
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $hash -Design ColorfulGrid -Columns 'My Name', 'My Value'
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $hash1 -Design ColorfulGrid -Columns 'My Name', 'My Value' -AutoFit Window
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $hash2 -Design ColorfulGrid -Columns 'My Name', 'My Value'
Add-WordParagraph -WordDocument $WordDocument

Add-WordTable -WordDocument $WordDocument -DataTable $myArray -Design ColorfulList #-Verbose

Save-WordDocument $WordDocument -Language 'en-US'
Invoke-Item $FilePath