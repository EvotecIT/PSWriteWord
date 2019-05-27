Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables9.docx"

$WordDocument = New-WordDocument $FilePath

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services 1'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services 2'
$InvoiceEntry2.Amount = '$300'

$InvoiceEntry3 = @{}
$InvoiceEntry3.Description = 'IT Services 3'
$InvoiceEntry3.Amount = '$288'

$InvoiceEntry4 = @{}
$InvoiceEntry4.Description = 'IT Services 4'
$InvoiceEntry4.Amount = '$301'

$InvoiceEntry5 = @{}
$InvoiceEntry5.Description = 'IT Services 5'
$InvoiceEntry5.Amount = '$299'

$InvoiceData = @(
    $InvoiceEntry1
    $InvoiceEntry2
    $InvoiceEntry3
    $InvoiceEntry4
    $InvoiceEntry5
)

Add-WordText -WordDocument $WordDocument -Text "Invoice Data" -Alignment center -FontSize 15 -UnderlineColor Blue -UnderlineStyle doubleLine -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -AutoFit Window -Color Blue, Green, Red -FontSize 15, 10, 8 -Bold $true, $false, $false -FontFamily 'Arial', 'Tahoma' -Supress $True -Transpose

Add-WordText -WordDocument $WordDocument -Text "Invoice Data" -Alignment center -FontSize 15 -UnderlineColor Blue -UnderlineStyle doubleLine -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -AutoFit Window -Color Blue, Green, Red -FontSize 15, 10, 8 -Bold $true, $false, $false -FontFamily 'Arial', 'Tahoma' -ContinueFormatting -Supress $True -Transpose

Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Invoice Data with different formatting" -Alignment center -FontSize 15 -UnderlineColor Blue -UnderlineStyle doubleLine -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -AutoFit Window -Color Blue, Green, Red -FontSize 15, 10 -Bold $true, $true, $false -FontFamily 'Tahoma' -ContinueFormatting -Supress $True -Transpose

Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'Notice how ', 'Continue Formatting', ' switch takes over formatting for', `
    ' font family ', ',', 'font size', ' and ', `
    'bold', '. It takes over the last entry for each formatting and continues it. That way you can set ', 'FontFamily', `
    ' to ', 'Tahoma', ' for whole table and still have different row colors if needed.' `
    -Color Black, Blue, Black, Blue, Black, Blue, Black, Blue `
    -Bold $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false, $true -Supress $True

Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordText -WordDocument $WordDocument -Text 'Notice how ', 'Continue Formatting', ' switch takes over formatting for', `
    ' font family ', ',', 'font size', ' and ', `
    'bold', '. It takes over the last entry for each formatting and continues it. That way you can set ', 'FontFamily', `
    ' to ', 'Tahoma', ' for whole table and still have different row colors if needed.' `
    -Color Black, Blue, Black, Blue, Black, Blue, Black, Blue `
    -Bold $false, $false, $false, $false, $false, $false, $false, $false, $false, $true, $false, $true -Supress $True -ContinueFormatting

Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Invoice Data with different formatting" -Alignment center -FontSize 15 -UnderlineColor Blue -UnderlineStyle doubleLine -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -AutoFit Window -FontFamily 'Tahoma' -FontSize 10, 9 -ContinueFormatting -Supress $True -Transpose

Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Lots of different formatting" -Alignment center -FontSize 15 -UnderlineColor Red -UnderlineStyle dotDash -CapsStyle smallCaps -Supress $True
Add-WordTable -WordDocument $WordDocument -AutoFit Window -DataTable $InvoiceData -FontSize 10 -FontFamily Tahoma -ContinueFormatting -Design ColorfulList -StrikeThrough none, doubleStrike, none -Color Black, Black, Red, Black -Supress $True -Transpose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument