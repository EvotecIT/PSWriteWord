Import-Module .\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems2.docx"

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services'
$InvoiceEntry2.Amount = '$200'

$ListOfItems = @()
$ListOfItems += $InvoiceEntry1
$ListOfItems += $InvoiceEntry2


<#
function Test {
    param(
        [Array] $Test
    )

    $Test[0].GetType()

    if ($Test[0] -is [System.Collections.IDictionary]) {
        $true
    }
}

Test -Test $InvoiceEntry1
Test -Test $InvoiceEntry2
Test -Test $ListOfItems
#>


$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be bulleted list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True -Verbose
Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems -Supress $true -Verbose

Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $true

Add-WordText -WordDocument $WordDocument -Text 'This is text after which will be numbered list' -FontSize 15 -UnderlineStyle singleLine -HeadingType Heading2 -Supress $True
Add-WordList -WordDocument $WordDocument -ListType Numbered -ListData $InvoiceEntry1 -Supress $true -Verbose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument