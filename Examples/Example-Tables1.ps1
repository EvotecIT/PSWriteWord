Import-Module PSWriteWord -Force
Import-Module ActiveDirectory

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables1.docx"

# Get basic AD information
$ADSnapshot = @{}
$ADSnapshot.RootDSE = $(
    $Info = Get-ADRootDSE
    $Info | Select-Object `
    @{label = 'Configuration Naming Context'; expression = { $_.configurationNamingContext }},
    defaultNamingContext, dnsHostName, domainControllerFunctionality, domainFunctionality,
    forestFunctionality, supportedLDAPPolicies, subschemaSubentry, supportedLDAPVersion, supportedSASLMechanisms
)
$ADSnapshot.ForestInformation = $(
    Get-ADForest | Select-Object DomainNamingMaster, Domains, ForestMode, Sites
)
$ADSnapshot.DomainInformation = $(Get-ADDomain)
# Get basic Ad information end

Clear-Host

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text "Active Directory Root DSE" `
    -FontSize 15 -CapsStyle smallCaps -Alignment both -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.RootDSE `
    -Design LightShading -Bold $true -Color Blue -Supress $True -PivotRows
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Active Directory ", 'Domain', ' Forest Information' `
    -FontSize 12, 12, 12 -StrikeThrough none, strike, none -Alignment center -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.ForestInformation `
    -Design LightShading -Italic $true, $false -Bold $true, $false -ContinueFormatting -Supress $True -Verbose #-PivotRows -AutoFit Window
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Active Directory Domain Information" `
    -FontSize 15 -Color Green -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.DomainInformation `
    -Design LightShading -Supress $True -PivotRows
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True

### Start Word with file
Invoke-Item $FilePath