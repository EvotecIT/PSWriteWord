Import-Module PSWriteWord -Force
Import-Module ActiveDirectory

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-ListItems4.docx"

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

$WordDocument = New-WordDocument $FilePath

Add-WordText -WordDocument $WordDocument -Text "Active Directory Root DSE" -FontSize 15 -CapsStyle smallCaps -Alignment both -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordList -WordDocument $WordDocument -ListData $ADSnapshot.RootDSE -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Active Directory ", 'Domain', ' Forest Information' -FontSize 12, 12, 12 -StrikeThrough none, strike, none -Alignment center -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordList -WordDocument $WordDocument -DataTable $ADSnapshot.ForestInformation -Supress $True # -Italic $true, $false -Bold $true, $false -ContinueFormatting -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak  -Supress $True
Add-WordText -WordDocument $WordDocument -Text "Active Directory Domain Information"  -FontSize 15 -Color Green -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True
Add-WordList -WordDocument $WordDocument -DataTable $ADSnapshot.DomainInformation -Supress $true -Verbose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument