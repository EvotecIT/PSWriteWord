Import-Module PSWriteWord #-Force
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

#Clear-Host

$WordDocument = New-WordDocument $FilePath
Add-WordSection -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -Text "Active Directory Root DSE" -FontSize 15
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.RootDSE -Design LightShading  #-Verbose
Add-WordSection -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -Text "Active Directory Forest Information" -FontSize 15
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.ForestInformation -Design LightShading #-Verbose
Add-WordSection -WordDocument $WordDocument -PageBreak
Add-WordText -WordDocument $WordDocument -Text "Active Directory Domain Information" -FontSize 15
Add-WordParagraph -WordDocument $WordDocument
Add-WordTable -WordDocument $WordDocument -DataTable $ADSnapshot.DomainInformation -Design LightShading #-Verbose

Save-WordDocument $WordDocument

### Start Word with file
Invoke-Item $FilePath