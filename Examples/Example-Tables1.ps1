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

Add-Section -WordDocument $WordDocument -PageBreak
$p = $WordDocument.InsertParagraph("Active Directory Root DSE").FontSize(15)
$p = $WordDocument.InsertParagraph("")
Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.RootDSE -Design LightShading -Verbose
Add-Section -WordDocument $WordDocument -PageBreak
$p = $WordDocument.InsertParagraph("Active Directory Forest Information").FontSize(15)
$p = $WordDocument.InsertParagraph("")
Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.ForestInformation -Design LightShading -Verbose
Add-Section -WordDocument $WordDocument -PageBreak
$p = $WordDocument.InsertParagraph("Active Directory Domain Information").FontSize(15)
$p = $WordDocument.InsertParagraph("")
Add-WordTable -WordDocument $WordDocument -Table $ADSnapshot.DomainInformation -Design LightShading -Verbose

Save-WordDocument $WordDocument