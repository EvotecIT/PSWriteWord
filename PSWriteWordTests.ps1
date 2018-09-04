$PSVersionTable.PSVersion

$ModuleVersion = (Get-Content -Raw $PSScriptRoot\PSWriteWord.psd1)  | Invoke-Expression | ForEach-Object ModuleVersion

#$dest = "Builds\PSWriteWord-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss")
#Compress-Archive -Path . -DestinationPath .\$dest

if ((Get-Module -ListAvailable pester) -eq $null) {
    Write-Warning 'PSWriteWordTests - Downloading Pester from PSGallery'
    Install-Module -Name Pester -Repository PSGallery -Force -SkipPublisherCheck
}
if ((Get-Module -ListAvailable PSSharedGoods) -eq $null) {
    Write-Warning 'PSWriteWordTests - Downloading PSSharedGoods from PSGallery'
    Install-Module -Name PSSharedGoods -Repository PSGallery -Force
}

$result = Invoke-Pester -Script $PSScriptRoot\Tests -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}