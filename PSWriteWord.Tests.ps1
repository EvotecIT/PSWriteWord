$PSVersionTable.PSVersion

$ModuleName = (Get-ChildItem $PSScriptRoot\*.psd1).BaseName
$ModuleVersion = (Get-Content -Raw $PSScriptRoot\*.psd1)  | Invoke-Expression | ForEach-Object ModuleVersion

#$Dest = "Builds\$ModuleName-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss")
#Compress-Archive -Path . -DestinationPath .\$dest

if ((Get-Module -ListAvailable pester) -eq $null) {
    Write-Warning "$ModuleName - Downloading Pester from PSGallery"
    Install-Module -Name Pester -Repository PSGallery -Force -SkipPublisherCheck
}
if ((Get-Module -ListAvailable PSSharedGoods) -eq $null) {
    Write-Warning "$ModuleName - Downloading PSSharedGoods from PSGallery"
    Install-Module -Name PSSharedGoods -Repository PSGallery -Force
}

$result = Invoke-Pester -Script $PSScriptRoot\Tests -Verbose -EnableExit

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}