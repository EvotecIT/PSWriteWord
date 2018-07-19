$PSVersionTable.PSVersion

$ModuleVersion = (Get-Content -Raw .\PSWriteWord.psd1)  | Invoke-Expression | ForEach-Object ModuleVersion

$dest = "PSWriteWord-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss")
Compress-Archive -Path . -DestinationPath .\$dest

if ((Get-Module -ListAvailable pester) -eq $null) {
    Install-Module -Name Pester -Repository PSGallery -Force
}

$result = Invoke-Pester -Script $PSScriptRoot\Tests -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}