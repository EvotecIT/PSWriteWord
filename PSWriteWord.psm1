#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

#Dot source the files
Foreach ($import in @($Public + $Private)) {
    Try {
        . $import.fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

$AssemblyName = "$($PSScriptRoot)\Lib\Xceed.Words.NET.dll"
Add-Type -Path $AssemblyName
#Export-ModuleMember -Function $Public.Basename -Verbose
#Export-ModuleMember -Function * #-Verbose # 'Start-ADReporting', 'Get-KerberosLogonEvents', 'Get-GroupPolicyChanges', 'Get-EventLogClearedLogs'
Export-ModuleMember -Function '*'