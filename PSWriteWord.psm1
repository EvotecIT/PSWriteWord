#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
$Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\*.dll -ErrorAction SilentlyContinue )

#Dot source the files
Foreach ($import in @($Public + $Private)) {
    Try {
        . $import.fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}
Foreach ($import in @($Assembly)) {
    Try {
        Add-Type -Path $import.fullname
    } Catch {
        Write-Error -Message "Failed to import DLL $($import.fullname): $_"
    }
}

Export-ModuleMember -Function '*'