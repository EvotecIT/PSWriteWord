#using namespace Xceed.Words.NET
#using namespace Xceed.Document.NET

#Get public and private function definition files.
$Public = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue -Recurse )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue -Recurse )
if ($PSEdition -eq 'Core') {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Core\*.dll -ErrorAction SilentlyContinue )
} else {
    $Assembly = @( Get-ChildItem -Path $PSScriptRoot\Lib\Default\*.dll -ErrorAction SilentlyContinue )
}

Foreach ($import in @($Assembly)) {
    Try {
        Add-Type -Path $import.fullname
    } Catch {
        Write-Error -Message "Failed to import DLL $($import.fullname): $_"
    }
}

#Dot source the files
Foreach ($import in @($Public + $Private)) {
    Try {
        . $import.fullname
    } Catch {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

Export-ModuleMember -Function '*' -Alias '*'