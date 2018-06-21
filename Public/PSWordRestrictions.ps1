function Add-WordProtection {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [EditRestrictions] $EditRestrictions,
        [string] $Password
    )
    if ($Password -eq $null) {
        $WordDocument.AddProtection($EditRestrictions)
    } else {
        $WordDocument.AddPasswordProtection($EditRestrictions, $Password)
    }
}