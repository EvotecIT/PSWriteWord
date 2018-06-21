function Add-WordProtection {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [EditRestrictions] $EditRestrictions
    )
    $WordDocument.AddProtection($EditRestrictions)
}