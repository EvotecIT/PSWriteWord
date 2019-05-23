function Add-WordTocItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [alias('Level')] [ValidateRange(0, 8)] [int] $ListLevel,
        [alias('ListType')][Xceed.Words.NET.ListItemType] $ListItemType = [Xceed.Words.NET.ListItemType]::Bulleted,
        [alias('Value', 'ListValue')]$Text,
        [alias ("HT")] [Xceed.Words.NET.HeadingType] $HeadingType = [Xceed.Words.NET.HeadingType]::Heading1,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $true,
        [bool]$Supress = $false
    )
    $List = New-WordListItemInternal -WordDocument $WordDocument -List $null -Text $Text -ListItemType $ListItemType -ContinueNumbering $ContinueNumbering -ListLevel $ListLevel -StartNumber $StartNumber -TrackChanges $TrackChanges
    $List = Add-WordListItem -WordDocument $WordDocument -List $List
    $Paragraph = Convert-ListToHeadings -WordDocument $WordDocument -List $List -HeadingType $HeadingType
    if ($Supress) { return } else { return $Paragraph }
}

