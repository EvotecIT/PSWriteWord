function Add-WordTOC {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [alias ('BeforeParagraph')][parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $Title = 'Table of contents',
        [TableOfContentsSwitches[]] $Switches = @(),
        [alias ('Heading', 'HeadingType')][HeadingType] $HeaderStyle = [HeadingType]::Heading1,
        [int] $MaxIncludeLevel = 3,
        [int] $RightTabPos = $null,
        [bool] $Supress = $false
    )
    if ($WordDocument -ne $null) {
        #public TableOfContents InsertTableOfContents( string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
        #public TableOfContents InsertTableOfContents( Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )

        $TableOfContentSwitch = 0
        foreach ($S in $switches) {
            $TableOfContentSwitch += $s -As [Int]
        }
        Write-Verbose $TableOfContentSwitch


        #if ($Title -eq $null -or $Title -eq '') {
        #    $Toc = $WordDocument.InsertDefaultTableOfContents()
        #    if ($Supress -eq $false) { return $Toc } else { return }
        #}
        if ($Paragraph -eq $null) {
            $toc = $WordDocument.InsertTableOfContents($Title, $TableOfContentSwitch, $HeaderStyle, $MaxIncludeLevel, $RightTabPos)
        } else {
            $toc = $WordDocument.InsertTableOfContents($Paragraph, $Title, $TableOfContentSwitch, $HeaderStyle, $MaxIncludeLevel, $RightTabPos)
        }
    }
    if ($Supress -eq $false) { return $Toc } else { return }
}

function Add-WordTocItem {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [alias('Level')] [ValidateRange(0, 8)] [int] $ListLevel,
        [alias('ListType')][ListItemType] $ListItemType = [ListItemType]::Bulleted,
        [alias('Value', 'ListValue')]$Text,
        [alias ("HT")] [HeadingType] $HeadingType = [HeadingType]::Heading1,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $true,
        [bool]$Supress = $false
    )
    $List = New-WordListItem -WordDocument $WordDocument -List $null -Text $Text -ListItemType $ListItemType -ContinueNumbering $ContinueNumbering -ListLevel $ListLevel -StartNumber $StartNumber -TrackChanges $TrackChanges
    $List = Add-WordListItem -WordDocument $WordDocument -List $List
    $Paragraph = Convert-ListToHeadings -WordDocument $WordDocument -List $List -HeadingType $HeadingType
    if ($Supress) { return } else { return $Paragraph }
}