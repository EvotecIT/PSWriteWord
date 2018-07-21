function Add-WordTOC {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $BeforeParagraph,
        [string] $Title = 'Table of contents',
        [TableOfContentsSwitches[]] $Switches = @(),
        [HeadingType] $HeaderStyle = [HeadingType]::Heading1,
        [int] $MaxIncludeLevel = 3,
        [int] $RightTabPos = $null,
        [bool] $Supress = $false
    )

    #public TableOfContents InsertTableOfContents( string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    #public TableOfContents InsertTableOfContents( Paragraph reference, string title, TableOfContentsSwitches switches, string headerStyle = null, int maxIncludeLevel = 3, int? rightTabPos = null )
    $TableOfContentSwitch = 0;
    foreach ($S in $switches) {
        $TableOfContentSwitch += $s -As [Int]
    }
    Write-Verbose $TableOfContentSwitch

    #if ($Title -eq $null -or $Title -eq '') {
    #    $Toc = $WordDocument.InsertDefaultTableOfContents()
    #    if ($Supress -eq $false) { return $Toc } else { return }
    #}
    if ($BeforeParagraph -eq $null) {
        $toc = $WordDocument.InsertTableOfContents($Title, $TableOfContentSwitch, $HeaderStyle, $MaxIncludeLevel, $RightTabPos)
    } else {
        $toc = $WordDocument.InsertTableOfContents($BeforeParagraph, $Title, $TableOfContentSwitch, $HeaderStyle, $MaxIncludeLevel, $RightTabPos)
    }

    if ($Supress -eq $false) { return $Toc } else { return }

}