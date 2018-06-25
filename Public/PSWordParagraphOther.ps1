<#
ParentContainer           : None
ListItemType              : Bulleted
Pictures                  : {}
Hyperlinks                : {}
StyleName                 : Normal
DocumentProperties        : {}
Direction                 : RightToLeft
IndentationFirstLine      : 0
IndentationHanging        : 0
IndentationBefore         : 0
IndentationAfter          : 0
Alignment                 : left
Text                      : Like me like i do
MagicText                 : {Xceed.Words.NET.FormattedText}
FollowingTable            :
LineSpacing               : 22
LineSpacingBefore         : 15
LineSpacingAfter          : 10
ParagraphNumberProperties :
IsListItem                : False
IndentLevel               :
IsKeepWithNext            : False
Xml                       : <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                              <w:pPr>
                                <w:spacing w:before="300" />
                                <w:bidi />
                                <w:ind />
                              </w:pPr>
                              <w:r>
                                <w:rPr>
                                  <w:lang w:val="pl-PL" />
                                  <w:sz w:val="42"></w:sz>
                                  <w:szCs w:val="42"></w:szCs>
                                  <w:b></w:b>
                                </w:rPr>
                                <w:t>Like me like i do</w:t>
                              </w:r>
                            </w:p>
PackagePart               : System.IO.Packaging.ZipPackagePart
#>


#$Formatting = [Xceed.Words.NET.Formatting]
#$Formatting.Attributes.

Function Add-WordParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [bool] $Supress = $true
    )
    $Data = $WordDocument.InsertParagraph()
    if ($Supress -eq $true) { return } else { return $Data }
}

Function Set-WordParagraph {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Alignment] $Alignment,
        [Direction] $Direction,
        [string] $Language
    )
    if ($Paragraph -ne $null) {
        Write-Verbose "Set-WordParagraph - Paragraph is not null"
        if ($Alignment -ne $null) {
            Write-Verbose "Set-WordParagraph - Setting Alignment to $Alignment"
            $Paragraph.Alignment = $Alignment
        }
        if ($Direction -ne $null) {
            Write-Verbose "Set-WordParagraph - Setting Direction to $Direction"
            $Paragraph.Direction = $Direction
        }
        if ($Language -ne $null) {
            $Culture = [System.Globalization.CultureInfo]::GetCultureInfo($Language)
            $Paragraph = $Paragraph.Culture($Culture)
        }
    }
}
function Get-WordParagraphForList {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        $ListID
    )
    $IDs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        #Write-Color "testtting " -Color Yellow
        if ($p.ParagraphNumberProperties -ne $null) {
            $ListNumber = $p.ParagraphNumberProperties.LastNode.LastAttribute.Value
            if ($ListNumber -eq $ListID) {
                $IDs += $p
            }
        }
        #

        #$p.StyleName = 'Heading1'
    }
    return $Ids
}
function Get-WordParagraphs {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument
    )
    $IDs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        $p
    }
    return $Ids
}