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

Function Set-WordParagraph {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Alignment] $Alignment
    )
    if ($Paragraph -ne $null) {
        if ($Alignment -ne $null) {
            $Paragraph.Alignment = $Alignment
        }
    }
}