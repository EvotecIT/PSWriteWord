function Add-WordPageCount {
    [alias('Add-WordPageNumber')]
    param(
        [PageNumberFormat] $PageNumberFormat = [PageNumberFormat]::normal,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Footers] $Footer,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Headers] $Header,
        [Alignment] $Alignment,
        [ValidateSet('All', 'First', 'Even', 'Odd')][string] $Type = 'All',
        [ValidateSet('Both', 'PageCountOnly', 'PageNumberOnly')][string] $Option = 'Both',
        [string] $TextBefore,
        [string] $TextMiddle,
        [string] $TextAfter,
        [bool] $Supress
    )
    $Paragraphs = New-GenericList
    if ($Footer -or $Header -or $Paragraph) {
        if ($null -eq $Paragraph) {

            if ($Type -eq 'All') {
                $Types = 'First', 'Even', 'Odd'
                foreach ($T in $Types) {
                    if ($Footer) {
                        $Paragraphs.Add($Footer.$T.InsertParagraph())
                    }
                    if ($Header) {
                        $Paragraphs.Add($Header.$T.InsertParagraph())
                    }
                }
            } else {
                if ($Footer) {
                    $Paragraphs.Add($Footer.$Type.InsertParagraph())
                }
                if ($Header) {
                    $Paragraphs.Add($Header.$Type.InsertParagraph())
                }
            }
        } else {
            $Paragraphs.Add($Paragraph)
        }
        foreach ($CurrentParagraph in $Paragraphs) {
            $CurrentParagraph = Add-WordText -Paragraph $CurrentParagraph -Text $TextBefore -AppendToExistingParagraph -Alignment $Alignment

            if ($Option -eq 'Both' -or $Option -eq 'PageNumberOnly') {
                $CurrentParagraph.AppendPageNumber($PageNumberFormat)
            }
            $CurrentParagraph = Add-WordText -Paragraph $CurrentParagraph -Text $TextMiddle -AppendToExistingParagraph
            if ($Option -eq 'Both' -or $Option -eq 'PageCountOnly') {
                $CurrentParagraph.AppendPageCount($PageNumberFormat)
            }
            $CurrentParagraph = Add-WordText -Paragraph $CurrentParagraph -Text $TextAfter -AppendToExistingParagraph

            #$CurrentParagraph = Set-WordTextAlignment -Paragraph $CurrentParagraph
        }
        if ($Supress) { return } else {return $Paragraphs }
    } else {
        Write-Warning -Message 'Add-WordPageCount - Footer or Header or Paragraph is required.'
    }
}