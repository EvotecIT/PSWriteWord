# Paragraph InsertParagraph()
# Paragraph InsertParagraph( int index, string text, bool trackChanges )
# Paragraph InsertParagraph( Paragraph p )
# Paragraph InsertParagraph( int index, Paragraph p )
# Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
# Paragraph InsertParagraph( string text )
# Paragraph InsertParagraph( string text, bool trackChanges )
# Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )

function Add-WordText {

    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [alias ("T")] [String[]]$Text,
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("N")] [string[]] $FontName = @(),
        [alias ("B")] [bool[]] $Bold = @(),
        [alias ("I")] [bool[]] $Italic = @(),
        [alias ("U")] [UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [highlight[]] $Highlight = @(),
        [alias ("CA")] [CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [HeadingType[]] $HeadingType = @(),
        [bool[]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph
    )
    $p = $WordDocument.InsertParagraph()
    for ($i = 0; $i -lt $Text.Length; $i++) {
        if ($NewLine[$i] -ne $null -and $NewLine[$i] -eq $true) {
            if ($i -gt 0) { $p = $WordDocument.InsertParagraph() }
            $p = $p.Append($Text[$i])
        } else {
            $p = $p.Append($Text[$i])
        }
        if ($Color[$i] -ne $null) {
            $p = $p.Color($Color[$i])
        }
        if ($FontSize[$i] -ne $null) {
            $p = $p.FontSize($FontSize[$i])
        }
        if ($FontName[$i] -ne $null) {
            $p = $p.Font($FontName[$i])
        }
        if ($Bold[$i] -ne $null -and $Bold[$i] -eq $true) {
            $p = $p.Bold()
        }
        if ($Italic[$i] -ne $null -and $Italic[$i] -eq $true) {
            $p = $p.Italic()
        }
        if ($UnderlineStyle[$i] -ne $null) {
            $p = $p.UnderlineStyle($UnderlineStyle[$i])
        }
        if ($SpacingAfter[$i] -ne $null) {
            $p = $p.SpacingAfter($SpacingAfter[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $p = $p.SpacingBefore($SpacingBefore[$i])
        }
        if ($SpacingBefore[$i] -ne $null) {
            $p = $p.SpacingBefore($SpacingBefore[$i])
        }
        if ($Spacing[$i] -ne $null) {
            $p = $p.Spacing($Spacing[$i])
        }
        if ($Highlight[$i] -ne $null) {
            $p = $p.Highlight($Highlight[$i])
        }
        if ($CapsStyle[$i] -ne $null) {
            $p = $p.CapsStyle($CapsStyle[$i])
        }
        if ($StrikeThrough[$i] -ne $null) {
            $p = $p.StrikeThrough($StrikeThrough[$i])
        }
        if ($HeadingType[$i] -ne $null) {
            $p = $p.HeadingType($HeadingType[$i])
        }
    }



    #$DefaultColor = $Color[0]
    #if ($BackGroundColor -ne $null -and $BackGroundColor.Count -ne $Color.Count) { Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated." ; return }
    if ($Text.Count -eq 0) { return }
    return
    #if ($LinesBefore -ne 0) {  for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host "`n" -NoNewline } } # Add empty line before
    #if ($ShowTime) { Write-Host "[$([datetime]::Now.ToString($TimeFormat))]" -NoNewline} # Add Time before output
    #if ($StartTab -ne 0) {  for ($i = 0; $i -lt $StartTab; $i++) { Write-Host "`t" -NoNewLine } }  # Add TABS before text

    if ($Color.Count -ge $Text.Count) {
        # the real deal coloring
        if ($BackGroundColor -eq $null) {
            $p = $WordDocument.InsertParagraph() #| Out-Null
            for ($i = 0; $i -lt $Text.Length; $i++) {
                # Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine
                $p = $p.Append($Text[$i]).FontSize($FontSize[$i]).Color($Color[$i]) #| Out-Null
            }
        } else {
            for ($i = 1; $i -lt $Text.Length; $i++) {
                #Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine
                #$WordDocument.InsertParagraph($Text[$i]).FontSize($FontSize[$i]) #| Out-Null
            }
        }
    } else {
        if ($BackGroundColor -eq $null) {
            #   for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
            #  for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -NoNewLine }
        } else {
            # for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine }
            #for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackGroundColor[0] -NoNewLine }
        }
    }
    <#
    if ($NoNewLine -eq $true) { Write-Host -NoNewline } else { Write-Host } # Support for no new line
    if ($LinesAfter -ne 0) {  for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host "`n" } }  # Add empty line after
    if ($LogFile -ne "") {
        # Save to file
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        try {
            Write-Output "[$([datetime]::Now.ToString($TimeFormat))]$TextToFile" | Out-File $LogFile -Encoding unicode -Append
        } catch {
            $_.Exception
        }
    }
    #>
}