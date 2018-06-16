# Paragraph InsertParagraph()
# Paragraph InsertParagraph( int index, string text, bool trackChanges )
# Paragraph InsertParagraph( Paragraph p )
# Paragraph InsertParagraph( int index, Paragraph p )
# Paragraph InsertParagraph( int index, string text, bool trackChanges, Formatting formatting )
# Paragraph InsertParagraph( string text )
# Paragraph InsertParagraph( string text, bool trackChanges )
# Paragraph InsertParagraph( string text, bool trackChanges, Formatting formatting )

function Add-WordText {
    <#
	.SYNOPSIS
	Write-Color is a wrapper around Write-Host.
	It provides:
	- easy manipulation of colors,
	- logging output to file (log)
	- nice formatting options out of the box.
	.DESCRIPTION
	Author: przemyslaw.klys at evotec.pl
	Project website: https://evotec.xyz/hub/scripts/write-color-ps1/
	Project support: https://github.com/EvotecIT/PSWriteColor

	Original idea: Josh (https://stackoverflow.com/users/81769/josh)

    version 0.5 (25th April 2018
    - added backgroundcolor
    - added aliases T/B/C to shorter code
    - added alias to function (can be used with "WC")
	- fixes to module publishing

	version 0.4.0-0.4.9 (25th April 2018)
	- published as module
	- fixed small issues

	version 0.31 (20th April 2018)
	- Added Try/Catch for Write-Output (might need some additional work)
	- Small change to parameters

	version 0.3 (9th April 2018)
	- added -ShowTime
	- added -NoNewLine
	- added function description
	- changed some formatting

	version 0.2
	- added logging to file

	version 0.1
	- first draft

	Notes:
	- TimeFormat https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx

	.EXAMPLE
	Write-Color -Text "Red ", "Green ", "Yellow " -Color Red,Green,Yellow

	Write-Color -Text "This is text in Green ",
					"followed by red ",
					"and then we have Magenta... ",
					"isn't it fun? ",
					"Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan

	Write-Color -Text "This is text in Green ",
					"followed by red ",
					"and then we have Magenta... ",
					"isn't it fun? ",
					"Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan -StartTab 3 -LinesBefore 1 -LinesAfter 1

	Write-Color "1. ", "Option 1" -Color Yellow, Green
	Write-Color "2. ", "Option 2" -Color Yellow, Green
	Write-Color "3. ", "Option 3" -Color Yellow, Green
	Write-Color "4. ", "Option 4" -Color Yellow, Green
	Write-Color "9. ", "Press 9 to exit" -Color Yellow, Gray -LinesBefore 1

	Write-Color -LinesBefore 2 -Text "This little ","message is ", "written to log ", "file as well." `
				-Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt" -TimeFormat "yyyy-MM-dd HH:mm:ss"
	Write-Color -Text "This can get ","handy if ", "want to display things, and log actions to file ", "at the same time." `
				-Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt"


    # Added in 0.5
    Write-Color -T "My text", " is ", "all colorful" -C Yellow, Red, Green -B Green, Green, Yellow
    wc -t "my text" -c yellow -b green
    wc -text "my text" -c red

    #>
    [Alias("wc")]
    param (
        $WordDocument,
        [alias ("T")] [String[]]$Text,
        [alias ("C")] [ConsoleColor[]]$Color = "White",
        [alias ("B")] [ConsoleColor[]]$BackGroundColor = $null,
        [alias ("S")] [double[]] $FontSize = 10,
        [int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [alias ("L")] [string] $LogFile = "",
        [string] $TimeFormat = "yyyy-MM-dd HH:mm:ss",
        [switch] $ShowTime,
        [switch] $NoNewLine
    )

    $DefaultColor = $Color[0]
    #if ($BackGroundColor -ne $null -and $BackGroundColor.Count -ne $Color.Count) { Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated." ; return }
    if ($Text.Count -eq 0) { return }

    #if ($LinesBefore -ne 0) {  for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host "`n" -NoNewline } } # Add empty line before
    #if ($ShowTime) { Write-Host "[$([datetime]::Now.ToString($TimeFormat))]" -NoNewline} # Add Time before output
    #if ($StartTab -ne 0) {  for ($i = 0; $i -lt $StartTab; $i++) { Write-Host "`t" -NoNewLine } }  # Add TABS before text

    if ($Color.Count -ge $Text.Count) {
        # the real deal coloring
        if ($BackGroundColor -eq $null) {
            $p = $WordDocument.InsertParagraph($Text[0]).FontSize($FontSize[0]) #| Out-Null
            for ($i = 1; $i -lt $Text.Length; $i++) {
                # Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine
                $p = $p.Append($Text[$i]).FontSize($FontSize[$i]) #| Out-Null
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