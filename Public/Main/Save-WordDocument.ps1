function Save-WordDocument {
    [CmdletBinding()]
    param (
        [alias('Document')][parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $false)][Xceed.Words.NET.Container]$WordDocument,
        [alias('Path')][string] $FilePath,
        [string] $Language,
        [switch] $KillWord,
        [switch] $OpenDocument,
        [bool] $Supress = $false
    )
    if ($Language) {
        Write-Verbose -Message "Save-WordDocument - Setting Language to $Language"
        $Paragraphs = Get-WordParagraphs -WordDocument $WordDocument
        foreach ($p in $Paragraphs) {
            Set-WordParagraph -Paragraph $p -Language $Language -Supress $True
        }
    }
    if (($KillWord) -and ($FilePath)) {
        $FileName = Split-Path $FilePath -leaf
        #$Process = get-process | Where { $_.MainWindowTitle -like "$FileName*"} | Select-Object id, name, mainwindowtitle | Sort-Object mainwindowtitle
        #$Process.MainWindowTitle
        Write-Verbose -Message "Save-WordDocument - Killing Microsoft Word with text $FileName"
        $Process = Stop-Process -Name "$FileName*" -Confirm:$false -PassThru
        Write-Verbose -Message "Save-WordDocument - Killed Microsoft Word: $FileName"
    }

    ### Saving PART
    if (-not $FilePath) {
        try {
            $FilePath = $WordDocument.FilePath
            Write-Verbose -Message "Save-WordDocument - Saving document (Save: $FilePath)"
            $Data = $WordDocument.Save()
        } catch {
            $ErrorMessage = $_.Exception.Message
            if ($ErrorMessage -like "*The process cannot access the file*because it is being used by another process.*") {
                $FilePath = Get-FileName -Temporary -Extension 'docx'
                Write-Warning -Message "Couldn't save file as it was in use. Trying different name $FilePath"
                $Data = $WordDocument.SaveAs($FilePath)
            }
        }
    } else {
        try {
            Write-Verbose "Save-WordDocument - Saving document (Save AS: $FilePath)"
            $Data = $WordDocument.SaveAs($FilePath)
        } catch {
            $ErrorMessage = $_.Exception.Message
            if ($ErrorMessage -like "*The process cannot access the file*because it is being used by another process.*") {
                $FilePath = Get-FileName -Temporary -Extension 'docx'
                Write-Warning -Message "Couldn't save file as it was in use. Trying different name $FilePath"
                $Data = $WordDocument.SaveAs($FilePath)
            }
        }
    }
    ### Saving PART

    If ($OpenDocument) {
        if (($FilePath -ne '') -and (Test-Path -LiteralPath $FilePath)) {
            Invoke-Item -Path $FilePath
        } else {
            Write-Warning -Message "Couldn't open file as it doesn't exists - $FilePath"
        }
    }
    if ($Supress) { return } else { return $FilePath }
}