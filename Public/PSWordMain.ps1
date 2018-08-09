function New-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath = ''
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    $WordDocument | Add-Member -MemberType NoteProperty -Name FilePath -Value $FilePath
    return $WordDocument
}

function Get-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Load($FilePath)
    $WordDocument | Add-Member -MemberType NoteProperty -Name FilePath -Value $FilePath
    return $WordDocument
}

function Save-WordDocument {
    [CmdletBinding()]
    param (
        [alias('Document')][parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        [alias('Path')][string] $FilePath,
        [string] $Language,
        [switch] $KillWord,
        [switch] $OpenDocument,
        [bool] $Supress = $false
    )
    if ($Language) {
        Write-Verbose "Save-WordDocument - Setting Language to $Language"
        $Paragraphs = Get-WordParagraphs -WordDocument $WordDocument
        foreach ($p in $Paragraphs) {
            Set-WordParagraph -Paragraph $p -Language $Language -Supress $True
        }
    }
    if (($KillWord) -and ($FilePath)) {
        $FileName = Split-Path $FilePath -leaf
        #$Process = get-process | Where { $_.MainWindowTitle -like "$FileName*"} | Select-Object id, name, mainwindowtitle | Sort-Object mainwindowtitle
        #$Process.MainWindowTitle
        Write-Verbose "Save-WordDocument - Killing Microsoft Word with text $FileName"
        $Process = Stop-Process -Name "$FileName*" -Confirm:$false -PassThru
        Write-Verbose "Save-WordDocument - Killed Microsoft Word: $FileName"
    }

    ### Saving PART
    if (-not $FilePath) {
        try {
            $FilePath = $WordDocument.FilePath
            Write-Verbose "Save-WordDocument - Saving document (Save: $FilePath)"
            $Data = $WordDocument.Save()
        } catch {
            $ErrorMessage = $_.Exception.Message
            if ($ErrorMessage -like "*The process cannot access the file*because it is being used by another process.*") {
                $FilePath = "$($([System.IO.Path]::GetTempFileName()).Split('.')[0]).docx"
                Write-Warning "Couldn't save file as it was in use. Trying different name $FilePath"
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
                $FilePath = "$($([System.IO.Path]::GetTempFileName()).Split('.')[0]).docx"
                Write-Warning "Couldn't save file as it was in use. Trying different name $FilePath"
                $Data = $WordDocument.SaveAs($FilePath)
            }
        }
    }
    ### Saving PART

    If ($OpenDocument) { Invoke-Item $FilePath }
    if ($Supress) { return } else { return $FilePath }
}