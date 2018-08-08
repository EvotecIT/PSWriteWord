function New-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath = ''
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    return $WordDocument
}

function Get-WordDocument {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][alias('Path')][string] $FilePath
    )
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Load($FilePath)
    return $WordDocument
}

function Save-WordDocument {
    [CmdletBinding()]
    param (
        [alias('Document')][parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory = $true)][Xceed.Words.NET.Container]$WordDocument,
        [alias('Path')][string] $FilePath,
        [string] $Language,
        [switch] $KillWord,
        [bool] $Supress = $false
    )
    if ($Language) {
        Write-Verbose "Save-WordDocument - Setting Language to $Language"
        $Paragraphs = Get-WordParagraphs -WordDocument $WordDocument
        foreach ($p in $Paragraphs) {
            Set-WordParagraph -Paragraph $p -Language $Language -Supress $Supress
        }
    }
    if (($KillWord) -and ($FilePath -ne '')) {
        $FileName = Split-Path $FilePath -leaf
        #$Process = get-process | Where { $_.MainWindowTitle -like "$FileName*"} | Select-Object id, name, mainwindowtitle | Sort-Object mainwindowtitle
        #$Process.MainWindowTitle
        Write-Verbose "Save-WordDocument - Killing Microsoft Word with text $FileName"
        $Process = Stop-Process -Name "$FileName*" -Confirm:$false -PassThru
        Write-Verbose "Save-WordDocument - Killed Microsoft Word: $FileName"
    }
    if (-not $FilePath) {
        Write-Verbose 'Save-WordDocument - Saving document without FilePath'
        $WordDocument.Save()
    } else {
        Write-Verbose "Save-WordDocument - Saving document to $FilePath"
        $WordDocument.SaveAs($FilePath)
    }
    #if ($Supress) { return } else { return $WordDocument }
}