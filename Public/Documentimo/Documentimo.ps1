#using namespace Xceed.Words.NET
#using namespace Xceed.Document.NET

function Documentimo {
    [CmdletBinding()]
    [alias('Doc', 'New-Documentimo')]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Documentimo requires opening and closing brace."),
        [string] $FilePath,
        [alias('Show')][switch] $Open,
        [string] $Language = 'en-US'
    )
    $WordDocument = New-WordDocument -FilePath $FilePath
    if ($null -ne $Content) {
        $Array = Invoke-Command -ScriptBlock $Content
        New-WordProcessing -Content $Array -WordDocument $WordDocument
    }
    Save-WordDocument -WordDocument $WordDocument -Supress $true -Language $Language -Verbose -OpenDocument:$Open
}