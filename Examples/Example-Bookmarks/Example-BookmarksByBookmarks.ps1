Import-Module PSWriteWord #-Force

### Before running this script make sure to run Example-CreateWord first
$FilePathTemplate = "$PSScriptRoot\Templates\CV-WithBookmarks.docx"

$WordDocument = Get-WordDocument -FilePath $FilePathTemplate

$WordBookmarks = Get-WordBookmark -WordDocument $WordDocument
$WordBookmarks | Format-Table -AutoSize

<#
Name        Paragraph
----        ---------
Name        Xceed.Document.NET.Paragraph
PhoneNumber Xceed.Document.NET.Paragraph
School1     Xceed.Document.NET.Paragraph
School2     Xceed.Document.NET.Paragraph

#>

Set-WordBookmark -BookmarkName 'Name' -BookmarkText 'Przemysław Kłys'
Set-WordBookmark -BookmarkName 'PhoneNumber' -BookmarkText '+48 500 500 500'
Set-WordBookmark -BookmarkName 'School1' -BookmarkText 'My super school'
Set-WordBookmark -BookmarkName 'School2' -BookmarkText 'My other super school'

Save-WordDocument -WordDocument $WordDocument -FilePath "$PSSCriptRoot\Output\CV-WithReplacedBookmarks.docx" -OpenDocument -Supress