### Before running this script make sure to run Example-CreateWord first
$FilePath = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word.docx"
$FilePath1 = "C:\Users\przemyslaw.klys\OneDrive - Evotec\Desktop\Word1.docx"
$doc = Get-WordDocument -FilePath $FilePath
$word = "Sample"
$formatObj = New-Object Xceed.Document.NET.Formatting
$formatObj.FontColor = "Red"
foreach ($p in $doc.Paragraphs) {
    Set-WordTextReplace -Paragraph $p -SearchValue $word -ReplaceValue $word -NewFormatting $formatObj -Supress $false
}
Save-WordDocument -Document $doc -FilePath $FilePath1