Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent5.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

Add-WordToc -WordDocument $WordDocument -Title 'Test' -Switches C, A -RightTabPos 15 -HeaderStyle Heading5
Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1
Add-Section -WordDocument $WordDocument -PageBreak
$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is my second title' -HeadingType Heading1 -Color Red -CapsStyle caps -Supress $false
Add-Section -WordDocument $WordDocument -PageBreak
Add-WordText  -WordDocument $WordDocument -Text 'This is my third title' -HeadingType Heading2 -Italic $true -Bold $true

Add-WordToc -WordDocument $WordDocument -BeforeParagraph $Paragraph -Title 'Test' -Switches C, A -RightTabPos 15 -HeaderStyle Heading3 -MaxIncludeLevel 3

Save-WordDocument $WordDocument