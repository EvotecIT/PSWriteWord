Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent5.docx"

$WordDocument = New-WordDocument -FilePath $FilePath

Add-WordToc -WordDocument $WordDocument -Title 'Test' -Switches C, A -RightTabPos 15 -HeaderStyle Heading5 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1 -Supress $True
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is my second title' -HeadingType Heading1 -Color Red -CapsStyle caps -Supress $false
Add-WordSection -WordDocument $WordDocument -PageBreak -Supress $True
Add-WordText  -WordDocument $WordDocument -Text 'This is my third title' -HeadingType Heading2 -Italic $true -Bold $true -Supress $True

Add-WordToc -WordDocument $WordDocument -BeforeParagraph $Paragraph -Title 'Test' -Switches C, A -RightTabPos 15 -HeaderStyle Heading3 -MaxIncludeLevel 3 -Supress $True

Save-WordDocument $WordDocument -Supress $True
Invoke-Item $FilePath