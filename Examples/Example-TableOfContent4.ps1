Import-Module PSWriteWord -Force
Clear-Host
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-TableOfContent4.docx"
$ListOfItems = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')
$ListOfHeaders = @('This is 1st section', 'This is 2nd section', 'This is 3rd section', 'This is 4th section', 'This is 5th section')

$WordDocument = New-WordDocument -FilePath $FilePath
$toc = $WordDocument.InsertTableOfContents("Table of content", 1)

#Add-WordText -WordDocument $WordDocument -Text 'This is my first title' -HeadingType Heading1

### This list will be converted into Headings for Numbered Table of Contents
$ListHeaders = Add-List -WordDocument $WordDocument -ListType Numbered -ListData $ListOfHeaders -Supress $false
### This list will be added added but....it will appear in the end... since we will use Add-WordText with $Paragraph
$List2 = Add-List -WordDocument $WordDocument -ListType Numbered -ListData $ListOfItems -Supress $false

### Converts List into numbered Headings for Table of Content
$Headings = Convert-ListToHeadings -WordDocument $WordDocument -List $ListHeaders

### Notice how this gets added under Test2 in 1st numbered list... essentially putting $List2 to the end.
$paragraph2 = Add-WordText -WordDocument $WordDocument `
    -Paragraph $Headings[2] -Text 'This is a text that will be added to ', ' 2nd ', 'section' `
    -Color Black, Red, Black -Supress $false
$paragraph2 = Add-WordText -WordDocument $WordDocument -Paragraph $Paragraph2 -Text 'This will continue getting added to that section... of course colored as red' -Color Red -Supress $false
$paragraph2 = Add-WordText -WordDocument $WordDocument `
    -Paragraph $Paragraph2 `
    -Text 'You need to use', ' Supress ', '$false ', 'to get paragraph values...', '' `
    -Color Black, Green, Red, Black -Supress $false
$paragraph2 = Add-WordText -WordDocument $WordDocument `
    -Paragraph $Paragraph2 `
    -Text "If you won't do that it will not send ", '$paragraph', ' value back ', 'essentially adding text to end of the document.' `
    -Color Black, Green, Red `
    -Supress $false # important...

$paragraph4 = Add-WordText -WordDocument $WordDocument `
    -Paragraph $Headings[4] -Text 'This is a text that will be added to ', ' 4th ', 'section' `
    -Color Black, Red, Black -Supress $false

$paragraph1 = Add-WordText -WordDocument $WordDocument `
    -Paragraph $Headings[1] -Text 'This is a text that will be added to ', ' 1st ', 'section' `
    -Color Black, Red, Black -Supress $false

Save-WordDocument $WordDocument