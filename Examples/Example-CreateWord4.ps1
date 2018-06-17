Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord4.docx"

$WordDocument = New-WordDocument $FilePath
$p1 = $WordDocument.InsertParagraph("This is a text").FontSize("10").SpacingAfter(50).UnderlineStyle([UnderlineStyle]::singleLine)
$p1.Alignment = 'center'


$p2 = $WordDocument.InsertParagraph("Like me like i do").FontSize("21").SpacingBefore(15).Bold()
$p2.Alignment = 'left'
$p2.Direction = [Direction]::RightToLeft # 'RightToLeft'
#$p2

$p3 = $WordDocument.InsertParagraph("Process").FontSize("15").Color([System.Drawing.Color]::Brown).Font('Arial').Italic()
$p3.Direction = [Direction]::RightToLeft
$p3.Alignment = 'both'

Save-WordDocument $WordDocument