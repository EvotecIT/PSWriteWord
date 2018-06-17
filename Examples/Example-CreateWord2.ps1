Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord2.docx"

$WordDocument = New-WordDocument $FilePath
$p1 = $WordDocument.InsertParagraph("This is a text").FontSize("10").SpacingAfter(50)
$p1.Alignment = 'center'


$p2 = $WordDocument.InsertParagraph("Like me like i do").FontSize("21")
$p2.Alignment = 'left'
$p2.Direction = [Direction]::RightToLeft # 'RightToLeft'

$p3 = $WordDocument.InsertParagraph("Process").FontSize("15")
$p3.Direction = [Direction]::RightToLeft
$p3.Alignment = 'both'

Save-WordDocument $WordDocument