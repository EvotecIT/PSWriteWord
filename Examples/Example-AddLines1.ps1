Import-Module PSWriteWord #-Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddLines1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and double. Horizonal Border Position by default bottom' -FontSize 10
Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType double

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be blue and single. Horizontal Border Position by default bottom' -FontSize 10
Add-WordLine -WordDocument $WordDocument -LineColor Blue -LineType single -LineSize 10

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and triple. Horizontal Border Position by default bottom' -FontSize 10
Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType triple

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be blue and single with Horizonal Border Position top' -FontSize 10
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition top -LineColor Blue -LineType single -LineSize 10

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and triple with Horizonal Border Position bottom' -FontSize 10
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition bottom -LineColor Red -LineType triple

Add-WordText -WordDocument $WordDocument -Text  'Adding line that should be blue and single with Horizonal Border Position top' -FontSize 10
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition top -LineColor Blue -LineType single -LineSize 10

Add-WordText -WordDocument $WordDocument -Text  'Adding line that should be blue and single with Horizonal Border Position bottom' -FontSize 10
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition bottom -LineColor Blue -LineType single -LineSize 10

Save-WordDocument $WordDocument -Language 'en-US'

### Start Word with file
Invoke-Item $FilePath