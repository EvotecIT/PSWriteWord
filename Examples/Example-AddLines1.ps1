Import-Module PSWriteWord -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddLines1.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and double. Horizonal Border Position by default bottom' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType double -Supress $True

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be blue and single. Horizontal Border Position by default bottom' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -LineColor Blue -LineType single -LineSize 10 -Supress $True

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and triple. Horizontal Border Position by default bottom' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType triple -Supress $True

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be blue and single with Horizonal Border Position top' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition top -LineColor Blue -LineType single -LineSize 10 -Supress $True

Add-WordText -WordDocument $WordDocument -Text 'Adding line that should be red and triple with Horizonal Border Position bottom' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition bottom -LineColor Red -LineType triple -Supress $True

Add-WordText -WordDocument $WordDocument -Text  'Adding line that should be blue and single with Horizonal Border Position top' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition top -LineColor Blue -LineType single -LineSize 10 -Supress $True

Add-WordText -WordDocument $WordDocument -Text  'Adding line that should be blue and single with Horizonal Border Position bottom' -FontSize 10 -Supress $True
Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition bottom -LineColor Blue -LineType single -LineSize 10 -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True

### Start Word with file
Invoke-Item $FilePath