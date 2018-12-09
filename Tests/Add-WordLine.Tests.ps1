Describe 'Add-WordLine' {
    It 'Given 4 new lines document should hold proper amount of paragraphs' {
        $WordDocument = New-WordDocument
        Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType double -Supress $True
        Add-WordLine -WordDocument $WordDocument -LineColor Blue -LineType single -LineSize 10 -Supress $True
        Add-WordLine -WordDocument $WordDocument -LineColor Red -LineType triple -Supress $True
        Add-WordLine -WordDocument $WordDocument -HorizontalBorderPosition top -LineColor Blue -LineType single -LineSize 10 -Supress $True
        $WordDocument.Paragraphs.Count | Should -Be 4
    }
}