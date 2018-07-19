Import-Module PSWriteWord #-Force
#Requires -Modules Pester

Describe 'Add-WordText' {

    It 'Given -Text parameter should create 1 paragraph' {
        $WordDocument = New-WordDocument $FilePath
        $WordDocument | Add-WordText -Text 'This is text'
        $WordDocument.Paragraphs.Count | Should -Be 1
    }
    It 'Given -Text parameter should create 1 paragraph with a text field This is text' {
        $WordDocument = New-WordDocument $FilePath
        $WordDocument | Add-WordText -Text 'This is text'
        $WordDocument.Paragraphs[0].Text | Should -Be 'This is text'
    }

}