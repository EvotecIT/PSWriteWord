#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 #-Force

Describe 'Add-WordTable' {
    It 'Given 2 tables, document should have 2 tables with proper design' {
        $WordDocument = New-WordDocument
        $Object2 = Get-PSDrive
        $WordDocument | Add-WordTable -DataTable $Object2 -Design 'ColorfulList' #-Verbose
        $WordDocument | Add-WordTable -DataTable $Object2 -Design "LightShading" #-Verbose
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulList'
        $WordDocument.Tables[1].Design | Should -Be 'LightShading'
        $WordDocument.Tables.Count | Should -Be 2
    }
}
