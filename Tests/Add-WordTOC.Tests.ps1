Describe 'Add-WordTOC' {

    Context 'Testing Table Of Contents' {
        $ListOfHeaders = @('This is 1st section', 'This is 2nd section', 'This is 3rd section', 'This is 4th section', 'This is 5th section')
        $WordDocument = New-WordDocument
        $HasElements = $WordDocument.Xml.HasElements
        $ListsCount = $WordDocument.Lists.Count
        $CountItems = $WordDocument.Lists[0].Items.Count
        $Toc = $WordDocument | Add-WordToc -Title 'Table of content' -Switches C, A -RightTabPos 15 -HeaderStyle Heading1 -Supress $false
        foreach ($Section in $ListOfHeaders) {
            $Paragraph = $WordDocument | Add-WordTocItem -Text $Section -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
            $Paragraph = $WordDocument | Add-WordText -Text 'This is my test. Added after TOC Item.' -Color Orange
        }
        $Paragraph = $WordDocument | Add-WordTocItem -Text 'Adding another one' -ListLevel 0 -ListItemType Numbered -HeadingType Heading1
        $Paragraph = $WordDocument | Add-WordText -Text 'This is my test - outside of loop. Added after TOC Item.' -Color Red
        $WordDocument.Lists.Count | Should -Be 6
        for ($i = 0; $i -le $WordDocument.Lists.Count; $i++) {
            $WordDocument.Lists[0].Items.Count | Should -Be 1
        }


        It 'Creating New-WordDocument should not throw errors' {
            $HasElements | Should -Be $True
        }
        It 'Lists count on empty document should be 1' {
            $ListsCount | Should -Be 1
        }
        It 'Lists items count on empty document should be 0' {
            $CountItems | Should -Be 0
        }
        It 'Word Document Table of Content Value should contain TOC entry' {
            $Toc.Xml.Value.Trim() | Should -BeLike '*TOC*'
        }
        It 'Word Document should contain 6 lists' {
            $WordDocument.Lists.Count | Should -Be 6
        }
        for ($i = 0; $i -le $WordDocument.Lists.Count; $i++) {
            It "List #$i should contain just 1 element" {
                $WordDocument.Lists[0].Items.Count | Should -Be 1
            }
        }

    }
}