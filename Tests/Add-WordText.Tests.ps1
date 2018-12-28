Describe 'Add-WordText' {
    It 'Given Text parameter should create 1 paragraph with a text field This is text' {
        $WordDocument = New-WordDocument
        $WordDocument | Add-WordText -Text 'This is text'
        $WordDocument.Paragraphs[0].Text | Should -Be 'This is text'
        $WordDocument.Paragraphs.Count | Should -Be 1
    }


    It 'Should add 2 texts to Headers, 2 texts to footers and 1 text to content. Should test different scenarios.' {
        $WordDocument = New-WordDocument
        $Header = Add-WordHeader -WordDocument $WordDocument
        $Footer = Add-WordFooter -WordDocument $WordDocument
        $WordDocument | Add-WordText -Text 'This is text'

        Add-WordText -WordDocument $WordDocument -Paragraph $Footer.First.Paragraphs[0] -AppendToExistingParagraph -Text 'My Text in Footer - 1st paragraph' -Color Orange -Supress $True
        Add-WordText -WordDocument $WordDocument -Paragraph $Header.First.Paragraphs[0] -AppendToExistingParagraph -Text 'My Text in Header - 1st paragraph' -Color Orange -Supress $True

        # this adds new paragraph into header with new text
        $WordDocument | Add-WordText -Text 'This is text in header' -Header $Header.First -Color Red
         # this adds new paragraph into footer with new text
        $WordDocument | Add-WordText -Text 'This is text in footer' -Footer $Footer.First -Color Red

        $Header.First.Paragraphs[0].Text | Should -Be 'My Text in Header - 1st paragraph'
        $Footer.First.Paragraphs[0].Text | Should -Be 'My Text in Footer - 1st paragraph'

        $Header.First.Paragraphs[1].Text | Should -Be 'This is text in header'
        $Footer.First.Paragraphs[1].Text | Should -Be 'This is text in footer'

        $WordDocument.Paragraphs[0].Text | Should -Be 'This is text'
        $WordDocument.Paragraphs.Count | Should -Be 1
    }
}