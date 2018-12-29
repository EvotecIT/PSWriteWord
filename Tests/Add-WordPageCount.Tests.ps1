Describe 'Add-WordPageCount/Add-WordPageNumber' {
    It 'Should add 1 text to content. Should add headers and footers. Should add page numbers and page count in different ways' {
        $WordDocument = New-WordDocument
        $Header = Add-WordHeader -WordDocument $WordDocument
        $Footer = Add-WordFooter -WordDocument $WordDocument
        $WordDocument | Add-WordText -Text 'This is text'

        $WordDocument.Paragraphs[0].Text | Should -Be 'This is text'
        $WordDocument.Paragraphs.Count | Should -Be 1

        Add-WordPageCount -Header $Header -PageNumberFormat normal -TextBefore 'Page Nr ' -TextMiddle ' of ' -TextAfter '' -Alignment center

        $Header.First.Paragraphs[1].Text | Should -Be 'Page Nr 1 of 1'
        $Header.Odd.Paragraphs[1].Text | Should -Be 'Page Nr 1 of 1'
        $Header.Even.Paragraphs[1].Text | Should -Be 'Page Nr 1 of 1'

        Add-WordPageCount -Footer $Footer -Type First -PageNumberFormat normal -Option PageNumberOnly

        $Footer.First.Paragraphs[1].Text | Should -Be '1'
        $Footer.First.Paragraphs[1].Alignment | Should -Be 'left'

        Add-WordPageCount -Footer $Footer -Type Odd -PageNumberFormat normal -Option Both -Alignment right -TextMiddle ' of '

        $Footer.Odd.Paragraphs[1].Text | Should -Be '1 of 1'

        $ParagraphArray = Add-WordPageNumber -Footer $Footer -Type All -PageNumberFormat roman -Option PageNumberOnly -Alignment left -TextBefore 'Page Number '

        # paragraph is 2nd because it was already taken above
        $Footer.First.Paragraphs[2].Text | Should -Be 'Page Number 1'

        $Footer.Even.Paragraphs[1].Text | Should -Be 'Page Number 1'

        # paragraph is 2nd because it was already taken above
        $Footer.Odd.Paragraphs[2].Text | Should -Be 'Page Number 1'

        # this will be overwrtitten by next one
        $Footer.Odd.Paragraphs[2].Alignment | Should -Be 'left'

        $ParagraphArray.Count | Should -Be 3

        # this basically takes only 1 paragraph on the first footer (odd, even footers have their own paragraphs)
        # and adds page count only to first footer along with text
        # it also changes alingment on that 1 paragraph
        Add-WordPageCount -Paragraph $ParagraphArray[0] -PageNumberFormat normal -Alignment right -TextMiddle ' / ' -Option PageCountOnly

        $Footer.First.Paragraphs[2].Alignment | Should -Be 'right'
        # with last command we just added "/ 1", and "Page Number 1" was already there so we expect 'Page Number 1 / 1'
        $Footer.First.Paragraphs[2].Text | Should -Be 'Page Number 1 / 1'

    }
}