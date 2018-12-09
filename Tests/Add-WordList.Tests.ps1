$ListOfItems1 = @('Test1')
$ListOfItems2 = @('Test1', 'Test2')
$ListOfItems5 = @('Test1', 'Test2', 'Test3', 'Test4', 'Test5')

$ListOfItemsNotArray1 = 'Test1'
$ListOfItemsNotArray2 = $false
$ListOfItemsNotArray3 = $false, $true

Describe 'Add-WordList' {
    It 'Given single string to Add-WordList should properly create a list' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray1 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 1
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'Test1'
    }
    It 'Given single bool to Add-WordList should properly create a list' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray2 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 1
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'False'
    }
    It 'Given two bools ($false/$True) string to Add-WordList should properly create a list with 2 entries' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItemsNotArray3 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 2
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'False'
        $WordDocument.Lists[0].Items[1].Text | Should -Be 'True'
    }

    It 'Given Array with 1 element to Add-WordList should properly create a list with 1 entries' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems1 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 1
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'Test1'
    }

    It 'Given Array with 2 elements to Add-WordList should properly create a list with 2 entries' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems2 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 2
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'Test1'
        $WordDocument.Lists[0].Items[1].Text | Should -Be 'Test2'
    }

    It 'Given Array with 5 elements to Add-WordList should properly create a list with 5 entries' {
        $WordDocument = New-WordDocument
        Add-WordList -WordDocument $WordDocument -ListType Bulleted -ListData $ListOfItems5 -Supress $True #-Verbose
        $WordDocument.Lists.Count | Should -Be 1
        $WordDocument.Lists[0].Items.Count | Should -Be 5
        $WordDocument.Lists[0].Items[0].Text | Should -Be 'Test1'
        $WordDocument.Lists[0].Items[1].Text | Should -Be 'Test2'
        $WordDocument.Lists[0].Items[2].Text | Should -Be 'Test3'
        $WordDocument.Lists[0].Items[3].Text | Should -Be 'Test4'
        $WordDocument.Lists[0].Items[4].Text | Should -Be 'Test5'
    }
}