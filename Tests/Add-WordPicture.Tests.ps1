#Requires -Modules Pester
Import-Module $PSScriptRoot\..\PSWriteWord.psd1 -Force

$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-AddPicture2.docx"
$FilePathImage1 = "$PSScriptRoot\..\Examples\Images\Logo-Evotec-Small.png"
$FilePathImage2 = "$PSScriptRoot\..\Examples\Images\Logo-Evotec-Small.jpg"

$TemporaryFolder = [IO.Path]::GetTempPath()

Describe 'Add-WordPicture' {
    It 'Given 2 pictures, one with rotation - Adding them in memory should work properly' {
        $FilePath = [IO.Path]::Combine($TemporaryFolder, "1.docx")

        $WordDocument = New-WordDocument $FilePath

        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Verbose
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage2 -Rotation 25

        $AllPictures = Get-WordPicture -WordDocument $WordDocument -ListPictures

        $AllPictures.Count | Should -Be 2
        $AllPictures[1].Rotation | Should -Be 25
    }
    It 'Given 2 pictures, one with rotation - Adding them, saving, loading work properly' {
        $FilePath = [IO.Path]::Combine($TemporaryFolder, "2.docx")

        $WordDocument = New-WordDocument $FilePath
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Verbose
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage2 -Rotation 25
        $AllPictures = Get-WordPicture -WordDocument $WordDocument -ListPictures

        $AllPictures.Count | Should -Be 2
        $AllPictures[1].Rotation | Should -Be 25

        Save-WordDocument $WordDocument -Language 'en-US' -Supress $true

        $ReadWord = Get-WordDocument -FilePath $FilePath
        $Pictures = Get-WordPicture -WordDocument $ReadWord -ListPictures
        $Pictures.Count | Should -Be 2
        $Pictures[1].Rotation | Should -Be 25
    }

    It 'Given 2 pictures, one with rotation - Adding them, saving, loading work properly, adding more pictures that already existed' {
        $FilePath = [IO.Path]::Combine($TemporaryFolder, "2.docx")

        $WordDocument = New-WordDocument $FilePath
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Verbose
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture... with rotation' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage2 -Rotation 25
        $AllPictures = Get-WordPicture -WordDocument $WordDocument -ListPictures

        $AllPictures.Count | Should -Be 2
        $AllPictures[1].Rotation | Should -Be 25

        Save-WordDocument $WordDocument -Language 'en-US' -Supress $true

        $ReadWord = Get-WordDocument -FilePath $FilePath
        $Pictures = Get-WordPicture -WordDocument $ReadWord -ListPictures
        $Pictures.Count | Should -Be 2
        $Pictures[1].Rotation | Should -Be 25

        $PlaceToAddPicture = Add-WordText -WordDocument $ReadWord -Text 'Adding a picture...' -Supress $false
        $PlaceToAddPicture.Text | Should -Be 'Adding a picture...'

        Add-WordText -WordDocument $ReadWord -Text 'This is text' -Supress $true
        Add-WordText -WordDocument $ReadWord -Text 'This is another text' -Supress $true

        Add-WordPicture -WordDocument $ReadWord -Picture $Pictures[0] -Paragraph $PlaceToAddPicture # add copy of picture to paragraph

        Add-WordText -WordDocument $ReadWord -Text 'Here we copy 1st picture from WordDocument and add it again'  -Supress $true
        Add-WordPicture -WordDocument $ReadWord -Picture $Pictures[1] # add copy of picture


        $OnePicture = Get-WordPicture -WordDocument $ReadWord -PictureID 0
        Add-WordPicture -WordDocument $ReadWord -Picture $OnePicture # add copy of picture

        $Pictures = Get-WordPicture -WordDocument $ReadWord -ListPictures
        $Pictures.Count | Should -Be 5
        $Pictures[3].Rotation | Should -Be 25
        Save-WordDocument $ReadWord -Language 'en-US' -Supress $true #-OpenDocument
    }
    It 'Given 3 pictures, set it align to center, right, both' {
        $FilePath = [IO.Path]::Combine($TemporaryFolder, "2.docx")

        $WordDocument = New-WordDocument $FilePath
        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Alignment center -Verbose

        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Alignment right -Verbose

        Add-WordText -WordDocument $WordDocument -Text 'Adding a picture...' -Supress $true
        Add-WordPicture -WordDocument $WordDocument -ImagePath $FilePathImage1 -Alignment both -Verbose

        $AllPictures = Get-WordPicture -WordDocument $WordDocument -ListParagraphs
        $AllPictures[0].Alignment | Should -Be Center
        $AllPictures[1].Alignment | Should -Be right
        $AllPictures[2].Alignment | Should -Be both
    }

}