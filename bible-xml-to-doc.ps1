# Update below file name variable accordinglly 
$fileName = "English_nkjv"
#$fileName ="TeluguBible"

# word document files will be written to the below folder
$folderName = "output"

# Get the content of the xml file
$doc = "$PSScriptRoot\bibles\$fileName.xml"
[xml]$xmlDoc = Get-Content -Path $doc -Encoding UTF8

$books = $xmlDoc.GetElementsByTagName("BIBLEBOOK")

foreach($book in $books){

    $word = New-Object -ComObject Word.Application
    $word.Visible = $True
    $Document = $word.Documents.Add()
    $Selection = $word.Selection    
    $Selection.TypeParagraph()
    $Selection.Font.Bold = 1
    $Selection.Font.Size = 18
    #$Selection.Font.Name = "Mallanna"
    $Selection.TypeText($book.bname)

    $chapters = $book.GetElementsByTagName("CHAPTER")

    foreach ($chapter in $chapters) {

        $verses = $chapter.GetElementsByTagName("VERS")

        foreach ($verse in $verses) {

            $Selection.Font.Size = 12
            $Selection.TypeParagraph()
            $Selection.Font.Italic = 1  
            $Selection.Font.Bold = 1          
            $Selection.Font.Underline = 1          
            $verseheading = $book.bname + " " + $chapter.cnumber + ":" + $verse.vnumber
            $Selection.TypeText($verseheading)
            $Selection.TypeParagraph()
            $Selection.Font.Bold = 0
            $Selection.Font.Italic = 0
            $Selection.Font.Underline = 0
            $Selection.TypeText($verse.InnerText)  
        }
        
    }

    # Creating the document file and writing the data
    $Report = "$PSScriptRoot\$folderName\$($book.bname).docx"
    $Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
    $word.Quit()
    
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    Remove-Variable word
}

<# Code to Create folders for each book and Chapter with an Empty file 
foreach($book in $books){

    New-Item -Path "$PSScriptRoot\$folderName\$($book.bnumber)" -ItemType Directory -Force
    $chapters = $book.GetElementsByTagName("CHAPTER")

    foreach ($chapter in $chapters) {
        New-Item -Path "$PSScriptRoot\$folderName\$($book.bnumber)\$($chapter.cnumber)" -ItemType Directory -Force
        New-Item -Path "$PSScriptRoot\$folderName\$($book.bnumber)\$($chapter.cnumber)\$($chapter.cnumber).txt" -ItemType File -Force
        $verses = $chapter.GetElementsByTagName("VERS")
        foreach ($verse in $verses) {
            New-Item -Path "$PSScriptRoot\$folderName\$($book.bname)\$($chapter.cnumber)\$($verse.vnumber)" -ItemType Directory -Force
        }
    }
}
#>