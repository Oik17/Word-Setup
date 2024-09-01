$word = New-Object -ComObject word.application
$word.visible = $true
$document = $word.documents.add()
$selection = $word.Selection

# Set the margins (in points; 72 points = 1 inch)
$Document.PageSetup.TopMargin = 72   # 1 inch
$Document.PageSetup.BottomMargin = 72 # 1 inch
$Document.PageSetup.LeftMargin = 72   # 1 inch
$Document.PageSetup.RightMargin = 72  # 1 inch

# Add a heading
$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "THIS DA"
$Range.Style = "Heading 1"  git 
$Range.InsertParagraphAfter()

# Add body text
$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "This is the body of the document."
$Range.Style = "Normal"
$Range.InsertParagraphAfter()


# Save the document
$Document.SaveAs([ref] "C:\path\to\your\document.docx")

$Word.Quit()
