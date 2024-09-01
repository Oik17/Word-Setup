# Create a new Word application object
$Word = New-Object -ComObject word.application

# Make the application visible (optional)
$Word.Visible = $true

# Add a new document
$Document = $Word.Documents.Add()

# Set the margins (in points; 72 points = 1 inch)
$Document.PageSetup.TopMargin = 72   # 1 inch
$Document.PageSetup.BottomMargin = 72 # 1 inch
$Document.PageSetup.LeftMargin = 72   # 1 inch
$Document.PageSetup.RightMargin = 72  # 1 inch

# Add a centered heading
$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "THIS DA"
$Range.Style = "Heading 1"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
$Range.Font.Size = 43
$Range.InsertParagraphAfter()

# Add text aligned to the right side of the page
$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "This is the text on the right side."
$Range.Style = "Normal"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphRight
$Range.Font.Size = 28
$Range.InsertParagraphAfter()

# Add a border around the page
$Document.Sections.Item(1).Borders.Enable = $true

# Save the document
$Document.SaveAs([ref] "C:\path\to\your\document.docx")


# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
