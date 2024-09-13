param (
    [string]$subject = "",
    [string]$name ="Akshat Gupta",
    [string]$regno="22BCE2173"
)

$Word = New-Object -ComObject word.application

$Word.Visible = $true

$Document = $Word.Documents.Add()

$Document.PageSetup.TopMargin = 72
$Document.PageSetup.BottomMargin = 72 
$Document.PageSetup.LeftMargin = 72   
$Document.PageSetup.RightMargin = 72  

$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = $subject+" DA"
$Range.Style = "Heading 1"
$Range.Font.Name = "Abadi Extra Light"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
$Range.Font.Size = 32
$Range.InsertParagraphAfter()

$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = $name
$Range.Style = "Normal"
$Range.Font.Name = "Abadi Extra Light"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphRight
$Range.Font.Size = 26
$Range.InsertParagraphAfter()

$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = $regno
$Range.Style = "Normal"
$Range.Font.Name = "Abadi Extra Light"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphRight
$Range.Font.Size = 26
$Range.InsertParagraphAfter()

$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "Question"
$Range.Style = "Normal"
$Range.Font.Name = "Abadi Extra Light"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
$Range.Font.Size = 19
$Range.InsertParagraphAfter()

$Range = $Document.Content.Paragraphs.Add().Range
$Range.Text = "Answer"
$Range.Style = "Normal"
$Range.Font.Name = "Abadi Extra Light"
$Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
$Range.Font.Size = 17
$Range.InsertParagraphAfter()

$Document.Sections.Item(1).Borders.Enable = $true

$Document.SaveAs([ref] "C:\path\to\your\document.docx")

# Release COM objects (optional but recommended to free up resources)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Range)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
