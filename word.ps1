$word = New-Object -ComObject word.application
$word.visible = $true
$document = $word.documents.add()
$selection = $word.Selection
