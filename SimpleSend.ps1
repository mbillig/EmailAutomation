$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "you@you.com"
$Mail.Subject = "Test"
$Mail.Body ="testing testing 123"
$Mail.Send()