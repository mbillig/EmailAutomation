# Open excel file and use specific sheet
$objExcel=New-Object -ComObject Excel.Application
$objExcel.Visible=$false
$WorkBook=$objExcel.Workbooks.Open("C:\path\to\volunteer_assignments.xlsx")
$worksheet = $WorkBook.sheets.Item(1)
#Generate all email files
for($iEmail = 0; $iEmail -le 74; $iEmail++){
	$pythonPath = "C:\path\to\python.exe"
	$pythonArgs = @("C:\path\to\formatScheduleEmail.py", $iEmail)
	& $pythonPath $pythonArgs
}

# loop for each row of the excel file
for($intRow = 72 ; $intRow -le 74 ; $intRow++)
{
	$name = $worksheet.cells.item($intRow,2).value2
	$email	= $worksheet.cells.item($intRow,3).value2
	$filePath = "C:\path\to\IndividualEmails\"
	$fileExt = ".txt"
	$fileName = $filePath + $email + $fileExt
	Write-Host $fileName "`r`n"
	$text = [IO.File]::ReadAllText($fileName)
	
	$Outlook = New-Object -ComObject Outlook.Application
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = $email
	$Mail.Subject = "Volunteering Info for " + $name

	$Mail.HTMLBody = $text
	
	$Mail.Send()

}  
$WorkBook.close()
$Outlook.quit()
$objexcel.quit()