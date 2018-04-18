# Open excel file and use specific sheet
    $objExcel=New-Object -ComObject Excel.Application
    $objExcel.Visible=$false
    $WorkBook=$objExcel.Workbooks.Open("C:\path\to\Volunteer\emails.xlsx")
    $worksheet = $WorkBook.sheets.Item(1)

# loop for each row of the excel file
    $intRowMax = ($worksheet.UsedRange.Rows).count
    for($intRow = 2 ; $intRow -le 75 ; $intRow++)
    {
    $name	= $worksheet.cells.item($intRow,2).value2
    $email	= $worksheet.cells.item($intRow,3).value2
    $times	= $worksheet.cells.item($intRow,5).value2
	
	$Outlook = New-Object -ComObject Outlook.Application
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = $email
	$Mail.Subject = "Volunteering Reminder for " + $name
	
	Write-Host $name "`r`n"

	$Mail.Body = "Dear " + $name + ",`r`n
	
Thank you for volunteering to help with the ACM/UPE Programming Contest! We are one week away from the event (Wednesday 4/18) and this is a reminder that you have signed up for the following times: `r`n" + 
	$times + 
"`r`n`r`nExpect another email later this week with additional information about your individual role and instructions on where to check-in. If you need to change your times, please let us know immediately so we can account for that.`r`n`
	
We really appreciate your time, and are depending on your help for this event to run smoothly. If you have any questions please let us know!`r`n`r`n
	
Thank you!`r`n
UPE & ACM E-Boards `r`n`r`n
"
	
    $Mail.Send()
	
    }  
    $WorkBook.close()
    $objexcel.quit()