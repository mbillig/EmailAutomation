# Open excel file and use specific sheet
    $objExcel=New-Object -ComObject Excel.Application
    $objExcel.Visible=$false
    $WorkBook=$objExcel.Workbooks.Open("Volunteer_Allocation_per_Volunteer.xlsx")
    $worksheet = $WorkBook.sheets.Item(4)

# loop for each row of the excel file
    $intRowMax = ($worksheet.UsedRange.Rows).count
    for($intRow = 1 ; $intRow -le 70 ; $intRow++)
    {
    $email	= $worksheet.cells.item($intRow,1).value2
	
	$Outlook = New-Object -ComObject Outlook.Application
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = $email
	$Mail.Subject = "Thank you!"

	Write-Host $email

	$Mail.Body = "Hello Volunteers!`r`n`r`nThank you very much for all your help today! A lot of time and effort goes into planning this event, but we know there were still some rough edges today, and we really appreciate your patience and flexibility throughout the day. Overall, the event was a huge success and we couldn't have done it without you! `r`n`r`nBest wishes,`r`nACM and UPE E-boards"
    $Mail.Send()
	
    }  
    $WorkBook.close()
    $objexcel.quit()