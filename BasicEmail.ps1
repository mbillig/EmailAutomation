# Open excel file and use specific sheet
    $objExcel=New-Object -ComObject Excel.Application
    $objExcel.Visible=$false
    $WorkBook=$objExcel.Workbooks.Open(pathToXLSXFileOfEmails)
    $worksheet = $WorkBook.sheets.Item(1)


# loop for each row of the excel file
    $intRowMax = ($worksheet.UsedRange.Rows).count
    for($intRow = 1 ; $intRow -le 37 ; $intRow++)
    {
    $lastName	= $worksheet.cells.item($intRow,1).value2
    $firstName	= $worksheet.cells.item($intRow,2).value2
    $email	= $worksheet.cells.item($intRow,3).value2
	
	$Outlook = New-Object -ComObject Outlook.Application
	$Mail = $Outlook.CreateItem(0)
	$Mail.To = $email
	$Mail.Subject = "UPE Invitation Correction"
	

	Write-Host $firstName " " $lastName "`r`n"

	$Mail.HTMLBody = "Dear " + $firstName + " " + $lastName + ":<br><br>

It has been brought to my attention that the first email omitted the actual cost of the UPE dues. Dues this year will be 80 dollars. I realize that is rather pricey, but this is because most of the cost goes towards to your international UPE membership fees. If you have any concerns about the cost, please let me know. <br><br>

Please make checks payable to 'Upsilon Pi Epsilon Wisconsin Beta Chapter'. Payment must be received by Thursday, April 26, and may be dropped off at the Upsilon Pi Epsilon mailbox in Cudahy Hall, Room 340, or brought to the induction ceremony.

We will hold an official induction ceremony and reception: <br><br>

Thursday, April 26, 2018, at 8:00 PM in Cudahy 401 <br><br>

If you are able to join us, please RSVP to the <a href=https://docs.google.com/forms/d/e/1FAIpQLSc7kebbIE28AZNPzj0fjj-iqA1uEJZdB-xxDsT0cdP5HAqOjA/viewform?usp=sf_link>google form</a> by Monday, April 24 (ASAP preferred).  <br><br>

Also, if you are graduating in May, please note so in your RSVP. <br><br>

If you have questions concerning Upsilon Pi Epsilon or the evening of the initiation, please let us know! <br><br><br>


Again congratulations! <br><br>

Best Regards, <br><br>

Upsilon Pi Epsilon (Wisconsin Beta Chapter), Marquette University <br>
Marielle Billig - President <br>
Alex Gattone - Vice President <br>
Shivani Kohli - Treasurer <br>
Dennis Brylow - Faculty Advisor <br><br><br>"

 
    $Mail.Send()
	
    }  
    $WorkBook.close()
    $objexcel.quit()