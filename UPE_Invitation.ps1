# Open excel file and use specific sheet
    $objExcel=New-Object -ComObject Excel.Application
    $objExcel.Visible=$false
    $WorkBook=$objExcel.Workbooks.Open("complete\path\to\emails.xlsx")
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
	$Mail.Subject = "Congratulations " + $firstName + "! UPE Invitation"
	

	Write-Host $firstName " " $lastName "`r`n"

	$Mail.HTMLBody = "Dear " + $firstName + " " + $lastName + ":<br><br>

Congratulations! We are excited to extend an invitation to you to join the International Honor Society of Upsilon Pi Epsilon. <br><br>

This invitation comes in recognition of your excellent academic record at Marquette University. Your hard work toward your degree has met the highest expectations of the University, and you have been nominated by at least three of your professors or classmates; we invite you to national and chapter membership in Upsilon Pi Epsilon. <br><br>

<a href=http://upe.acm.org/>Upsilon Pi Epsilon</a> is the first and only international honor society for the computing and information disciplines. It has been endorsed by both the ACM (Association for Computing Machinery) and the IEEE Computer Society, our field's largest professional societies. UPE chapters across the country and around the world work to recognize outstanding scholarship, promote the advancement of computing, and support members in their continued education. The <a href=http://www.mscs.mu.edu/~upsilon/>Marquette Chapter</a> inducts students in the computer science and computer engineering majors. Members must uphold at least 5 hours of service per semester in order to maintain good standing. These hours are defaulted to scheduled tutoring hours organized by the UPE officers, in addition to the various computer science volunteering opportunities offered throughout the semester. <br><br>

We will hold an official induction ceremony and reception: <br><br>

Tuesday, April 26, 2018, at 8:00 PM in Cudahy 401 <br><br>

If you are able to join us, please RSVP to the <a href=https://docs.google.com/forms/d/e/1FAIpQLSc7kebbIE28AZNPzj0fjj-iqA1uEJZdB-xxDsT0cdP5HAqOjA/viewform?usp=sf_link>google form</a> by Monday, April 24 (ASAP preferred). A one-time $80 fee is required to cover membership dues. Dues primarily cover the cost of international membership as well as smaller items like the UPE pin, certificate, and honor cords for graduation.. Please make checks payable to 'Upsilon Pi Epsilon Wisconsin Beta Chapter'. Payment must be received by Monday, April 24 , and may be dropped off at the Upsilon Pi Epsilon mailbox in Cudahy Hall, Room 340. <br><br>

Also, if you are graduating in May, please note so in your RSVP. <br><br>

If you have questions concerning Upsilon Pi Epsilon or the evening of the initiation, please contact our officers  at upsilon-officers@mscs.mu.edu or Dr. Brylow at brylow@mscs.mu.edu. <br><br><br>


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