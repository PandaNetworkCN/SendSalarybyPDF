# Need install
# https://wkhtmltopdf.org/downloads.html
# C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe
# Install module ImportExcel Install-Module -Name ImportExcel via admin

# Import the module
Import-Module -Name ImportExcel

#Get current running folder. 
$currentFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Specify the path to the Excel file
$excelFilePath = "$currentFolder\salrylist.xlsx"

# Read the Excel file
$excelData = Import-Excel -Path $excelFilePath

#Get all properties for object $excelData 
#$properties = $excelData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# Process each row
foreach ($row in $excelData) {

   $ID = $row.'2.员工 ID'

   if($ID -ne $null)
   {
    #Get Emplyee ID as the default file name
    $mail = $row.'23.Email'
    $Name = $row.'1.姓名'
    
#PDF file path
$pdfFilePath = "$currentFolder\$ID.pdf"

#Email settings
$smtpServer = "service.demo.com"
$senderEmail = "BeijingHR@demo.com"
$recipientEmail = $mail
$emailSubject = "工资单"
$emailBody = "
$Name 收

请查收附件工资单
 
"
echo $recipientEmail
echo $pdfFilePath

# Send the email with the PDF attachment
Send-MailMessage -SmtpServer $smtpServer -From $senderEmail -To $recipientEmail -Subject $emailSubject -Body $emailBody -Attachments $pdfFilePath -Encoding UTF8 -Verbose

}
}




