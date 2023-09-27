# 1. Need install wkhtmltopdf.exe x 64 from https://wkhtmltopdf.org/downloads.html
#    C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe
# 2. Need Install module ImportExcel Install-Module -Name ImportExcel via admin

# Import the module
Import-Module -Name ImportExcel

#Get current running folder. 
$currentFolder = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Specify the path to the Excel file
$excelFilePath = "$currentFolder\salrylist.xlsx"

# Read the Excel file
$excelData = Import-Excel -Path $excelFilePath

#$properties = $excelData | Get-Member -MemberType NoteProperty |Select-Object -ExpandProperty Name 
$properties = $excelData | Get-Member -MemberType NoteProperty | Sort-Object { [int]($_.Name.Split('.')[0]) } |Select-Object -ExpandProperty Name 
write-host $properties

# Process each rowcls

foreach ($row in $excelData) {

   $ID = $row.'2.员工 ID'

   if ($ID -ne $null)
   {

   $TxtFilePath = "$currentFolder\$ID.txt"
   $pdfFilePath = "$currentFolder\$ID.pdf"

   Write-Output "********************工资清单*****************" | Out-File -FilePath $TxtFilePath -Force

   foreach ($property in $properties) {

   Write-Output $property : $row.$property "`r`n" | Out-File -FilePath $TxtFilePath -Append -NoNewline

   }
   }

#Convert the text file to a PDF using wkhtmltopdf
Start-Process -FilePath "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" -ArgumentList $TxtFilePath , $pdfFilePath -WindowStyle Hidden

}

#Delete all txt files
Start-Sleep -Seconds 10
Remove-Item -Path $currentFolder\*.txt -Force

