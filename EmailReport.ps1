#email report
#first grab info from excel
#INFORMATIONAL
#$date = Get-Date -Format "MMdd"
$date = Get-Date -Format "MMddyyyy"
Install-Module ImportExcel -Scope CurrentUser
Import-Module ImportExcel

#region inboxstuff for getting informational alerts

$outlook = new-object -com Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$namespace.Folders.Item('IT NOC')

$Email = $namespace.Folders.Item('IT NOC').Folders.Item('InBox').Items

#$EmailSCOM = $namespace.Folders.Item('IT NOC').Folders.Item('Alerts').Folders.Item('SCOM Alerts').Items

$countyellow = $Email | Where-Object {$_.Categories -eq "Yellow category"}

$ArrayForReport = @()
$BodyForReport = @()



foreach ($yallow in $countyellow) {
    
    $ArrayForReport += $yallow.Subject + ' '
    $BodyForReport += $yallow.Body + ' '
   }


$ArrayForReport

$DateForReports = @()

foreach ($dayte in $countyellow){

   
    $DateForReports += $dayte.ReceivedTime | Get-Date -Format 'MM/dd'

}


$DateForReports

$PlusThreeDays = @()

Foreach ($alicia in $DateForReports){

    
    $alicia = $alicia -as [DateTime]
    
    $PlusThreeDays += $alicia.AddDays(3) | Get-Date -Format 'MM/dd'
    
    }

$PlusThreeDays





$ArrayForReport | Export-Excel -Path ".\$($date).xlsx" -WorksheetName "INFORMATIONAL" -StartColumn 2 -StartRow 2

$PlusThreeDays | Export-Excel -Path ".\$($date).xlsx" -WorksheetName "INFORMATIONAL" -StartColumn 1 -StartRow 2

#endregion 


$Escalations = Import-Excel -Path "$($date).xlsx" -WorkSheetName "ESCALATIONS" | ConvertTo-Html -Fragment -PreContent "<h2>ESCALATIONS</h2>" -PostContent " "
$RequiresAttention = Import-Excel -Path "$($date).xlsx" -WorkSheetName "REQUIRES ATTENTION" | ConvertTo-Html -Fragment -PreContent "<h2>REQUIRES ATTENTION</h2>" -PostContent " "
$Informational = Import-Excel -Path "$($date).xlsx" -WorkSheetName "INFORMATIONAL" | ConvertTo-Html -Fragment -PreContent "<h2>INFORMATIONAL</h2>" -PostContent " "
$ItemsCompleted = Import-Excel -Path "$($date).xlsx" -WorkSheetName "ITEMS COMPLETED" | ConvertTo-Html -Fragment -PreContent "<h2>ITEMS COMPLETED</h2>" -PostContent " "

$header = @"
<style>
TABLE {border-width: 2px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 2px; padding: 10px; border-style: solid; border-color: black; background-color: powderblue;}
TD {border-width: 2px; padding: 20px; boder-style: solid; border-color: black;}
</style>
"@




$EmailReport = ConvertTo-Html -Head $header -Body "$Escalations $RequiresAttention $Informational $ItemsCompleted"

$EmailReport | Out-File -FilePath "$($date).html"

#Invoke-Item .\0122.html

#creates email draft and loads it

$ol = New-Object -comObject Outlook.Application
$mail = $ol.CreateItem(0)
$emaildate = Get-Date -Format "MM/dd/yyyy"
$mail.Subject = "Transition Report $($emaildate)"
$mail.To = "SidleyNOC@sidley.com"
$mail.CC = "rgarcia@sidley.com"
$mail.HTMLBody= "$EmailReport"
$mail.save()
$inspector = $mail.GetInspector
$inspector.Display()


#The above works without issue except getting a few extra zeros for dates / now we need move it into its own folder
#check the first two digits of the file name and move to the appropriate folder

#need to reinsert the original date values / add conditional formatting to properly show date

$huhhuh | Export-Excel -Path ".\$($date).xlsx" -WorksheetName "INFORMATIONAL" -StartColumn 1 -StartRow 2 

#region reformat and save
# Specify the path to the Excel file and the WorkSheet Name
#$FilePath = "C:\Users\$($env:USERNAME)\Desktop\Transition Reports\$($date).xlsx"
# Create an Object Excel.Application using Com interface
#$objExcel = New-Object -ComObject Excel.Application
# Disable the 'visible' property so the document won't open in excel
#$objExcel.Visible = $false
# Open the Excel file and save it in $WorkBook
#$WorkBook = $objExcel.Workbooks.Open($FilePath)
#$ws = $WorkBook.Worksheets[3]
#$ws.Columns.Item('a').NumberFormat = "m/d"
#$WorkBook.SaveAs($filepath)
#$objExcel.Quit()
#endregion 


$regexdate = "$($date).xlsx"

Switch -regex ($regexdate) {
    "^01"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\01_January}
    "^02"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\02_February}
    "^03"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\03_March}
    "^04"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\04_April}
    "^05"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\05_May}
    "^06"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\06_June}
    "^07"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\07_July}
    "^08"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\08_August}
    "^09"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\09_September}
    "^10"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\10_October}
    "^11"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\11_November}
    "^12"
    {Move-Item ".\$($date).xlsx" -Destination .\2021\12_December}
  
  }
