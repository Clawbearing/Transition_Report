$date = Get-Date -Format "MMddyyyy"
Set-Location "C:\Users\$($ENV:username)\Desktop\Transition Reports\"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show('Please select your most recent transition report')
#region beginning directory

$FileBrowsertuu = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = "C:\Users\$($ENV:username)\Desktop\Transition Reports\2021\"}
$null = $FileBrowsertuu.ShowDialog()


#endregion 

#$transitionfolder = Get-Item -Path "C:\Users\ecowgill\Desktop\Transition Reports"
#$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
#$null = $FileBrowser.ShowDialog()
#$recenttransitionreport = $FileBrowser.FileName
$recenttransitionreport = $FileBrowsertuu.FileName


Copy-Item -Path $recenttransitionreport -Destination "C:\Users\$($env:USERNAME)\Desktop\Transition Reports\$($date).xlsx"

#instead of line 9 - copy from template and then just insert the items completed page
Copy-Item -Path .\Template\TransitionTemplate.xlsx -Destination "C:\Users\$($env:USERNAME)\Desktop\Transition Reports\$($date).xlsx"


#region insert ITEMS COMPLETED page from recent transition report that was selected at the top

#can we do conditional formatting to get to display dates
#Import-Excel -Path $recenttransitionreport -WorksheetName "INFORMATIONAL" | Export-Excel -Path ".\$($date).xlsx" -WorksheetName "INFORMATIONAL"

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
#$objExcel.DisplayAlerts = $false;
#$objExcel.ActiveWorkbook.SaveAs($xlsFile);
#$WorkBook.SaveAs($filepath)
#$objExcel.Quit()
#endregion 


#endregion 

#Start-Process -FilePath ".\$($date).xlsx" -WindowStyle Maximized

# Specify the path to the Excel file and the WorkSheet Name
$FilePath = "C:\Users\$($env:USERNAME)\Desktop\Transition Reports\$($date).xlsx"
# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $true
$WorkBook = $objExcel.Workbooks.Open($FilePath)
$WorkBook
$objExcel.DisplayAlerts = $true;











