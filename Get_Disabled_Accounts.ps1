##################################################################
# Help                                                                                                        #
##################################################################
<#  
    Nom: Get_Disabled_accounts.ps1
    Descripttion: 
    Get Disabled Accounts, sort them by OU and Export by Excel or CSV
    .EXAMPLE
    Get_Disabled_Accounts.ps1
    Version 1.3
#>
##################################################################
# Main                                                                                                       #
##################################################################
# Get file path 
$RootFolder = Split-Path -Path $MyInvocation.MyCommand.Definition
# Test if active directory module is installed
if((Get-Module ActiveDirectory) -eq $null){
    try{
        Import-Module ActiveDirectory
    }catch{
        Write-Host "The execution computer doesn't have ActiveDirectory Powershell Module. The script can't continue." -ForegroundColor Red
        return
    }
}
# Check if excel is installed
try{
    $objExcel = new-object -comobject excel.application
    Write-Host "Excel is installed on this Computer, disabled Users will be export in a fashioned excel file."
    $ExcelTest = $true
}catch{
    Write-Host "Excel is not installed on this Computer, disabled Users will be export in a plain old CSV file."
    $ExcelTest = $false
}
# Get the current date for file
$Date = Get-Date -Format ddMMyyyy
# check if excel is available
if($ExcelTest){
    # set path of excel file
    Write-Host "Create Excel file"
    $ExcelPath = "$RootFolder\DisabledAccounts_$Date.xlsx"
    if (Test-Path $ExcelPath) {
        $finalWorkBook = $objExcel.WorkBooks.Open($ExcelPath)
       # select the worksheet we working on 
        $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
        # set a name for the worksheet 
        $finalWorkBook.Worksheets.Item(1).Name = "DisabledAccounts"
    }else{
        # creation of a new file
        $finalWorkBook = $objExcel.Workbooks.Add()
        $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
        $finalWorkBook.Worksheets.Item(1).Name = "DisabledAccounts"
    }
    #$objExcel.Visible =$true
    Write-Host "Create header"
   # write the first line
    $finalWorkSheet.Cells.Item(1,1) = "SamAccountName"
    # set the text to Gras
    $finalWorkSheet.Cells.Item(1,1).Font.Bold = $True
    $finalWorkSheet.Cells.Item(1,2) = "FirstName";
    $finalWorkSheet.Cells.Item(1,2).Font.Bold = $True
    $finalWorkSheet.Cells.Item(1,3) = "LastName"
    $finalWorkSheet.Cells.Item(1,3).Font.Bold = $True
    $finalWorkSheet.Cells.Item(1,4) = "LastLogonDate"
    $finalWorkSheet.Cells.Item(1,4).Font.Bold = $True
    $finalWorkSheet.Cells.Item(1,5) = "DistinguishedName"
    $finalWorkSheet.Cells.Item(1,5).Font.Bold = $True
}else{
    # Creation of the CSV file
    $CSVPath = "$RootFolder\DisabledAccounts_$Date.csv"
    # Creation of the header
    $Header = "SamAccountName;FirstName;LastName;LastLogonDate;DistinguishedName"
    # writing of header
    $Header | Out-File -FilePath $CSVPath
}
Write-Host "Rrieving data..." -ForegroundColor Green
# getting all the disbled accounts
$ListUser = Get-ADUser -Filter {Enabled -eq $false} -Properties * | Select CanonicalName, CN, DistinguishedName, SamAccountName, GivenName, SurName, LastLogonDate, `
@{Name="OU";Expression={(($_.CanonicalName).Substring($_.CanonicalName.IndexOf("/")+1)).replace("/$($_.CN)","")}} `
| Sort-Object OU
Write-Host "Writing data..." -ForegroundColor Green
# we start at the second line (the first is for the Header)
$FinalExcelRow = 2
# select of a color for highlight
$ColorIndex = 41
# getting the OU of the first user
if($ListUser.Count -gt 0){
    $OU = $ListUser[0].OU
}
$i = 0
#loop on every users
ForEach($User in $ListUser){
    #progress bar for every users
    $PercentComplete = [System.Math]::Round($($i*100/($ListUser.Count)),2)
    Write-Progress -Activity "Exporting data to Excel" -status "Effectué : $PercentComplete %" -percentcomplete $($i*100/($ListUser.Count))
    $i++
    if($ExcelTest){
        # if OU has changed we change the color
        if($OU -ne $User.OU){
            $ColorIndex++
            if($ColorIndex -ge 56){
                $ColorIndex = 3
            }
        }
       # getting the OU of the current user
        $OU = $User.OU
        #Store the differents values
        $finalWorkSheet.Cells.Item($FinalExcelRow,1) = $User.SamAccountName
        #attribute each different color for the values
        $finalWorkSheet.Cells.Item($FinalExcelRow,1).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,2) = $User.GivenName
        $finalWorkSheet.Cells.Item($FinalExcelRow,2).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,3) = $User.SurName
        $finalWorkSheet.Cells.Item($FinalExcelRow,3).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,4) = $User.LastLogonDate
        $finalWorkSheet.Cells.Item($FinalExcelRow,4).Interior.ColorIndex = $ColorIndex
        $finalWorkSheet.Cells.Item($FinalExcelRow,5) = $User.DistinguishedName
        $finalWorkSheet.Cells.Item($FinalExcelRow,5).Interior.ColorIndex = $ColorIndex
       $FinalExcelRow++
    }else{
        $Result = $User.SamAccountName+";"+$User.GivenName+";"+ `                            $User.SurName+ ";"+$User.LastLogonDate+";"+$User.DistinguishedName
        $Result | Out-File -FilePath $CSVPath -Append 
    } 
}
Write-Host "Saving data and closing Excel." -ForegroundColor Green
if($ExcelTest){
    # Select used cells
    $UR = $finalWorkSheet.UsedRange
   # Auto adjust of the column size    
    $null = $UR.EntireColumn.AutoFit()
    if (Test-Path $ExcelPath) {
        # we save if the file exist
        $finalWorkBook.Save()
    }else{
        # ifnot we give a name
        $finalWorkBook.SaveAs($ExcelPath)
    }
    # closing the file
    $finalWorkBook.Close()
}
# The Excel process used to process the transaction is terminated
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)