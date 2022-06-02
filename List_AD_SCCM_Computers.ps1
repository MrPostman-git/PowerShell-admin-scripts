################################################################################################################################
#                                                                                                                              #
################################################################################################################################
<#                                                                                                                             #
##                                                                                                                             #
##   Name: List_AD_SCCM_Computers                                                                                              #
##   Description:                                                                                                              #
##   Get ADComputer && Get CM-Device                                                                                           #
##   .DESCRIPTION                                                                                                              #
##   Récupérer les postes de SCCM et de l'AD, Les classe dans l'ordre et l'exporte en deux colonnes                            #
     dans un fichier Excel ou CSV                                                                                              #
                                                                                                                               #
     Un AD et configuration Manager est nécéssaire pour l'utilisation de ce script                                             #
     modifier la variable SCCM_Server_Name par le nom de votre serveur                                                         #
##                                                                                                                             #
##   Version : 1.3                                                                                                             #
##   Auteur : Mr Postman                                                                                                       #
##                                                                                                                             #
##                                                                                                                             #
##                                                                                                                             #
##                                                                                                                             #
#>                                                                                                                             # 
################################################################################################################################
#                                                                                                                              #
################################################################################################################################
# Récupération du chemin du fichier du script
$RootFolder = Split-Path -Path $MyInvocation.MyCommand.Definition


# Test de la présence du module Active Directory
if((Get-Module ActiveDirectory) -eq $null){
    try{
        Import-Module ActiveDirectory
    }catch{
        Write-Host "The execution computer doesn't have ActiveDirectory Powershell Module. The script can't continue." -ForegroundColor Red
        return
    }
}


# Test de la présence du module Configuration Manager
if((Get-Module ConfigurationManager) -eq $null){
    try{
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    }catch{
        Write-Host "The execution computer doesn't have the Configuration Manager Powershell Module. The script can't continue." -ForegroundColor Red
        return
    }
}


# Création de l'objet application Excel sinon on réalise un export au format CSV
try{
    $objExcel = new-object -comobject excel.application
    Write-Host "Excel is installed on this Computer, a list of all the computers will be export in a fashioned excel file."
    $ExcelTest = $true
}catch{
    Write-Host "Excel is not installed on this Computer, a list of all the computers will be export in a plain old CSV file."
    $ExcelTest = $false
}
# Génération de la date du jour pour le nom du fichier d'export
$Date = Get-Date -Format ddMMyyyy
# Si Excel est disponible
if($ExcelTest){
    # Génération du chemin du fichier d'export
    Write-Host "Create Excel file"
    $ExcelPath = "$RootFolder\ListComputers_$Date.xlsx"
   # Si le fichier Existe on l'ouvre
    if (Test-Path $ExcelPath) {
        $finalWorkBook = $objExcel.WorkBooks.Open($ExcelPath)
       # On choisi l'onglet sur lequel on travaille
        $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
        # Donne un nom au à l'onglet
        $finalWorkBook.Worksheets.Item(1).Name = "ListComputers"
    }else{
        # Création d'un nouveau fichier
        $finalWorkBook = $objExcel.Workbooks.Add()
        $finalWorkSheet = $finalWorkBook.Worksheets.Item(1)
        $finalWorkBook.Worksheets.Item(1).Name = "ListComputers"
    }
    #$objExcel.Visible =$true
    Write-Host "Create header"
   # Rempli la première ligne
    $finalWorkSheet.Cells.Item(1,1) = "Nom postes AD"
    # Met le texte en gras
    $finalWorkSheet.Cells.Item(1,1).Font.Bold = $True

    $finalWorkSheet.Cells.Item(1,3) = "Nom postes SCCM"
    $finalWorkSheet.Cells.Item(1,3).Font.Bold = $True
    
}else{
    # Création du chemin du fichier CSV
    $CSVPath = "$RootFolder\DisabledAccounts_$Date.csv"
    # Création du header du fichier CSV
    $Header = "Nom  Postes AD;LastLogonDate;LastName;FirstName;DistinguishedName"
    # Ecriture du header
    $Header | Out-File -FilePath $CSVPath
}
Write-Host "Rrieving AD computers..." -ForegroundColor Green
# Récupération des ordianateur de l'AD
$ListADComputer = Get-ADComputer -Filter * -Property * | Sort-Object Name -Descending |
Select-Object Name | 
Where-Object {$_.name -NotMatch "-TL" -And $_.name -NotMatch "VIRT" -And $_.name -NotMatch "SERV"-And $_.name -NotMatch "SRV" -And $_.name -NotMatch "PCL"}
Write-Host "Writing data..." -ForegroundColor Green
# On commence à la seconde ligne (la 1ère est consacrée au Header)
$FinalExcelRow = 2
# Choix d'une couleur pour surligner la ligne
$ColorIndex = 00


$i = 0
#On boucle sur chaque ordianateur
ForEach($Computer in $ListADComputer){
    #On affiche une barre de progression montrant le nombre d'ordianateur déjà traités
    $PercentComplete = [System.Math]::Round($($i*100/($ListADComputer.Count)),2)
    Write-Progress -Activity "Exporting AD computers to Excel" -status "Effectué : $PercentComplete %" -percentcomplete $($i*100/($ListADComputer.Count))
    $i++
    
    
        if($ExcelTest){
      
        #On stocke les différentes valeurs
        $finalWorkSheet.Cells.Item($FinalExcelRow,1) = $Computer.Name
        #On attribut la couleur définit plus haut pour la case concerné
        $finalWorkSheet.Cells.Item($FinalExcelRow,1).Interior.ColorIndex = $ColorIndex
        
        
        # On incrémente le numéro de la ligne en cours d'écriture
       $FinalExcelRow++
    }else{
        $Result = $Computer.Name+";"+$Computer.GivenName+";"+ `                            $Computer.SurName+ ";"+$Computer.LastLogonDate+";"+$Computer.DistinguishedName
        $Result | Out-File -FilePath $CSVPath -Append 
    } 
} 



#Affiche les postes SCCM

# on importe le module de ConfigurationManager.psd1
#Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"


#Connexion au lecteur du site
$SCCM_Server_Name='SRV-SCCM-PRI'
New-PSDrive -Name 'PRI' -PSProvider CMSite -Root $SCCM_Server_Name| Out-Null


#on définie l'emplacement actuelle par le code du site 
Set-Location 'PRI:\'


Write-Host "Rrieving SCCM computers..." -ForegroundColor Green
# Récupération des ordianateur d'SCCM
$ListSCCMComputer = Get-CMDevice  | Sort-Object Name -Descending | Select-Object Name, LastLogonDate |
Where-Object {$_.name -NotMatch "VIRT" -And $_.name -NotMatch "SERV"-And $_.name -NotMatch "SRV" -And $_.name -NotMatch "PCL"}
Write-Host "Writing data..." -ForegroundColor Green
# On commence à la seconde ligne (la 1ère est consacrée au Header)
$FinalExcelRow = 2
# Choix d'une couleur pour surligner la ligne
$ColorIndex = 00


$i = 0
#On boucle sur chaque ordinateur
ForEach($Computer in $ListSCCMComputer){
    #On affiche une barre de progression montrant le nombre d'ordinateur déjà traités
    $PercentComplete = [System.Math]::Round($($i*100/($ListSCCMComputer.Count)),2)
    Write-Progress -Activity "Exporting SCCM computers to Excel" -status "Effectué : $PercentComplete %" -percentcomplete $($i*100/($ListSCCMComputer.Count))
    $i++
    
    
    
        if($ExcelTest){
       
        #On stocke les différentes valeurs
        $finalWorkSheet.Cells.Item($FinalExcelRow,3) = $Computer.Name
        #On attribut la couleur définit plus haut pour la case concerné
        $finalWorkSheet.Cells.Item($FinalExcelRow,3).Interior.ColorIndex = $ColorIndex
        
       
        # On incrémente le numéro de la ligne en cours d'écriture
       $FinalExcelRow++
    }else{
        $Result = $Computer.Name+";"+$Computer.GivenName+";"+ `                            $Computer.SurName+ ";"+$Computer.LastLogonDate+";"+$Computer.DistinguishedName
        $Result | Out-File -FilePath $CSVPath -Append 
    } 
 
}
Write-Host "Saving data and closing Excel." -ForegroundColor Green
if($ExcelTest){
    # Sélectionne les cellules utilisées
    $UR = $finalWorkSheet.UsedRange
   # Auto ajustement de la taille de la colonne    
    $null = $UR.EntireColumn.AutoFit()
    if (Test-Path $ExcelPath) {
        # Si le fichier existe déjà, on le sauvegarde
        $finalWorkBook.Save()
    }else{
        # Sinon on lui donne un nom de fichier au moment de la sauvegarde
        $finalWorkBook.SaveAs($ExcelPath)
    }
    # On ferme le fichier
    $finalWorkBook.Close()
}
# Le processus Excel utilisé pour traiter l'opération est arrêté
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

Set-Location $RootFolder