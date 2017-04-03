<#
Office365 batch user activation
Verion 0.1
OCWS
BSD 3-Clause License
#>

param([string]$userCsv = "C:\Sources\users.csv",[string]$MSOLUser = $null,[string]$MSOLPassword = $null)

Import-Module msonline

#Logging Function
Function Write-Log {
    Param ([string]$User="",[string]$Type="",[string]$Content)
    $time = get-date -format "yyyy-MM-dd HH:mm:ss[ffffff]"
    [string]$Path=$PSScriptRoot+"\ActivationOffice365-"+$date+".log"
    if ($User -eq "") {
        $Output = $time + ", " + $Type + " - " + $Content
    }
    else {
        $Output = $time + ", " + $Type + " - " + $User + " : " + $Content
    }
   Add-content -encoding utf8 -Path $Logfile -value $Output -ErrorAction SilentlyContinue
   write-host $Output
}

#Office365 connection
try {
    if($MSOLUser -ne $null -and $MSOLPassword -ne $null)
    {
        $secpasswd = ConvertTo-SecureString $MSOLPassword -AsPlainText -Force
        $MsolCred = New-Object System.Management.Automation.PSCredential ($MSOLUser,$MSOLPassword)
        Connect-MsolService -cred $MsolCred
        Write-log -Type "Information" -Contenu "Connecté au service Office 365"   
    }
    else
    {
        $MsolCred = Get-Credential -Message "Identifiant Office 365"
        Connect-MsolService -cred $MsolCred
        Write-log -Type "Information" -Contenu "Connecté au service Office 365"
    }
    }
catch {
    $errorMessage = $_.Exception.Message
    Write-log -Type "Error" -Contenu $errorMessage
    Exit
    }

#Chargement du fichier utilisateur
try{
    $users = $null
    $users = Import-Csv $userCsv
    $count = $users.count
    #Write-Host "Nombre d'utilisateurs dans le fichier CSV = " $count
    Write-log -Type "Information" -Contenu "Nombre d'utilisateurs dans le fichier CSV = $count"
}
catch{
    $errorMessage = $_.Exception.Message
    Write-log -Type "Error" -Contenu $errorMessage
    Exit
}

foreach ($user in $users)
{
    New-MsolUser -UserPrincipalName $user.UserPrincipalName -DisplayName $_.Nom_Complet -FirstName $_.Prenom -LastName $_.Nom -MobilePhone $_.tel_mobile -Office $_.Bureau -PhoneNumber $_.Tel_bureau -Title $_.Fonction -UsageLocation FR -Country $_.Pays
}