<#
Office365 batch user activation
Verion 0.1
OCWS
BSD 3-Clause License
#>

param([string]$userCsv = "C:\Sources\users.csv")

Import-Module msonline

#Office365 session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session 

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

