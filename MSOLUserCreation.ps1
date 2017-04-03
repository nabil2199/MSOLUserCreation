<#
Office365 batch user activation
Verion 0.1
OCWS
BSD 3-Clause License
#>

param([string]$userCsv = "C:\Sources\users.csv")

#Open Office365 session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session 

