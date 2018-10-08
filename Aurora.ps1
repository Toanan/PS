<#
.Synopsis
   Migrate document to Sharepoint-Online saving and adding metadata
   *RunTime flow:
        1-load-AEnvironment
        2-Connect-ASPO
        3-Get-ALibrary
.DESCRIPTION
   Migrate document from local/FS to Sharepoint-Online saving and adding metadata
   Require a CSV file:
        -Delimliter ","
        -Header {
            'Created Date'
            'Modified Date'
            'Author'
        }
+++++++++++++++++++++/!\NECESSITE UN PARAMETRAGE AVEC LE CREDENTIAL MANAGER /!\+++++++++++++++++++++
.EXAMPLE
   Exemple d’usage de cette applet de commande
.EXAMPLE
   Autre exemple de l’usage de cette applet de commande
#>
function Start-Aurora 
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Targetted SharePoint Online WebSite to upload files to
        [Parameter(Mandatory=$true,Position=0)][alias("Site","SPOSite")][string]$sUrl,
        # Targetted SharePoint Online Library to upload files to
        [Parameter(Mandatory=$false,Position=1)][alias("Library","List")][string]$sLib
    )

    Begin{
    
    #Hello World
    Write-Verbose "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    Write-Verbose "++++++++++++++++++++++++ Program Starts +++++++++++++++++++++++"
    Write-Verbose "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    Write-Verbose (Get-Date) 
    
    }Process{

    load-AEnvironment
    
    }End{
    
    
    }
}

<#
.Synopsis
   Load ExecutionPolicy and ensure module SharePointPnPPowerShellOnline is loaded
.EXAMPLE
   load-AEnvironment
#>
function load-AEnvironment
{
    #Configuring ExecutionPolicy
    Write-Verbose "#Checking-ExecutionPolicy"
    try{

        Set-ExecutionPolicy RemoteSigned -EA Stop
        Write-Verbose "#Ok-ExecutionPolicy"   
    
    }catch [System.UnauthorizedAccessException] {
        
        $err = "+/!\++You need to run Powershell as an Admin to launch this program++/!\+"
        Write-Host $err -ForegroundColor Red
    
    }catch{
        
        $err = $error[0].exception.message
        Write-Host $err -ForegroundColor Red
    
    }   
    #Loading module
    Write-Verbose "#Importing-Module"
    $oModule = (Get-module)
    if (-not $oModule.name.Contains("SharePointPnPPowerShellOnline")){
        
        try{

            Import-Module "SharePointPnPPowerShellOnline" -EA Stop
            Write-Verbose "#Ok-Module" 
        
        }catch{
            
            $err = $error[0].exception.message
            Write-Host $err -ForegroundColor Red
        
        }
    }
    
    Connect-ASPO($sUrl)
}

<#
.Synopsis
   Launch the connexion to SPOSite
.EXAMPLE
   Connect-ASPO -sUrl $sUrl
#>
function Connect-ASPO([string]$sUrl)
{
    #Initialising SharePoint Online Connexion
    Write-Verbose "#Checking-SharepointOnline"       
    try{

        Connect-PnPOnline -Url $sUrl -EA Stop
        Write-Verbose "#Ok-SharePointOnline"
    
    }catch{

        $err = $error[0].exception.message
        Write-Host $err -ForegroundColor Red
    
    }

    $a = Get-ALibrary($sLib)
    
    Write-Host $a.itemCount "Items in $sLib"
}

<#
.Synopsis
   Print the SPOLibrary itemCount
.EXAMPLE
   Get-ALibrary -sLib $sLib
#>
function Get-ALibrary([string]$sLib)
{

    #Getting the library object
    Write-Verbose "#Checking-Library"       
    try{

        $oList = Get-PnPList -Identity $sLib -EA Stop
        Write-Verbose "#Ok-Library : $sLib"
        return $oList

    }catch{

        $err = $error[0].exception.message
        Write-Host $err -ForegroundColor Red
        return false

    } 
}



<#

#Ce script est un exemple de migration de document et de création d'ensemble de document dans sharepoint.
#Les élément dans la propriété -Values dépendent de la bibliothèque sur SharePoint et doivent etre présent dans le CSV

Set-ExecutionPolicy RemoteSigned
Update-Module -Name SharePointPnPPowerShellOnline
#Install-Module -Name SharePointPnPPowerShellOnline
$UserCredential = Get-Credential
Connect-PnPOnline https://quanticcloud.sharepoint.com/sites/hubqtc/quanticsupport -Credential $UserCredential	

#Variables à modifier
$ContentType = '0x0120D52000413AEE2C2594D34D9F3648C3A08112A7'
$TargetList = "Affaires"
$LogFile = 'C:\Users\Agalland\OneDrive - QUANTIC SUPPORT\Clients\Quantic interne\Migration SP\Affaires\log\MigrationAffaireError.txt'
$Csv = Import-Csv -Path 'C:\Users\Agalland\OneDrive - QUANTIC SUPPORT\Clients\Quantic interne\Migration SP\Affaires\final export affaire copie.csv' -Delimiter ';' -ErrorAction Continue

#Log début de migration
$Time = Get-Date
"$($Time) Début de la migration" | Out-File $LogFile -Append

foreach ($row in $Csv) {

#Check Type = Affaire
    
    If ($row.Type.StartsWith("Aff")){
#Create The Document Set
        Add-PnPDocumentSet -List $TargetList -ContentType $ContentType -Name $row.Affaire 
#Save the last Document Set created in a variable
        $CurrentItem = Get-PnPListItem -list $TargetList | Where-Object {$_["HTML_x0020_File_x0020_Type"] -eq 'Sharepoint.DocumentSet'} | Sort-Object Id -Descending | Select-Object -First 1 
#Edit the last document set created using the variable
        $ffolder = "$($CurrentItem.FieldValues.Title) - $($CurrentItem.FieldValues.ID)"
#Try to modify the Current Affaire        
        try {
            Set-PnPListItem -List $TargetList -Identity $CurrentItem -Values @{FileLeafRef="$($ffolder)";Author=$row.'SPuser';Editor=$row.'SPuser';AssignedTo=$row.'SPuser';Etat=$row.'_Status';Created=$row.'Date';Modified=$row.'Date';Comptes=$row.'CompteID'} -Ea 'Stop'
        }
#catch : write error to the log file and the shell       
        catch { 
            $Time = Get-Date
            "$($Time) -- $($CurrentItem.FieldValues.Title) --> $($_)" | Out-File $LogFile -Append
            write-host "$($CurrentItem.FieldValues.Title) --> $($_)" -ForegroundColor Red
        }
        
        $Folder = "$($TargetList) / $($CurrentItem.fieldValues.FileLeafRef)"
        Write-Host "L'affaire " $($Folder) "à été créée" -ForegroundColor Yellow
    }
    
#Check Type = Document
    Elseif ($row.Type.StartsWith("Doc")){
#Try to copy file to the curent Document Set
    try {
        $file = Add-PnPFile -Path $row.Path -Folder $Folder -Values @{Document=$row.'TypeDeDocument';Created=$row.'ModifiedDoc';Modified=$row.'ModifiedDoc';Author=$row.'SPuser';Editor=$row.'SPUser'} -Ea 'Stop'
    }
#catch : write error to the log file and the shell    
    catch {
        $Time = Get-Date
        "$($Time) -- $($file.Name) --> $($_)" | Out-File $LogFile -Append
        Write-Host "$($file.Name) --> $($_)" -ForegroundColor Red
    }

    write-host "le document" $($file.Name) "à été copié" -ForegroundColor Green

    }
}

$Time = Get-Date
"$($Time) Fin de la migration" | Out-File $LogFile -Append


#Set-PnPListItem -List $TargetList -Identity 6133 -Values @{FileLeafRef="PC PORTABLES - 6133";Comptes="696";AssignedTo="18"} -Ea 'Stop'
#>