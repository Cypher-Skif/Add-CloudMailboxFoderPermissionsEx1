param (
    [Parameter(Mandatory=$true)]
    [string]$GrantAccessTo

)
<#----------------Variables that should be set-----------------#>
$accessLevel = "Reviewer" #access level
$targetUser = <userEmail> #whose folders will be shared
$ExcludedFolder = "Private" #name of folder what needs to be excluded
<#-------------------------------------------------------------#>

<#----------------Connect to exchange online-----------------#>
$cloudCredentials = Get-Credential
Connect-ExchangeOnline -Credential $cloudCredentials -Prefix cl 
<#-----------------------------------------------------------#>

#Get list of folders from inbox
$folders = Get-clMailboxFolderStatistics $targetUser -folderscope inbox | Select name, folderpath 

#Grant access to all folders in Inbox except of $ExcludedFolder
foreach ($folder in $folders) {
    if ($folder.name -ne $ExcludedFolder){
        $grantedFolder = $GrantAccessTo + ":\" + $folder.name
        try{
		Add-clMailboxFolderPermission -Identity $grantedfolder -user $GrantAccessTo -AccessRights $accessLevel
		write-host "Permission to" $grantedFolder "granted to" $grantaccessto}
	catch{
		Write-Error "Can't grant access"
	}
        }else{
            write-host "Excluded folders" $folder.name
        }
    
}


