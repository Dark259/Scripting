#-------------------------------------------
# This script is meant to be run automatically.
# This will go through all computers in the Domain Computers OU and disable.
# Those over 30 days for last login will have a notice sent to the user.
# Those over 45 days for last login will have a final notice sent to the user.
# (If a user is not found, the Service Desk email will receive all notices and the user will need to discovered and contacted manually.)
# Those with a last login date older than 2 months will be disabled and moved to the disabled OU. A notice will be sent to the user with the Service Desk cc'd.
#
# After it is in the disabled OU for 1 months it will be up for deletion.
# The script will check the description which is set whenever the script is to move a computer to the disabled OU.
# If the description date is older than 1 month, the BitLocker Key and LAPS password will be backed up and the computer will be deleted from AD.
# In this sitation an email to the user with the manager and Service Desk cc'd will be sent in essence treating the computer as a lost asset looking to be retrieved at that point.
#
# This script also will move enabled computers out of disabled and disabled computers into
# enabled as a backup method to keep the OUs clean and organized. Along with that it will move
# computers to their respective OUs based on naming convention.
#-------------------------------------------

#Requires -Modules ActiveDirectory, configurationmanager, mgraph

Import-Module ConfigurationManager
New-PSDrive -Name "" -PSProvider "CMSite" -Root "" -Description "Primary site" #Setup for your SCCM site
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All"

#Setup variables

#Functions
function Append-ChangeLog {
    param (
        [parameter(Mandatory=$false)]
        [string]$LogPath="C:\Scripts\ADCleanup\$(Get-Date -Format yyyy-MM-dd)-ComputerCleanup.log",
        [parameter(Mandatory=$true)]
        [string]$FilePath,
        [parameter(Mandatory=$true)]
        [string]$ActionType,
        [parameter(Mandatory=$false)]
        [string]$Destination,
        [parameter(Mandatory=$false)]
        [string]$UserName,
        [parameter(Mandatory=$false)]
        [string]$ComputerName,
        [parameter(Mandatory=$false)]
        [string]$BitLockerKey='N/A',
        [parameter(Mandatory=$false)]
        [string]$LAPSPassword='N/A'
    )
    $out=[pscustomobject]@{
        Date=(Get-Date);
        Action=$ActionType;
        ComputerName = $ComputerName;
        Destination=$Destination;
        BitLockerKey=$BitLockerKey;
        LAPSPassword=$LAPSPassword
    } 
    $out | Export-Csv -Append -Path $FilePath

    $out | Out-File -Append -FilePath $LogPath
}

function Create-ComputerLists {
    param (
        [parameter(Mandatory=$false)]
        [string]$Disabled = "" #Disabled Computers OU
    )
    $Exceptions = @() #To be added in the case that we have exceptions within domain computers.
    $lists =@()
    $placeholder = foreach ($computer in $(Get-ADComputer -Filter * -SearchBase "" -Properties name, lastlogondate, description)) {
        $desc=$null
        if ((-not [string]::IsNullOrWhiteSpace($computer.description)) -and ($computer.description -like "Disabled on ????-??-??")) {
            $desc=([DateTime]($computer.description).split(' ')[-1])
        }
        if(($computer.DistinguishedName -eq ("CN=$($computer.name),$Disabled")) -and ($desc  -le ((Get-Date).AddMonths(-1)))) {
                $DecommissionStatus = $TRUE
        } else {
            $DecommissionStatus = $FALSE
        }
        $list=$null
        $list+= @{
            Name = $computer.name
            DistinguishedName = $computer.DistinguishedName;
            LastLogon = $computer.lastlogondate;
            Decommission = $DecommissionStatus
            Description = $computer.description;
            Enabled = $computer.enabled;
            Destination = (Get-FolderDestination -Name $computer.name -Decommission $DecommissionStatus);
        }
        $list+=@{CorrectLocation = [bool]($list.Destination -eq (($computer.distinguishedname).replace("CN=$($computer.name),",'')))}
        $list+=@{Age = [int](Get-Age -LastLogon ($computer.LastLogonDate))}
        $lists+=[pscustomobject]$list

    }
    return $lists
}

function Create-UserLists {
    param (
        [parameter(Mandatory=$false)]
        [string]$Disabled = "" #Terminated OU
    )
    $Exceptions = @() #To be added in the case that we have exceptions within domain computers.
    $lists =@()
    $EntraLogins =@()
    Get-MgBetaUser -All -Property Mail, Displayname, SignInActivity | % {
        $EntraLogins += [pscustomobject]@{Email=$_.mail;LastSignin=(($_.signinactivity).lastsignindatetime);NonInteractive=(($_.signinactivity).LastNonInteractiveSignInDateTime)}
    }
    $OnPremPlaceholder = foreach ($user in $(Get-ADUser -Filter * -SearchBase "" -Properties name, lastlogondate, passwordneverexpires)) {
        $EntraPlaceholder = $EntraLogins | where email -EQ $user.UserPrincipalName 

        if($user.lastlogondate -le (get-date).AddDays(-45)){ #Assign on-prem stale status
            $OnPremStale = $TRUE
        }
        else {
            $OnPremStale = $FALSE
        }

        if(($EntraPlaceholder.lastsignin -le (Get-Date).AddDays(-45)) -and ($EntraPlaceholder.NonInteractive -le (Get-Date).AddDays(-45))){ #Assign Azure stale status
            $AzureStale = $TRUE
        }
        else {
            $AzureStale = $FALSE
        }
        
        $list=$null
        $list+= @{
            Name = $User.name
            DistinguishedName = $User.DistinguishedName;
            LastLogon = $User.lastlogondate;
            OnPremStale = $OnPremStale
            AzureStale = $AzureStale
            Enabled = $User.enabled;
            PasswordNeverExpires = $User.passwordneverexpires
        }
        $lists+=[pscustomobject]$list

    }
    return $lists
}

function Get-Age {
    param (
        [parameter(Mandatory=$true)]
        [DateTime]$LastLogon
    )

    $CurrentDate = [datetime](Get-Date)
    $Age = New-TimeSpan -Start $LastLogon -End $CurrentDate

    return $Age.Days
}
function Send-Emails {

    param (
        [parameter(Mandatory=$TRUE)]
        [string]$NoticeType,
        [parameter(Mandatory=$FALSE)]
        [List] $List,
        [parameter(Mandatory=$FALSE)]
        [List] $ComputerName
    )

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #0066CC;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$OmissionGroup = @() # (Get-ADGroupMember "").samaccountname Need a group of everyone
$RecipientUser = $null;
cd EXA:\\
$RecipientUser = "$(Get-CMDevice -Name $ComputerName | select -ExpandProperty lastlogonuser)@exac.com"
cd C:

    if($ComputerList){
        $Table = $ComputerList| Sort-Object Hostname | ConvertTo-Html -Head $Header
    }
    if($NoticeType -eq "Decommission" -and $ComputerList){
        $DecommissionComputersBodyString = 
        "<br>Computers deemed ready for Decommissioning:</b><br/><br/>
        $Table<br>"

        $MailMessage = @{
            To         = ""
            From       = ""
            Subject    = ""
            BodyAsHtml = $true
            Body       = $DecommissionComputersBodyString
            SMTPServer = ""
            Port       = ""
            Credential = 
            UseSsl     = $true
        }
        Write-Output "Decommission ready computers sent to freshservice as ticket..."
        #Send-MailMessage @MailMessage
    }elseif($NoticeType -eq "Unorganized"){
        $Script:Unorganized = "
        Computer:$($_.name) did not follow any known naming conventions. Due to this it was left in the upper 'Domain Computers' OU. <br/>
        Please move this computer to the correct country OU inside AD.<br/>
        <br/><br/>"
        $UnorganizedBodyString = $Script:Unorganized| Out-String
        $MailMessage = @{
            To         = ""
            From       = ""
            Subject    = ""
            BodyAsHtml = $true
            Body       = $UnorganizedBodyString
            SMTPServer = ""
            Port       = ""
            Credential = 
            UseSsl     = $true
        }
        Write-Output "$($_.name) left outside country folder."
        #Send-MailMessage @MailMessage
    }elseif($NoticeType -eq "StaleUsers"){
        $StaleUsers = 
        "<br>Computers deemed ready for Decommissioning:</b><br/><br/>
        $Table<br>"
        $MailMessage = @{
            To         = ""
            From       = ""
            Subject    = ""
            BodyAsHtml = $true
            Body       = $StaleUsers
            SMTPServer = ""
            Port       = ""
            Credential = 
            UseSsl     = $true
        }
        Write-Output "$($_.name) left outside country folder."
        #Send-MailMessage @MailMessage
    }
    if($OmissionGroup -notcontains $RecipientUser){
        if($NoticeType -eq "Initial"){
            $Script:NoticeBodyString = "
            Your computer:$($_.name) has not been seen on our systems for over 30 days. 
            Per policy, if your computer remains offline for 30 more days for a total of 60, it will be disabled from accessing any of our systems.
            To avoid  this happening, please turn on your computer, connect to VPN, and leave it on and awake for 1-2 hours at least once a month.<br/>
            If you have any questions please submit a ticket at exactech.freshservice.com or email IT as is.request@exac.com<br/>
            <br/>
            Thank you for your understanding and cooperation with our policies.<br/>
            From,<br/>
            IT Team<br/>
            <br/><br/>"
            $NoticeBodyString = $Script:NoticeBodyString | Out-String
            $MailMessage = @{
                To         = $RecipientUser
                From       = ""
                Subject    = ""
                BodyAsHtml = $true
                Body       = $NoticeBodyString
                SMTPServer = ""
                Port       = ""
                Credential = 
                UseSsl     = $true
            }
            Write-Output "$($_.name) offline 30+ days. Generating email notification..."
            write-Output $MailMessage
            #Send-MailMessage @MailMessage
        }elseif($NoticeType -eq "Final"){
            $Script:FinalNoticeBodyString = "
            Your computer:$($_.name) has not been seen on our systems for over 45 days. 
            Per policy, if your computer remains offline for 15 more days for a total of 60, it will be disabled from accessing any of our systems.
            To avoid  this happening, please turn on your computer, connect to VPN, and leave it on and awake for 1-2 hours at least once a month.<br/>
            If you have any questions please submit a ticket at exactech.freshservice.com or email IT as is.request@exac.com.<br/>
            <br/>
            Thank you for your understanding and cooperation with our policies.<br/>
            From,<br/>
            IT Team<br/>
            <br/><br/>"
            $FinalNoticeBodyString = $Script:FinalNoticeBodyString | Out-String
            $MailMessage = @{
                To         = $RecipientUser
                From       = ""
                Subject    = ""
                BodyAsHtml = $true
                Body       = $FinalNoticeBodyString
                SMTPServer = ""
                Port       = ""
                Credential = 
                UseSsl     = $true
            }
            Write-Output "$($_.name) offline 45+ days. Generating email notification..."
            #Send-MailMessage @MailMessage
        }elseif($NoticeType -eq "Disabled"){
            $Script:DisabledBodyString = "
            Your computer:$($_.name) has not been seen on our systems for over 60 days.
            Per policy, it has been disabled from accessing any of our systems.<br/>
            Please submit a ticket at exactech.freshservice.com or email IT at is.request@exac.com to have your computer assessed before it is reactivated.<br/>
            Thank you for your understanding and cooperation with our policies.<br/>
            <br/>
            From,<br/>
            IT Team<br/>
            <br/><br/>"
            $DisabledBodyString = $Script:DisabledBodyString | Out-String
            $MailMessage = @{
                To         = $RecipientUser
                From       = ""
                Subject    = ""
                BodyAsHtml = $true
                Body       = $DisabledBodyString
                SMTPServer = ""
                Port       = ""
                Credential = 
                UseSsl     = $true
            }
            Write-Output "$($_.name) disabled. Generating email notification..."
            #Send-MailMessage @MailMessage
        }
    }
    Start-Sleep -Seconds 10
}

function Disable-ADObject {
    param (
        [parameter(Mandatory=$true)]
        [string]$DistName,
        [parameter(Mandatory=$true)]
        [string]$ObjectType
    )

    if($ObjectType -eq "Computer"){

        $Description = "Disabled on $(Get-Date -Format yyyy-MM-dd))"

        try {
            Set-ADComputer $DistName -Enabled $False
            Set-ADComputer $DistName -Description "$Description"
            Append-ChangeLog -FilePath "" -ActionType "Disabled" -ComputerName $DistName
        } catch {
            Out-File -Append -FilePath error.log -InputObject "Failed to disable $DistName"
            Out-File -Append -FilePath error.log -InputObject $_
        }
    }

    if($ObjectType -eq "User"){
        try{
            Set-ADUser $DistName -Enabled $False
            Append-ChangeLog -FilePath "" -ActionType "Disabled" -UserName $DistName
        } catch {
            Out-File -Append -FilePath error.log -InputObject "Failed to disable $DistName"
            Out-File -Append -FilePath error.log -InputObject $_
        }
    }
}

function Backup-Computers {
    param (
        [parameter(Mandatory=$true)]
        [string]$DistName
    )

    try { 
        $BackupInfo = Get-ADComputer $DistName -Properties *
        $BitlockerDetails = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $BackupInfo.DistinguishedName -Properties 'msFVE-RecoveryPassword', 'distinguishedname'
        $Bitlocker = $BitlockerDetails | select msFVE-RecoveryPassword
        $LAPSPass = $BackupInfo | select -ExpandProperty ms-Mcs-AdmPwd
        if(($LAPSPass -notlike $null) -or ($Bitlocker -notlike $null)){
            $BackupCheck = Import-Csv ""
            $BackupInfo | %{
                if($Backupcheck.computername -notcontains $BackupInfo.DistinguishedName){
                    Append-ChangeLog -FilePath "" -ActionType "Backup" -ComputerName $_ -BitLockerKey $Bitlocker.ToString() -LAPSPassword $LAPSPass.ToString()
                }
            }
        }
    } catch {
        Out-File -Append -FilePath error.log -InputObject "Failed to backup $DistName"
        Out-File -Append -FilePath error.log -InputObject $_
    }
}

function Move-ADObject {
    param(
    [parameter(Mandatory=$true)]
    [string]$DistName,
    [parameter(Mandatory=$true)]
    [string]$Location
    )

    try {
        if($DistName -ne $NULL){
            Move-ADObject -Identity $DistName -TargetPath $Location
            Append-ChangeLog -FilePath "" -ActionType "Moved" -ComputerName $DistName -Destination $Location
        }
    } catch {
        Out-File -Append -FilePath error.log -InputObject "Failed to move $DistName"
        Out-File -Append -FilePath error.log -InputObject $_
    }
}

function Get-FolderDestination{

    param(
        [parameter(Mandatory=$True)]
        [String]$Name,
        [parameter(Mandatory=$True)]
        [bool]$Decommission
    )

    if($Decommission -eq $FALSE){
        if($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        elseif($Name -like "*"){
            $Destination = ""
        }
        else{
            $Destination = "No Match"
        }
    } else {
        $Destination = ""
    }
    return $Destination
}



#Script Body

$Creds = Import-Clixml C:\Scripts\Credential.xml
$Computers = Create-ComputerLists
$Users = Create-UserLists

#Disable and move accounts if stale
$UserChangeCount = ($Users | where {$_.OnPremStale -eq $TRUE -and $_.passwordNeverExpires -eq $FALSE -and $_.DistinguishedName -like "" -and $_.AzureStale -eq $TRUE -and $_.enabled -eq $TRUE}).count
$Users | where enabled -EQ $TRUE | %{
    if($UserChangeCount -le 50){
        if($_.OnPremStale -eq $TRUE -and $_.passwordNeverExpires -eq $FALSE -and $_.DistinguishedName -like "" -and $_.AzureStale -eq $TRUE -and $_.enabled -eq $TRUE){ #Check onprem and azure if the account is stale and still enabled
            #Disable-ADObject -DistName $_.DistinguishedName -ObjectType "Users"
            #Move-ADObject -DistName $_.DistinguishedName -Location #Setup location
        }
    } else {
        break
    }
}

#Send notices based on age
$ComputerChangeCount = ($Computers | where {$_.enabled -eq $TRUE -and $_.age -gt 60}).count
$Computers | where enabled -EQ $TRUE | %{
    if($_.age -GT 30){
        if($_.age -gt 45){
            if($_.age -gt 60){
                if($ComputerChangeCount -le 50){
                    #Send-Emails -NoticeType "Disabled" -ComputerName $_.Name
                    #Disable-ADObject -DistName $_.distinguishedname -ObjectType "Computer"
                } else {
                    break
                }
            }
            #Send-Emails -NoticeType "Final" -ComputerName $_.Name
        }
        #Send-Emails -NoticeType "Initial" -ComputerName $_.Name
    }
}

#Move devices to accurate locations
$Computers | %{ 
    if($_.CorrectLocation -eq $FALSE){
        if($_.enabled -eq $TRUE){
            if($_.Destination -eq "No Match"){
                $UOItem = [PSCustomObject]@{
                    Hostname = $_.name;
                    "Last Logon Date" = $_.LastLogon;
                }
            } else{
                #Move-ADObject -DistName $_.DistinguishedName -Location $_.Destination
            }
        } elseif($_.enabled -eq $FALSE) {
            if($_.Destination -eq "No Match"){
                $UOItem = [PSCustomObject]@{
                    Hostname = $_.name;
                    "Last Logon Date" = $_.LastLogon;
                }
            } elseif($_.Decommission -eq $FALSE){
                #Move-ADObject -DistName $_.DistinguishedName -Location "OU=Disabled,OU=Domain Computers,DC=exac,DC=com"
            } else {
                #Backup-Computers -Name $_.DistinguishedName #Backup Computers
                #Move-ADObject -DistName $_.DistinguishedName -Location "OU=Decommission,OU=Disabled,OU=Domain Computers,DC=exac,DC=com" #Move to Decommission, ready for deletion
                $DecomItem = [PSCustomObject]@{
                    Hostname = $_.name;
                    "Last Logon Date" = $_.LastLogon;
                }
            }
        }
    }
    $DecommissionComputers.Add($DecomItem)
    $UnorganizedComputers.Add($UOItem)
}

#Create Service Desk Notices
if($StaleUsers){
    #Send-Emails -NoticeType "StaleUsers" -List $StaleUsers
}

if($DecommissionComputers) {
    #Send-Emails -NoticeType "Decommission" -List $DecommissionComputers
} 

if($UnorganizedComputers) {
    #Send-Emails -NoticeType "Unorganized" -List $UnorganizedComputers
}