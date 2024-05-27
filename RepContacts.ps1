#Requires -Modules ExchangeOnlineManagement

#Connect-ExchangeOnline



Function Add-RepContact{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,
        [Parameter(Mandatory = $true)]
        [string]$FullName,
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )
    $cred = $Credential

    Connect-ExchangeOnline -Credential $cred
    
    $FN, $LN = $Fullname.split()

    try { 
         New-MailContact -Name "$FullName" -ExternalEmailAddress "$email" -FirstName "$FN" -LastName "$LN" -DisplayName "$FullName" -Alias $email.split('@')[0]
    }
    catch{
        Write-Host "Contact cannot be made because $_"
    }
    $firstLetter = $email.ToLower()[0]

    if($firstLetter -le "p") {
        if($firstLetter -le "h"){
            $group = "Rep Email Spam Bypass 1"
        }
        $group = "Rep Email Spam Bypass 2"
    }else { 
        $group = "Rep Email Spam Bypass 3"
    }

    #For when we want to complete automate since this will eliminate Any alphabetizing.
    <# $Sum1=(Get-TransportRule -Identity "Rep Email Spam Bypass 1").from
    ($Sum1| Measure-Object -Property length -Sum).sum
    $Sum2=(Get-TransportRule -Identity "Rep Email Spam Bypass 2").from
    ($Sum| Measure-Object -Property length -Sum).sum
    $Sum3=(Get-TransportRule -Identity "Rep Email Spam Bypass 3").from
    ($Sum3| Measure-Object -Property length -Sum).sum #>

    $Characters=(Get-TransportRule -Identity $group).from
    ($Characters| Measure-Object -Property length -Sum).sum

    if($Characters -gt 8000){
        try {
        $i=Get-TransportRule -Identity $group
        $i.From.Add($email) | Out-Null
        Set-TransportRule -Identity $i.Name -From $i.from
        Write-Host "$email added to $group"
        }
        catch{
            Write-Host "Rep was not added because $_"
        }
    } else {
        Write-Host "Cannot add due to character limit being $(($Characters| Measure-Object -Property length -Sum).sum)/8000"
    }

    Disconnect-ExchangeOnline -Confirm:$false

}

Function Remove-RepFilter{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    $cred=$Credential

    Connect-ExchangeOnline -Credential $cred
    $group = "None"
  <#   if($firstLetter -le "p") {
        if($firstLetter -le "h"){
            $group = "Rep Email Spam Bypass 1"
        }
        $group = "Rep Email Spam Bypass 2"
    }else { 
        $group = "Rep Email Spam Bypass 3"
    } 
  #>

    try {
        $rep1=Get-TransportRule -Identity "Rep Email Spam Bypass 1"
        $rep2=Get-TransportRule -Identity "Rep Email Spam Bypass 2"
        $rep3=Get-TransportRule -Identity "Rep Email Spam Bypass 3"
    }catch{
        Write-Host "Could not connect to Exchange"
    }

    if($rep1.from -contains $email){ 
        $group = "Rep Email Spam Bypass 1"
    } elseif($rep2.from -contains $email){
        $group = "Rep Email Spam Bypass 2"
    } elseif($rep3.from -contains $email){
        $group = "Rep Email Spam Bypass 3"
    } else {
        write-host "$Email has not been previously added to any of the existing groups."
    }

    try{
        $i=Get-TransportRule -Identity $group
        $i.From.Remove($email)   
        Set-TransportRule -Identity $i.Name -From $i.from
        Write-Host "Rep was removed from $group"
    } catch {
        write-host "Rep was not removed becuase $_"
    }

    Disconnect-ExchangeOnline -Confirm:$false

}