import-Module configurationmanager

cd #Drive

$SCCMAllSystems = Get-CMDevice -Name * | select name, lastActiveTime, isclient, currentlogonuser

#Have to sub out export path
$SCCMNonBlankUsers | where{(($Omission -notcontains $_.currentlogonuser) -and ($Omission -notcontains $_.username)) -and (($_.username -notlike $null ) -or ($_.currentlogonuser -notlike $null))}

$Omission = @()

$SCCMEndUsers = $SCCMNonBlankUsers | %{
    if(($Omission -notcontains $_.currentlogonuser) -or ($Omission -notcontains $_.username)){ 
        if(($Omission -contains $_.currentlogonuser) -and ($_.username -notlike "")){ 
            $_.currentlogonuser = $_.username; 
            $_.username = "" ;
            return $_
        } 
        if(($Omission -contains $_.username) -and ($_.currentlogonuser -notlike "")){
            $_.username = "" ;
            return $_
        } 
        if((($Omission -notcontains $_.currentlogonuser) -and ($Omission -notcontains $_.username)) -and (($_.username -notlike $null) -and ($_.currentlogonuser -notlike $null))){
            $_.currentlogonuser = $_.username;
            return $_
        }
    }
}

$SCCMEndUsers | Export-Csv .\Downloads\sccmsystems.csv #If in a different window to perform SCCM functions, otherwise you will need to switch drives after this step
$SCCMSystems = Import-Csv .\Downloads\sccmsystems.csv

#API has 120/min. limit on it
$credPair = "";
$encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair));
$headers = @{ Authorization = "Basic $encodedCredentials"};

# Page specific swapping for requesters #$users = @(); $usersURL = "https://*organization*/api/v2/requesters?page=90";echo $usersURL;$users += Invoke-RestMethod -Uri $usersURL -Method Get -Headers $headers -UseBasicParsing -ContentType 'application/json'; 
$count = 1
$users = @() 
do {
    $FinalRequester = $users.requesters | where #Final requester in dry run (will always be the same even after deactivation)
    $usersURL = "https://*organization website*/api/v2/requesters?page=$count"
    echo $usersURL; $users += Invoke-RestMethod -Uri $usersURL -Method Get -Headers $headers -UseBasicParsing -ContentType 'application/json'
    $count++
    Start-Sleep -Seconds 15
} while (!$FinalRequester)

#Page Specific swapping for assets#$Assets = Invoke-RestMethod -Uri https://*organization*/api/v2/assets/2 -Method Get -Headers $headers -UseBasicParsing -ContentType 'application/json'
$count = 1
$Assets = @() 
do {
    $FinalAsset = $Assets.assets | where asset_tag -EQ "Asset-2"
    $url = "https://*organization*/api/v2/assets?order_type=desc&page=$count"
    echo $url
    $Assets += Invoke-RestMethod -Uri $url -Method Get -Headers $headers -UseBasicParsing -ContentType 'application/json'
    $count++
    Start-Sleep -Seconds 15
} while (!$FinalAsset)

#After pulling all assets and users into variables
$NonNullAssets = $Assets.assets|where user_id -eq $null
$common = compare-Object -ReferenceObject $SCCMSystems.name -DifferenceObject $NonNullAssets.name  -IncludeEqual | where sideindicator -Like "=="
$common = $common |select -ExpandProperty inputobject

$AssetToEmail = $SCCMSystems |%{
    if ($common -contains $_.name){
        return $_
    }
}

$AssetToEmail| %{
    $_.currentlogonuser = ($_.currentlogonuser).replace('EXACTECH\', "")
}

$UserIDs=@{}
$UserIDs = $AssetToEmail.currentlogonuser| 
%{
    $UserIDTemp = $users.requesters | where primary_email -Like "$_@exac.com" | select primary_email, id
    $AssetName = $AssetToEmail | 
    %{
        if($UserIDTemp.primary_email -contains "$($_.currentlogonuser)@exac.com") {return $_.name}
    } 
    return [pscustomobject]@{Asset = $AssetName;Username = $UserIDTemp.primary_email;UID = $UserIDTemp.id}
}

$UserIDs = $UserIDs | where asset -NotLike $null

$UserIDs=$UserIDs| %{
    $i=$null
    $i=$_
    if ($i.asset.count -gt 1){
        $i.asset| %{
            return [pscustomobject]@{Asset=$_;Username=$i.Username;UID=$i.UID}
        }
    }else {return $i}
}

#$Body = @{user_id=11001650927;description='Updated via API'} example, included in apiupdates foreach loop

$AssetDisplayID = $NonNullAssets | %{
    return [pscustomobject]@{AssetName = $_.name;DisplayID = $_.display_id}
}

$APIUpdates = $UserIDs | %{
    $AssetNameTemp = $_.asset
    $disIDTemp = $AssetDisplayID | %{
        if($AssetNameTemp -contains $_.assetname){
            return $_.displayid}
        }
    return [pscustomobject]@{asset=$_.asset;username=$_.username;UID=$_.UID;displayID=$disIDTemp}
}

$APIUpdates | %{$APIReturnURL = $null;$APIReturnURL = "https://*organization*/api/v2/assets/$($_.displayID)" ;$Body = $null; $Body = @{user_id=$_.UID} ;$APIReturnURL; $Body;Invoke-RestMethod $APIReturnURL -Headers $headers -Method PUT -ContentType 'application/json' -Body ($Body | ConvertTo-Json)}