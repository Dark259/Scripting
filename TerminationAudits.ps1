$Audit = "Q124"
$Final = "$($Audit)F"

$helperlist = "" #Helper File location with names

$termList = Import-Csv "$helperlist\$audit.csv"

$lists = @()

$termList | %{
    $list = $null
    $user = Get-ADUser $_.name -Properties enabled, lastlogondate, description, name, distinguishedname
    $description = $user.description
    $list+=@{
        Name = $user.name
        "Account Terminated" = ((($description).split(' '))[2])
        "Termination Processed" = $_.termprocessed
        "Termination Date" = $_.Term
    }
    $lists+=[pscustomobject]$list
}

$lists | Export-Csv "$helperlist\$Final.csv" -NoTypeInformation