#Group List
$ProtectedGroups = Get-ADGroup -LDAPFilter "(admincount=1)"

#Variables
$ADProperties=@("DisplayName","LastLogonDate","enabled")
$Audit = "Q124"

#Results

#Iterate recursively for every group in list passed
function Get-ADNestedGroupMembers {
    param (
        [string]$DistinguishedName
    )
    $RetVal = @()
    $Group=Get-ADGroup -Identity $DistinguishedName
    foreach ($obj in $(Get-ADGroupMember -Identity $DistinguishedName)) {
        switch ($obj.objectclass) {
            "group" {
                #Get-ADNestedGroupMembers -DistinguishedName $obj.DistinguishedName
            }
            "user" {
                $User=$null;
                $User=Get-ADUser -Identity $obj.DistinguishedName -Properties $ADProperties
                $RetVal+=[pscustomobject]@{Group=$Group.Name;User=$User.samaccountname;LastLogon=$User.lastlogondate;Enabled=$User.enabled}
            }
        }
    }
    return $RetVal
}

$Protected = @()
$ProtectedGroups | % {
    $Protected += Get-ADNestedGroupMembers -DistinguishedName (Get-ADGroup $_).distinguishedname
}

$Protected | Export-Csv "" -NoTypeInformation #Input location of exported file