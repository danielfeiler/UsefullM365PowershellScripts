#$AllMailADObjects=Get-adobject -filter "msExchRecipientTypeDetails -ne 0" -Properties DisplayName,Name,samaccountname,UserPrincipalName,Mail,ProxyAddresses,msExchRecipientTypeDetails,msExchRecipientDisplayType,msExchRemoteRecipientType
class DomainUsageInfo {

    [string]$DomainName
    [uint64]$GroupCount=0
    [uint64]$UserCount=0
    [uint64]$PublicFolderCount=0
    [bool] $IsPrimary
    DomainUsageInfo([string]$DomainName,[string]$ObjectClass,[bool]$IsPrimary){
        $this.DomainName = $DomainName
        $this.IsPrimary = $IsPrimary
        $this.Add($ObjectClass)
    }

    IncGroup(){
        $this.GroupCount++
    }
    IncUser(){
        $this.userCount++
    }
    IncPublicFolder(){
        $this.PublicFolderCount++
    }

    Add($ObjectClass){
        switch($ObjectClass){
        'user' {$this.IncUser()}
        'group' {$this.IncGroup()}
        'publicFolder' {$this.IncPublicFolder()}
        }
    }

}
class DomainInfoList {
    [System.Collections.Generic.List``1[DomainUsageInfo]]$Domains
    Add($DomainName, $ObjectClass,$IsPrimary){
        if($this.Domains.Count -eq 0) {
            $this.Domains.Add(([DomainUsageInfo]::New($DomainName,$ObjectClass,$IsPrimary)))
        }else{
            
        }
    }

    }
foreach($AdMailObject in $AllMailADObjects){
    foreach ($MailAddress in $AdMailObject.ProxyAddresses){
        if($MailAddress -imatch '^SMTP:'){
            $Domain
            $isPrimary = $MailAddress -cmatch '^SMTP:'
            $domainName = $MailAddress -ireplace '^smtp:.+?@'
            $objectClass = $AdMailObject
        }
    }
}