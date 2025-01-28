[CmdletBinding()]
param(
    [string]$OutputFile = "C:\temp\DomainOverview.html"
)
$AllMailADObjects = Get-adobject -ldapfilter "msExchRecipientTypeDetails=*" -Properties DisplayName, Name, samaccountname, UserPrincipalName, Mail, ProxyAddresses, msExchRecipientTypeDetails, msExchRecipientDisplayType, msExchRemoteRecipientType
class DomainUsageInfo {

    [string]$DomainName
    [uint64]$GroupCount = 0
    [uint64]$UserCount = 0
    [uint64]$PublicFolderCount = 0
    [bool] $IsPrimary
    DomainUsageInfo([string]$DomainName, [string]$ObjectClass, [bool]$IsPrimary) {
        $this.DomainName = $DomainName
        $this.IsPrimary = $IsPrimary
        $this.Add($ObjectClass)
    }
    DomainUsageInfo() {
        $this.DomainName = $null
        $this.IsPrimary = $null

    }
    IncGroup() {
        $this.GroupCount++
    }
    IncUser() {
        $this.userCount++
    }
    IncPublicFolder() {
        $this.PublicFolderCount++
    }

    Add($ObjectClass) {
        switch ($ObjectClass) {
            'user' { $this.IncUser() }
            'group' { $this.IncGroup() }
            'publicFolder' { $this.IncPublicFolder() }
        }
    }

}
class DomainInfo {
    [string]$DomainName
    [System.Collections.Generic.List``1[DomainUsageInfo]]$DomainUsageInfo 
    DomainInfo([string]$DomainName, [string]$ObjectClass, [bool]$IsPrimary) {
        $this.DomainName = $DomainName -ireplace '^smtp:.+?@'
        $this.DomainUsageInfo = New-Object System.Collections.Generic.List``1[DomainUsageInfo]
        $this.DomainUsageInfo.Add(([DomainUsageInfo]::New($DomainName, $ObjectClass, $IsPrimary)))
    }
    DomainInfo([string]$DomainName) {
        $this.DomainName = $DomainName -ireplace '^smtp:.+?@'
        $this.DomainUsageInfo = New-Object System.Collections.Generic.List``1[DomainUsageInfo]
    }
    Add([string]$ObjectClass, [bool]$IsPrimary) {
        $DomainIdx = $this.DomainUsageInfo.FindIndex({ $args.DomainName -eq $this.DomainName -and $args.IsPrimary -eq $IsPrimary })
        if ($DomainIdx -ne -1) {
            $this.DomainUsageInfo[$DomainIdx].Add($ObjectClass)
        }
        else {
            $this.DomainUsageInfo.Add(([DomainUsageInfo]::New($this.DomainName, $ObjectClass, $IsPrimary)))
        }
    }
}
class DomainInfoList {
    [System.Collections.Generic.List``1[DomainInfo]]$Domains
    DomainInfoList() {
        $this.Domains = New-Object System.Collections.Generic.List``1[DomainInfo]
    }
    Add([string]$DomainName, [string]$ObjectClass, [bool]$IsPrimary) {
        $DomainName = $DomainName -ireplace '^smtp:.+?@'
        if ($this.Domains.Count -eq 0) {
            $this.Domains.Add(([DomainInfo]::New($DomainName, $ObjectClass, $IsPrimary)))
        }
        else {
            $DomainIdx = $this.Domains.FindIndex({ $args.DomainName -eq $DomainName }) 
            if ($DomainIdx -ne -1) {
                
                $this.Domains[$DomainIdx].Add($ObjectClass, $IsPrimary) 
            }
            else {
                $this.Domains.Add(([DomainInfo]::New($DomainName, $ObjectClass, $IsPrimary)))
            }
        }

    }
    [string] GetDomainSummary() {
        $filler = $null
        $domainCount = 0
        $returnValue = @'
    <!DOCTYPE html>
        <html lang="en">
            <head>
                <title>Domain Overview</title>
                <meta name="viewport" content="width=device-width, initial-scale=1">
                <meta name="description" content="List of Used Domains">
                <style>
                body {background-color:#ffffff;background-repeat:no-repeat;background-position:top left;background-attachment:fixed;}
                h1{font-family:Arial, sans-serif;color:#000000;background-color:#ffffff;}
                p {font-family:Georgia, serif;font-size:14px;font-style:normal;font-weight:normal;color:#000000;background-color:#ffffff;}
                .doaminhead {font-family:Arial, sans-serif;color:#000000;background-color:#ffffff;Text-align:center;}
                .text-left {text-align:left;}
                .text-right {text-align:right;}
                .text-center {text-align:center;}
                table,td,tr,th {font-family:Arial, sans-serif;color:#333333;border-width: 1px;border-style: solid;border-color: #666666;}
                th {background-color:#dedede;}
                td,th {min-width: 100px;padding: 5px;}
                th.number,td.number {text-align:right;}
                </style>
            </head>
            <body>
                <h1>List of all used domains</h1>
                <table>
        <thead>
        <tr>
            <th colspan="2" class="text-center">Domain overview</th>
            
        </tr>
        </thead>
        <tbody>
'@

        foreach ($Domain in $this.Domains) {
            $domainCount++
            $returnValue += @"
            <tr>
                <th Colspan="2" class="doaminhead">$($Domain.DomainName)</th>
            </tr>
            <tr>
"@
            foreach ($DomainUsageInfo in $Domain.DomainUsageInfo | Sort-Object -Property IsPrimary -Descending ) {
                
                if ($Domain.DomainUsageInfo.count -eq 1) {
                    $filler = @"
<td>
                
                <table>
                <tbody>
                <tr>
                <th colspan="2"  class="text-left">$(if(!$DomainUsageInfo.IsPrimary){"Domain is primary for:</th>"}else{"Domain is additional for:"})</th>
                </tr>
                <th class="text-left">Groups</th>
                <td class="number">0</td>
                </tr>
                <tr>
                <th class="text-left">Users</th>
                <td class="number">0</td>
                </tr>
                <tr>
                <th class="text-left">Public Folders</th>
                <td class="number">0</td>
                </tr>
                </tbody>
                </table>
                </td>
"@
                }
                else {
                    $filler = $null
                }

            
                $returnValue += @"
                <!--Primary Domain $($DomainCount)-->
                <td>
                
                <table>
                <tbody>
                <tr>
                <th colspan="2"  class="text-left">$(if($DomainUsageInfo.IsPrimary){"Domain is primary for:</th>"}else{"Domain is additional for:"})</th>
                </tr>
                <th class="text-left">Groups</th>
                <td class="number">$($DomainUsageInfo.GroupCount)</td>
                </tr>
                <tr>
                <th class="text-left">Users</th>
                <td class="number">$($DomainUsageInfo.UserCount)</td>
                </tr>
                <tr>
                <th class="text-left">Public Folders</th>
                <td class="number">$($DomainUsageInfo.PublicFolderCount)</td>
                </tr>
                </tbody>
                </table>
                </td>
                <!--Secondary Domain $($DomainCount)-->             
"@
                if ((!$DomainUsageInfo.IsPrimary) -and $Domain.DomainUsageInfo.count -eq 1) {
                    $returnValue = $returnValue -replace $('<!--Primary Domain '+ $DomainCount+'-->'), $filler
                }
                elseif ($DomainUsageInfo.IsPrimary -and $Domain.DomainUsageInfo.count -eq 1) {
                    $returnValue = $returnValue -replace $('<!--Secondary Domain '+$domainCount+'-->'), $filler
                }

            }
        }
        $returnValue += @'
    </tbody>
    </table>
    </body>
    </html>
'@
        return $returnValue 
    }
}
$DomainInfoList = New-Object  DomainInfoList
foreach ($AdMailObject in $AllMailADObjects) {
    foreach ($MailAddress in $AdMailObject.ProxyAddresses) {
        if ($MailAddress -imatch '^SMTP:') {
            $DomainInfoList.Add($MailAddress, $($AdMailObject.ObjectClass), $($MailAddress -cmatch '^SMTP:'))
        }
    }
}
$DomainInfoList.GetDomainSummary() | Out-File -FilePath $OutputFile -Encoding utf8
