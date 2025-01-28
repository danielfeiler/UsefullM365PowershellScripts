<#
.SYNOPSIS
    This script will create a HTML file with an overview of all used mail domains in your Active Directory.
.DESCRIPTION
    This script will create a HTML file with an overview of all used mail domains in your Active Directory.
    The script will list all used mail domains and show how many groups, users and public folders are using the domain.
    The script will also show if the domain is the primary domain for the object.
    The script will create a HTML file in the location specified in the OutputFile parameter.
.PARAMETER OutputFile
    The path to the file where the HTML output should be saved.
    Default is C:\temp\DomainOverview.html
.EXAMPLE
    Get-AllUsedMaildomains.ps1 -OutputFile "C:\temp\DomainOverview.html"
    This will create a HTML file with an overview of all used mail domains in your Active Directory and save it to C:\temp\DomainOverview.html
.NOTES
    File Name      : Get-AllUsedMaildomains.ps1
    Author         : Daniel Feiler
    Prerequisite   : PowerShell V5.1 and the Active Directory module installed. 
    You can install the Active Directory module on Windows Server with the following command: Install-WindowsFeature RSAT-AD-PowerShell
    for Windows 10/11 you can install the RSAT tools with Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~
    For windows 10 older than 1809 you need to download the RSAT tools from the Microsoft website.
#>

<#PSScriptInfo
.VERSION 1.0
.GUID 5d96470d-c421-4e20-82fe-9b833e6bc385
.AUTHOR Daniel Feiler
.COMPANYNAME
.COPYRIGHT Daniel Feiler 2025
.TAGS Active Directory, Mail, Domain
.LICENSEURI https://www.gnu.org/licenses/gpl-3.0
.PROJECTURI https://github.com/danielfeiler/UsefullM365PowershellScripts
.ICONURI
.EXTERNALMODULEDEPENDENCIES ActiveDirectory
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
.PRIVATEDATA

#>

#Requires -Version 5.1
#Requires -Modules ActiveDirectory

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
