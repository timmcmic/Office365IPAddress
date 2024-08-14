
<#PSScriptInfo

.VERSION 1.0

.GUID e5d18bf9-f775-4a7a-adff-f3da4de7f72f

.AUTHOR timmcmic

.COMPANYNAME Microsoft

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 This script tests to see if an IP address is contained within the Office 365 URL and IP address ranges. 

#> 
Param(
    [Parameter(Mandatory = $false)]
    [string]$IPAddressToTest=""
)


