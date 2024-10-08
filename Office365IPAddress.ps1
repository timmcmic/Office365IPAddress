
<#PSScriptInfo

.VERSION 1.1.5

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
    [string]$IPAddressToTest="NONE",
    [Parameter(Mandatory = $false)]
    [string]$URLToTest="NONE",
    [Parameter(Mandatory = $true)]
    [string]$logFolderPath=$NULL,
    [Parameter(Mandatory = $false)]
    [boolean]$allowQueryIPLocationInformationFromThirdParty=$TRUE
)

Function new-LogFile
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$logFileName,
        [Parameter(Mandatory = $true)]
        [string]$logFolderPath
    )

    $ErrorActionPreference = 'Stop'

    [string]$logFileSuffix=".log"
    [string]$fileName=$logFileName+$logFileSuffix

    # Get our log file path

    $logFolderPath = $logFolderPath+"\"+$logFileName+"\"
    
    #Since $logFile is defined in the calling function - this sets the log file name for the entire script
    
    $global:LogFile = Join-path $logFolderPath $fileName

    #Test the path to see if this exists if not create.

    [boolean]$pathExists = Test-Path -Path $logFolderPath

    if ($pathExists -eq $false)
    {
        try 
        {
            #Path did not exist - Creating

            New-Item -Path $logFolderPath -Type Directory
        }
        catch 
        {
            throw $_
        } 
    }
}
Function Out-LogFile
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        $String,
        [Parameter(Mandatory = $false)]
        [boolean]$isError=$FALSE
    )

    # Get the current date

    [string]$date = Get-Date -Format G

    # Build output string
    #In this case since I abuse the function to write data to screen and record it in log file
    #If the input is not a string type do not time it just throw it to the log.

    if ($string.gettype().name -eq "String")
    {
        [string]$logstring = ( "[" + $date + "] - " + $string)
    }
    else 
    {
        $logString = $String
    }

    # Write everything to our log file and the screen

    $logstring | Out-File -FilePath $global:LogFile -Append

    #Write to the screen the information passed to the log.

    if ($string.gettype().name -eq "String")
    {
        Write-Host $logString
    }
    else 
    {
        write-host $logString | select-object -expandProperty *
    }

    #If the output to the log is terminating exception - throw the same string.

    if ($isError -eq $TRUE)
    {
        #Ok - so here's the deal.
        #By default error action is continue.  IN all my function calls I use STOP for the most part.
        #In this case if we hit this error code - one of two things happen.
        #If the call is from another function that is not in a do while - the error is logged and we continue with exiting.
        #If the call is from a function in a do while - write-error rethrows the exception.  The exception is caught by the caller where a retry occurs.
        #This is how we end up logging an error then looping back around.

        write-error $logString

        #Now if we're not in a do while we end up here -> go ahead and create the status file this was not a retryable operation and is a hard failure.

        exit
    }
}

function get-ClientGuid
{
    $functionClientGuid = $NULL

    out-logfile -string "Entering new-ClientGuid"

    try
    {   
        out-logfile -string "Obtain client GUID."
        $functionClientGuid = new-GUID -errorAction STOP
        out-logfile -string "Client GUID obtained successfully."
    }
    catch {
        out-logfile -string $_
        out-logfile -string "Unable to obtain client GUID." -isError:$true
    }

    out-logfile -string "Exiting new-ClientGuid"

    return $functionClientGuid
}

function get-Office365IPInformation
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $baseURL
    )

    $functionVersionInfo = $NULL

    out-logfile -string "Entering get-Office365IPInformation"

    try
    {   
        out-logfile -string 'Invoking web request for information...'
        $functionVersionInfo = Invoke-WebRequest -Uri $baseURL -errorAction:STOP
        out-logfile -string 'Invoking web request complete...'
    }
    catch {
        out-logfile -string $_
        out-logfile -string "Unable to invoke web request for Office 365 URL and IP information." -isError:$TRUE
    }

    out-logfile -string "Exiting get-Office365IPInformation"

    return $functionVersionInfo
}

function get-webURL
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$baseURL,
        [Parameter(Mandatory = $true)]
        [string]$clientGuid
    )

    $functionURL = $NULL

    out-logfile -string "Entering get-webURL"

    $functionURL = $baseURL+$clientGuid

    out-logfile -string "Exiting get-webURL"

    return $functionURL
}

function get-jsonData
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $data
    )

    $functionData = $NULL

    out-logfile -string "Entering get-jsonData"

    try {
        $functionData = convertFrom-Json $data -errorAction Stop
    }
    catch {
        out-logfile -string $_
        out-logfile -string "Unable to convert json data." -isError:$true
    }

    out-logfile -string "Exiting get-jsonData"

    return $functionData
}

function test-IPSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $IPAddress,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionNetwork = $NULL

    out-logfile -string "Entering test-IPSpace"

    foreach ($entry in $dataToTest)
    {
        Out-logfile -string ("Testing entry id: "+$entry.id)

        if ($entry.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.ips)
            {
                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = [System.Net.IpNetwork]::Parse($ipEntry)

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    out-logfile -string "The IP to test is contained within the entry.  Log the service."

                    $outputObject = new-Object psObject -property @{
                        M365Instance = $regionString
                        ID = $entry.ID
                        ServiceAreaDisplayName = $entry.ServiceAreaDisplayName
                        URLs = $entry.URLs
                        IPs = $entry.ips
                        IPInSubnet = $ipEntry
                        TCPPorts = $entry.tcpports
                        ExpressRoute = $entry.expressRoute
                        Required = $entry.required
                    }

                    out-logfile -string $outputObject

                    $global:outputArray += $outputObject
                 }
                 else
                 {
                    out-logfile -string "The IP to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "IP count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-IPSpace"
}

function test-URLSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $IPAddress,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionNetwork = $NULL

    out-logfile -string "Entering test-IPSpace"

    foreach ($entry in $dataToTest)
    {
        Out-logfile -string ("Testing entry id: "+$entry.id)

        if ($entry.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.ips)
            {
                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = [System.Net.IpNetwork]::Parse($ipEntry)

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    out-logfile -string "The IP to test is contained within the entry.  Log the service."

                    $outputObject = new-Object psObject -property @{
                        M365Instance = $regionString
                        ID = $entry.ID
                        ServiceAreaDisplayName = $entry.ServiceAreaDisplayName
                        URLs = $entry.URLs
                        IPs = $entry.ips
                        IPInSubnet = $ipEntry
                        TCPPorts = $entry.tcpports
                        ExpressRoute = $entry.expressRoute
                        Required = $entry.required
                    }

                    out-logfile -string $outputObject

                    $global:outputArray += $outputObject
                 }
                 else
                 {
                    out-logfile -string "The IP to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "IP count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-IPSpace"
}

function test-IPChangeSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $changeDataToTest,
        [Parameter(Mandatory = $true)]
        $IPAddress,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionNetwork = $NULL
    $functionOriginalID = $null
    $functionServiceAreaDisplayName = $NULL

    out-logfile -string "Entering test-IPSpace"

    foreach ($entry in $changeDataToTest)
    {
        Out-logfile -string ("Testing change entry id: "+$entry.id)

        if ($entry.add.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.add.ips)
            {
                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = [System.Net.IpNetwork]::Parse($ipEntry)

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    out-logfile -string "The IP to test is contained within the entry.  Log the service."

                    $functionOriginalID = $datatoTest | where {$_.id -eq $entry.EndpointSetID}

                    if ($functionOriginalID.ServiceAreaDisplayName -eq $NULL)
                    {
                        $functionServiceAreaDisplayName = "Endpoint Set ID No Longer Active"
                    }
                    else
                    {
                        $functionServiceAreaDisplayName = $functionOriginalID.ServiceAreaDisplayName
                    }

                    $outputObject = new-Object psObject -property @{
                        M365Instance = $regionString
                        ChangeID = $entry.ID
                        Disposition = $entry.Disposition
                        EndpointSetID = $entry.endpointSetId
                        Version = $entry.Version
                        ServiceAreaDisplayName = $functionServiceAreaDisplayName
                        IPsAdded = $entry.add.ips
                        IPInSubnet = $ipEntry
                    }

                    out-logfile -string $outputObject

                    $global:outputChangeArray += $outputObject
                 }
                 else
                 {
                    out-logfile -string "The IP to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "IP count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-IPSpace"
}

function test-IPRemoveSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $changeDataToTest,
        [Parameter(Mandatory = $true)]
        $IPAddress,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionNetwork = $NULL
    $functionOriginalID = $null
    $functionServiceAreaDisplayName = $NULL

    out-logfile -string "Entering test-IPSpace"

    foreach ($entry in $changeDataToTest)
    {
        Out-logfile -string ("Testing change entry id: "+$entry.id)

        if ($entry.remove.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.remove.ips)
            {
                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = [System.Net.IpNetwork]::Parse($ipEntry)

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    out-logfile -string "The IP to test is contained within the entry.  Log the service."

                    $functionOriginalID = $datatoTest | where {$_.id -eq $entry.EndpointSetID}

                    if ($functionOriginalID.ServiceAreaDisplayName -eq $null)
                    {
                        $functionServiceAreaDisplayName = "Endpoint Set ID No Longer Active"
                    }
                    else
                    {
                        $functionServiceAreaDisplayName = $functionOriginalID.ServiceAreaDisplayName
                    }

                    $outputObject = new-Object psObject -property @{
                        M365Instance = $regionString
                        ChangeID = $entry.ID
                        Disposition = $entry.Disposition
                        EndpointSetID = $entry.endpointSetId
                        Version = $entry.Version
                        ServiceAreaDisplayName = $functionServiceAreaDisplayName
                        IPsRemove = $entry.remove.ips
                        IPInSubnet = $ipEntry
                    }

                    out-logfile -string $outputObject

                    $global:outputRemoveArray += $outputObject
                 }
                 else
                 {
                    out-logfile -string "The IP to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "IP count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-IPSpace"
}

function get-IPLocationInformation
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $ipAddress
    )

    $functionData = ""
    $ipLocationProvider = "https://api.country.is/"
    $ipLocationQuery = $ipLocationProvider+$ipAddress

    out-logfile -string "Entering get-IPLocationInformation"

    try {
        $functionData = invoke-WebRequest $ipLocationQuery
    }
    catch {
        out-logfile -string $_
        out-logfile -string "Unable to invoke web request for geolocation lookup."
        $functionData = "Failed"
    }

    out-logfile -string "Exiting get-IPLocationInformation"

    return $functionData
}

#Following function credit to author -> https://github.com/fleschutz/PowerShell/blob/main/docs/check-ipv4-address.md
function IsIPv4AddressValid { param([string]$IP)
	$RegEx = "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
	if ($IP -match $RegEx) {
		return $true
	} else {
		return $false
	}
}

#Follwoing function credit to author -> https://github.com/fleschutz/PowerShell/blob/main/docs/check-ipv6-address.md

function IsIPv6AddressValid { param([string]$IP)
    $IPv4Regex = '(((25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))'
    $G = '[a-f\d]{1,4}'
    $Tail = @(":",
    "(:($G)?|$IPv4Regex)",
    ":($IPv4Regex|$G(:$G)?|)",
    "(:$IPv4Regex|:$G(:$IPv4Regex|(:$G){0,2})|:)",
    "((:$G){0,2}(:$IPv4Regex|(:$G){1,2})|:)",
    "((:$G){0,3}(:$IPv4Regex|(:$G){1,2})|:)",
    "((:$G){0,4}(:$IPv4Regex|(:$G){1,2})|:)")
    [string] $IPv6RegexString = $G
    $Tail | foreach { $IPv6RegexString = "${G}:($IPv6RegexString|$_)" }
    $IPv6RegexString = ":(:$G){0,5}((:$G){1,2}|:$IPv4Regex)|$IPv6RegexString"
    $IPv6RegexString = $IPv6RegexString -replace '\(' , '(?:' # make all groups non-capturing
    [regex] $IPv6Regex = $IPv6RegexString
    if ($IP -imatch "^$IPv6Regex$") {
    	return $true
    } else {
    	return $false
    }
}
function test-Parameters
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $ipAddressToTest,
        [Parameter(Mandatory = $true)]
        $URLToTest
    )

    $functionIPV4 ="."
    $functionIPV6 = ":"

    if (($URLToTest -eq "NONE") -and ($IPAddressToTest -eq "NONE"))
    {
        out-logfile -string "No URL or IP entries were specified.  To perform an analysis specify either an IP address or URL."
        out-logfile -string "ERROR: NO IP OR URL SPECIFIED" -isERROR:$TRUE
    }
    elseif (($URLToTest -ne "NONE") -and ($IPAddressToTest -ne "NONE"))
    {
        out-logfile -string "Both a URL and IP address were specified to test.  Specify only a IP Address or URL to proceed - not both."
        out-logfile -string "ERROR: URL AND IP ADDRESS SPECIFIED" -isERROR:$TRUE
    }
    elseif ($URLToTest -ne "NONE")
    {
        out-logfile -string "URL specified to test - proceed."
    }
    elseif ($IPAddressToTest -ne "NONE")
    {
        out-logfile -string "IP specified to test - proceed."

        if ($ipAddressToTest.contains($functionIPV4))
        {
            if (IsIPv4AddressValid -ip $ipAddressToTest)
            {
                out-logfile -string "IPv4 address in proper format."
            }
            else 
            {
                out-logfile -string "IPv4 address is not in proper format."
            }
        }
        elseif ($ipAddressToTest.contains($functionIPV6))
        {
            if (IsIPv6AddressValid -ip $ipAddressToTest)
            {
                out-logfile -string "IPv6 address in proper format."
            }
            else 
            {
                out-logfile -string "IPv6 address is not in proper format."
            }
        }
    }
}

Function Test-PowershellVersion
{
    [cmdletbinding()]

    $functionPowerShellVersion = $NULL

    out-logfile -string "Entering Test-PowerShellVersion"

    #Write function parameter information and variables to a log file.

    $functionPowerShellVersion = $PSVersionTable.PSVersion

    out-logfile -string "Determining powershell version."
    out-logfile -string ("Major: "+$functionPowerShellVersion.major)
    out-logfile -string ("Minor: "+$functionPowerShellVersion.minor)
    out-logfile -string ("Patch: "+$functionPowerShellVersion.patch)
    out-logfile -string $functionPowerShellVersion

    if ($functionPowerShellVersion.Major -lt 7)
    {
        out-logfile -string "Powershell 7 and higher is required to run this script."
        out-logfile -string "Please run module from Powershell 7.x"
        out-logfile -string "" -isError:$true
    }
    else
    {
        out-logfile -string "Powershell version is not powershell 5.X proceed."
    }

    out-logfile -string "Exiting Test-PowerShellVersion"

}

function get-logFileName
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $ipAddressToTest,
        [Parameter(Mandatory = $true)]
        $URLToTest
    )

    [string]$logFileName = ""
    $functionURL = ""
    $functionURL1 = "//"
    $functionURL2 = "/"

    write-host ("IPAddressToTest: "+$ipAddressToTest)
    write-host ("URLToTest: "+$URLToTest)
    
    if (($ipAddressToTest -eq "NONE") -and ($urlToTest -eq "NONE"))
    {
        #Neither an IP address or URL was specified - used generic name.
        write-host "Log File Name Generic - neither URL or IP specified."
        $logFileName = "Office365IPAddress"
        write-host $logFileName
    }
    elseif ((($ipAddressToTest -ne "NONE") -and ($urlToTest -ne "NONE"))) 
    {
        #Both an IP and URL were specified - use a generic log name.
        write-host "Log File Name Generic - both URL or IP specified."
        $logFileName = "Office365IPAddress"
        write-host $logFileName
    }
    elseif ($ipAddressToTest -ne "")
    {
        write-host "Only IP address specified - use IP as log file name."
        if ($IPAddressToTest.contains("."))
        {
            $logFileName = $IPAddressToTest.replace(".","-")
            write-host $logFileName
        }
        else 
        {
            $logFileName = $IPAddressToTest.replace(":","-")
            write-host $logFileName
        }
    }
    elseif ($urlToTest -ne "")
    {
        write-host "Only URL was specified - use URL as name."
        #Test the string to see if it was specified in the format of a URL.

        if ($urlToTest.contains($functionURL1))
        {
            $functionURL = $urlToTest.split($functionURL1)

            foreach ($url in $functionURL)
            {
                write-host $url
            }

            if ($functionURL.contains("$functionURL2"))
            {
                $functionURL = $urlToTest.split($functionURL2)

                foreach ($url in $functionURL)
                {
                    write-host $url
                }
            }

            $functionURL = $functionURL[1]
            write-host $functionURL
        }
        else 
        {
            $functionURL = $URLToTest
            write-host $functionURL
        }

        $logfilename = $functionurl.replace(".","-")
        write-host $logFileName
    }

    return $logFileName
}

#=====================================================================================
#Begin main function body.
#=====================================================================================

#Define function variables.

$clientGuid = $NULL
$allVersionInfoBaseURL = "https://endpoints.office.com/version?ClientRequestId="
$allVersionInfoURL = $NULL
$allVersionInfo = $NULL

$allIPInformationWorldWideBaseURL = "https://endpoints.office.com/endpoints/Worldwide?ClientRequestId="
$allIPInformationChinaBaseURL = "https://endpoints.office.com/endpoints/China?ClientRequestId="
$allIPInformationUSGovGCCHighBaseURL = "https://endpoints.office.com/endpoints/USGovGCCHigh?ClientRequestId="
$allIPInformationUSGovDODBaseURL = "https://endpoints.office.com/endpoints/USGovDOD?ClientRequestId="

$allIPChangeInformationWorldWideBaseURL = "https://endpoints.office.com/changes/worldwide/0000000000?clientrequestid="
$allIPChangeInformationChinaBaseURL = "https://endpoints.office.com/changes/china/0000000000?clientrequestid="
$allIPChangeInformationUSGovGCCHighBaseURL = "https://endpoints.office.com/changes/usgovgcchigh/0000000000?clientrequestid="
$allIPChangeInformationUSGovDODBaseURL = "https://endpoints.office.com/changes/usgovdod/0000000000?clientrequestid="


$allIPInformationWorldWideURL = $NULL
$allIPInformationChinaURL = $NULL
$allIPInformationUSGovGCCHighURL = $NULL
$allIPInformationUSGovDODURL = $NULL

$allIPChangeInformationWorldWideURL = $NULL
$allIPChangeInformationChinaURL = $NULL
$allIPChangeInformationUSGovGCCHighURL = $NULL
$allIPChangeInformationUSGovDODURL = $NULL

$allIPInformationWorldWide = $NULL
$allIPInformationChina = $NULL
$allIPInfomrationUSGovGCCHigh = $NULL
$allIPInformationUSGovDOD = $NULL

$allIPChangeInformationWorldWide = $NULL
$allIPChangeInformationChina = $NULL
$allIPChangeInfomrationUSGovGCCHigh = $NULL
$allIPChangeInformationUSGovDOD = $NULL

$worldWideRegionString = "Microsoft 365 Worldwide (+GCC)"
$chinaRegionString = "Microsoft 365 operated by 21 Vianet"
$gccHighRegionString = "Microsoft 365 U.S. Government GCC High"
$dodRegionString = "Microsoft 365 U.S. Government DoD"

$ipLocation = ""

$global:outputArray = @()
$global:outputChangeArray=@()
$global:outputRemoveArray=@()

#Create the log file.

try {
    $logfileName = get-logFileName -ipAddressToTest $ipAddressToTest -urlToTest $URLToTest
}
catch {
    write-error "Error calculating log file name."
}

new-logfile -logFileName $logFileName -logFolderPath $logFolderPath

$outputXMLFile = $global:LogFile.replace(".log",".xml")
$outputChangeXMLFile = $global:LogFile.replace(".log","Adds.xml")
$outputRemoveXMLFile = $global:LogFile.replace(".log","Removes.xml")

out-logfile -string $global:LogFile
out-logfile -string $outputXMLFile
out-logfile -string $outputChangeXMLFile

#Start logging

out-logfile -string "*********************************************************************************"
out-logfile -string "Start Office365IPAddress"
out-logfile -string "*********************************************************************************"

out-logfile -string "Invoking powershell version test..."

Test-PowerShellVersion

out-logfile -string "Invoking parameter test..."

Test-Parameters -ipAddressToTest $ipAddressToTest -urlToTest $urlToTest

out-logfile -string "Obtaining client guid for web requests."

$clientGuid = get-ClientGuid
$clientGUid = $clientGuid.tostring()    

out-logfile -string $clientGuid

out-logfile -string "Obtain and log version information for the query."

$allVersionInfoURL = get-webURL -baseURL $allVersionInfoBaseURL -clientGuid $clientGuid

out-logfile -string $allVersionInfoURL

$allVersionInfo = get-Office365IPInformation -baseURL $allVersionInfoURL

$allVersionInfo = get-jsonData -data $allVersionInfo

foreach ($version in $allVersionInfo)
{
    out-logfile -string ("Instance: "+$version.instance+" VersionInfo: "+$version.latest)
}

out-logfile -string "Calculate URLS for Office 365 IP Addresses"

$allIPInformationWorldWideURL = get-webURL -baseURL $allIPInformationWorldWideBaseURL -clientGuid $clientGuid
out-logfile -string $allIPInformationWorldWideURL
$allIPInformationChinaURL = get-webURL -baseURL $allIPInformationChinaBaseURL -clientGuid $clientGuid
out-logfile -string $allIPInformationChinaURL
$allIPInformationUSGovGCCHighURL = get-webURL -baseURL $allIPInformationUSGovGCCHighBaseURL -clientGuid $clientGuid
out-logfile -string $allIPInformationUSGovGCCHighURL
$allIPInformationUSGovDODURL = get-webURL -baseURL $allIPInformationUSGovDODBaseURL -clientGuid $clientGuid
out-logfile -string $allIPInformationUSGovDODURL

out-logfile -string "Calculate URLS for Office 365 IP Addresses Change"

$allIPChangeInformationWorldWideURL = get-webURL -baseURL $allIPChangeInformationWorldWideBaseURL -clientGuid $clientGuid
out-logfile -string $allIPChangeInformationWorldWideURL
$allIPChangeInformationChinaURL = get-webURL -baseURL $allIPChangeInformationChinaBaseURL -clientGuid $clientGuid
out-logfile -string $allIPChangeInformationChinaURL
$allIPChangeInformationUSGovGCCHighURL = get-webURL -baseURL $allIPChangeInformationUSGovGCCHighBaseURL -clientGuid $clientGuid
out-logfile -string $allIPChangeInformationUSGovGCCHighURL
$allIPChangeInformationUSGovDODURL = get-webURL -baseURL $allIPChangeInformationUSGovDODBaseURL -clientGuid $clientGuid
out-logfile -string $allIPChangeInformationUSGovDODURL

out-logfile -string "Obtain IP information for all available Office 365 instances."

$allIPInformationWorldWide = get-Office365IPInformation -baseURL $allIPInformationWorldWideURL
$allIPInformationChina = get-Office365IPInformation -baseURL $allIPInformationChinaURL
$allIPInfomrationUSGovGCCHigh = get-Office365IPInformation -baseURL $allIPInformationUSGovGCCHighURL
$allIPInformationUSGovDOD = get-Office365IPInformation -baseURL $allIPInformationUSGovDODURL

out-logfile -string "Obtain IP information for all available Office 365 instances changes."

$allIPChangeInformationWorldWide = get-Office365IPInformation -baseURL $allIPChangeInformationWorldWideURL
$allIPChangeInformationChina = get-Office365IPInformation -baseURL $allIPChangeInformationChinaURL
$allIPChangeInfomrationUSGovGCCHigh = get-Office365IPInformation -baseURL $allIPChangeInformationUSGovGCCHighURL
$allIPChangeInformationUSGovDOD = get-Office365IPInformation -baseURL $allIPChangeInformationUSGovDODURL

out-logfile -string "Convert IP information from JSON."

$allIPInformationWorldWide = get-jsonData -data $allIPInformationWorldWide
$allIPInformationChina = get-jsonData -data $allIPInformationChina
$allIPInfomrationUSGovGCCHigh = get-jsonData -data $allIPInfomrationUSGovGCCHigh
$allIPInformationUSGovDOD = get-jsonData -data $allIPInformationUSGovDOD

out-logfile -string "Convert IP information from JSON changes."

$allIPChangeInformationWorldWide = get-jsonData -data $allIPChangeInformationWorldWide
$allIPChangeInformationChina = get-jsonData -data $allIPChangeInformationChina
$allIPChangeInfomrationUSGovGCCHigh = get-jsonData -data $allIPChangeInfomrationUSGovGCCHigh
$allIPChangeInformationUSGovDOD = get-jsonData -data $allIPChangeInformationUSGovDOD



out-logfile -string "Begin testing IP spaces for presence of the specified IP address."

test-IPSpace -dataToTest $allIPInformationWorldWide -IPAddress $IPAddressToTest -regionString $worldWideRegionString
test-IPSpace -dataToTest $allIPInformationChina -IPAddress $IPAddressToTest -regionString $chinaRegionString
test-IPSpace -dataToTest $allIPInfomrationUSGovGCCHigh -IPAddress $IPAddressToTest -regionString $gccHighRegionString
test-IPSpace -dataToTest $allIPInformationUSGovDOD -IPAddress $IPAddressToTest -regionString $dodRegionString

if ($global:outputArray.count -gt 0)
{
    out-logfile -string "IPs found in Active Service - test for changes."

    test-IPChangeSpace -dataToTest $allIPInformationWorldWide -IPAddress $IPAddressToTest -regionString $worldWideRegionString -changeDataToTest $allIPChangeInformationWorldWide
    test-IPChangeSpace -dataToTest $allIPInformationChina -IPAddress $IPAddressToTest -regionString $chinaRegionString -changeDataToTest $allIPChangeInformationChina
    test-IPChangeSpace -dataToTest $allIPInfomrationUSGovGCCHigh -IPAddress $IPAddressToTest -regionString $gccHighRegionString -changeDataToTest $allIPChangeInfomrationUSGovGCCHigh
    test-IPChangeSpace -dataToTest $allIPInformationUSGovDOD -IPAddress $IPAddressToTest -regionString $dodRegionString -changeDataToTest $allIPChangeInformationUSGovDOD
}
else 
{
    out-logfile -string "No IP addresses found -> no need to test for additions."
}

if ($global:outputArray.count -eq 0)
{
    out-logfile -string "Since the global output count is equal to 0 -> testing to see if the IP address was removed."

    test-IPRemoveSpace -dataToTest $allIPInformationWorldWide -IPAddress $IPAddressToTest -regionString $worldWideRegionString -changeDataToTest $allIPChangeInformationWorldWide
    test-IPRemoveSpace -dataToTest $allIPInformationChina -IPAddress $IPAddressToTest -regionString $chinaRegionString -changeDataToTest $allIPChangeInformationChina
    test-IPRemoveSpace -dataToTest $allIPInfomrationUSGovGCCHigh -IPAddress $IPAddressToTest -regionString $gccHighRegionString -changeDataToTest $allIPChangeInfomrationUSGovGCCHigh
    test-IPRemoveSpace -dataToTest $allIPInformationUSGovDOD -IPAddress $IPAddressToTest -regionString $dodRegionString -changeDataToTest $allIPChangeInformationUSGovDOD
}
else 
{
    out-logfile -string "The IP address was found - no need to test for removals."
}

if ($global:outputArray.count -gt 0)
{
    if ($allowQueryIPLocationInformationFromThirdParty -eq $TRUE)
    {
        $ipLocation = get-IPLocationInformation -ipAddress $ipAddressToTest

        out-logfile -string $ipLocation

        if ($ipLocation -ne "Failed")
        {
            out-logfile -string "Converting IP location JSON."
            $ipLocation = get-jsonData -data $ipLocation
        }
    }

    out-logfile -string "*"
    out-logfile -string "**"
    out-logfile -string "******************************************************"
    out-logfile -string ("The IP Address: "+$IPAddressToTest+ " was located in any Office 365 Services.")

    if ($ipLocation -ne "Failed")
    {   
        out-logfile -string ("The IP Address geo-location is: "+$ipLocation.country)
    }

    write-host "IP entries present in the following Office 365 Services:"

    foreach ($entry in $global:outputArray)
    {
        $entry
    }

    if ($global:outputChangeArray)
    {
        write-host "IP entries present in the changes file:"

        foreach ($entry in $global:outputChangeArray)
        {
            $entry
        }
    }

    out-logfile -string "A XML file containing the above entries is available in the log directory."
    out-logfile -string "******************************************************"
    out-logfile -string "**"
    out-logfile -string "*"

    if ($global:outputArray.count -gt 0)
    {
        out-logfile -string ""
        out-logfile -string "IPs was located in the following service description:"
        out-logfile -string ""

        foreach ($entry in $global:outputArray)
        {
            out-logfile -string $entry
        }
    
        $global:outputArray | Export-Clixml -Path $outputXMLFile
    }

    if ($global:outputChangeArray.count -gt 0)
    {
        out-logfile -string ""
        out-logfile -string "IP was located in the following version additions since 2018:"
        out-logfile -string ""

        foreach ($entry in $global:outputChangeArray)
        {
            out-logfile -string $entry
        }

        $global:outputChangeArray | Export-Clixml -Path $outputChangeXMLFile
    }  
}
else 
{
    out-logfile -string "******************************************************"
    out-logfile -string ("The IP Address: "+$IPAddressToTest+ " was NOT located in the following Office 365 Services.")
    out-logfile -string "******************************************************"

    if ($global:outputRemoveArray.count -gt 0)
    {
        write-host "The following changes removed the IP address:"

        foreach ($entry in $global:outputRemoveArray)
        {
            $entry
        }

        foreach ($entry in $global:outputRemoveArray)
        {
            out-logfile -string $entry
        }

        $global:outputRemoveArray | Export-Clixml -Path $outputRemoveXMLFile
    }

}

