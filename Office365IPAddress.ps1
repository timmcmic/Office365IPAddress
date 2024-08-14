
<#PSScriptInfo

.VERSION 1.1

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
    [string]$IPAddressToTest="",
    [Parameter(Mandatory = $true)]
    [string]$logFolderPath=$NULL,
    [Parameter(Mandatory = $true)]
    [boolean]$allowQueryIPLocationInformationFromThirdParty
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
        $IPAddress
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

#=====================================================================================
#Begin main function body.
#=====================================================================================

#Define function variables.

if ($IPAddressToTest.contains("."))
{
    $logFileName = $IPAddressToTest.replace(".","-")
}
else 
{
    $logFileName = $IPAddressToTest.replace(":","-")
}


$clientGuid = $NULL
$allVersionInfoBaseURL = "https://endpoints.office.com/version?ClientRequestId="
$allVersionInfoURL = $NULL
$allVersionInfo = $NULL

$allIPInformationWorldWideBaseURL = "https://endpoints.office.com/endpoints/Worldwide?ClientRequestId="
$allIPInformationChinaBaseURL = "https://endpoints.office.com/endpoints/China?ClientRequestId="
$allIPInformationUSGovGCCHighBaseURL = "https://endpoints.office.com/endpoints/USGovGCCHigh?ClientRequestId="
$allIPInformationUSGovDODBaseURL = "https://endpoints.office.com/endpoints/USGovDOD?ClientRequestId="

$allIPInformationWorldWideURL = $NULL
$allIPInformationChinaURL = $NULL
$allIPInformationUSGovGCCHighURL = $NULL
$allIPInformationUSGovDODURL = $NULL

$allIPInformationWorldWide = $NULL
$allIPInformationChina = $NULL
$allIPInfomrationUSGovGCCHigh = $NULL
$allIPInformationUSGovDOD = $NULL

$outputXMLFile = $global:LogFile.replace(".log",".xml")

$ipLocation = ""

$global:outputArray = @()

#Create the log file.

new-logfile -logFileName $logFileName -logFolderPath $logFolderPath

out-logfile -string $global:LogFile
out-logfile -string $outputXMLFile

#Start logging

out-logfile -string "*********************************************************************************"
out-logfile -string "Start Office365IPAddress"
out-logfile -string "*********************************************************************************"

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

out-logfile -string "Obtain IP information for all available Office 365 instances."

$allIPInformationWorldWide = get-Office365IPInformation -baseURL $allIPInformationWorldWideURL
$allIPInformationChina = get-Office365IPInformation -baseURL $allIPInformationChinaURL
$allIPInfomrationUSGovGCCHigh = get-Office365IPInformation -baseURL $allIPInformationUSGovGCCHighURL
$allIPInformationUSGovDOD = get-Office365IPInformation -baseURL $allIPInformationUSGovDODURL

out-logfile -string "Convert IP information from JSON."

$allIPInformationWorldWide = get-jsonData -data $allIPInformationWorldWide
$allIPInformationChina = get-jsonData -data $allIPInformationChina
$allIPInfomrationUSGovGCCHigh = get-jsonData -data $allIPInfomrationUSGovGCCHigh
$allIPInformationUSGovDOD = get-jsonData -data $allIPInformationUSGovDOD

out-logfile -string "Begin testing IP spaces for presence of the specified IP address."

test-IPSpace -dataToTest $allIPInformationWorldWide -IPAddress $IPAddressToTest
test-IPSpace -dataToTest $allIPInformationChina -IPAddress $IPAddressToTest
test-IPSpace -dataToTest $allIPInfomrationUSGovGCCHigh -IPAddress $IPAddressToTest
test-IPSpace -dataToTest $allIPInformationUSGovDOD -IPAddress $IPAddressToTest

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
    out-logfile -string ("The IP Address: "+$IPAddressToTest+ " was located in the following Office 365 Services:")

    if ($ipLocation -ne "Failed")
    {   
        out-logfile -string ("The IP Address geo-location is: "+$ipLocation.country)
    }

    foreach ($entry in $global:outputArray)
    {
        $entry
    }

    out-logfile -string "A XML file containing the above entries is available in the log directory."
    out-logfile -string "******************************************************"
    out-logfile -string "**"
    out-logfile -string "*"

    foreach ($entry in $global:outputArray)
    {
        out-logfile -string $entry
    }

    $global:outputArray | Export-Clixml -Path $outputXMLFile
}
else 
{
    out-logfile -string "******************************************************"
    out-logfile -string ("The IP Address: "+$IPAddressToTest+ " was NOT located in the following Office 365 Services.")
    out-logfile -string "******************************************************"
}

