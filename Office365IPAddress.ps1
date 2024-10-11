
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
    [string]$IPAddressToTest="0.0.0.0",
    [Parameter(Mandatory = $false)]
    [string]$URLToTest="nodomain.local",
    [Parameter(Mandatory = $false)]
    [string]$portToTest="0",
    [Parameter(Mandatory = $true)]
    [string]$logFolderPath=$NULL,
    [Parameter(Mandatory = $false)]
    [boolean]$allowQueryIPLocationInformationFromThirdParty=$TRUE
)

$ErrorActionPreference = 'Stop'

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

function create-OutputObject
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $M365Instance,
        [Parameter(Mandatory = $true)]
        $id,
        [Parameter(Mandatory = $true)]
        $ServiceArea,
        [Parameter(Mandatory = $true)]
        $ServiceAreaDisplayName,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $URLs,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $IPs,
        [Parameter(Mandatory = $true)]
        $IPInSubnetorURL,
        [Parameter(Mandatory = $true)]
        $TCPPorts,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $UDPPorts
        [Parameter(Mandatory = $true)]
        $ExpressRoute,
        [Parameter(Mandatory = $true)]
        $Required
    )
    
    $outputObject = new-Object psObject -property $([ordered]@{
        M365Instance = $M365Instance
        ID = $ID
        ServiceArea = $ServiceArea
        ServiceAreaDisplayName = $ServiceAreaDisplayName
        URLs = $URLs
        IPs = $ips
        IPInSubnetorURL = $IPInSubnetorURL
        TCPPorts = $tcpports
        UDPPorts = $udpPorts
        ExpressRoute = $expressRoute
        Required = $required
    })

    return $outputObject
}

function create-OutputChangebject
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $M365Instance,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $ChangeID,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Disposition,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $EndpointSetID,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Version,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $ServiceAreaDisplayName,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $IPsAddredorRemoved,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $URLsAddedOrRemoved,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $IPInSubnetOrURL,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $PreviousCategory,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $PreviousExpressRoute,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $PreviousServiceArea,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $PreviousRequire,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $PreviousTCPPort
    )
    
    $outputObject = new-Object psObject -property $([ordered]@{
        M365Instance = $M365Instance
        ChangeID = $ChangeID
        Disposition = $Disposition
        EndpointSetID = $EndpointSetID
        Version = $Version
        ServiceAreaDisplayName = $ServiceAreaDisplayName
        IPsAddedorRemoved = $IPsAddredorRemoved
        URLsAddedOrRemoved = $URLsAddedOrRemoved
        IPInSubnetorURL = $IPInSubnetOrURL
        PreviousCategory = $PreviousCategory
        PreviousExpressRoute = $PreviousExpressRoute
        PreviousServiceArea = $PreviousServiceArea
        PreviousRequire = $PreviousRequire
        PreviousTCPPorts = $PreviousTCPPort
    })

    return $outputObject
}


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
        $IPAddress,
        [Parameter(Mandatory = $true)]
        $portToTest,
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

                 $functionNetwork = get-IPEntry -ipEntry $ipEntry

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    if ($portToTest -ne 0)
                    {
                        out-logfile -string "The IP to test is contained within the entry.  Log the service."

                        $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required
                            
                        out-logfile -string $outputObject

                        $global:outputArray += $outputObject
                    }
                    else
                    {
                        if (($ipEntry.tcpPorts.contains($portToTest)) -or ($ipEntry.udpPorts.contains($portToTest)))
                        {
                            out-logfile -string "The IP to test is contained within the entry.  Log the service."

                            $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required
                                
                            out-logfile -string $outputObject

                            $global:outputArray += $outputObject
                        }
                        else {
                            out-logfile -string "A IP entry was found matching but not to the specified port - skipping."
                        }
                    }
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

function get-IPEntry
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        $ipEntry
    )

    $functionIPEntry = [System.Net.IpNetwork]::Parse($ipEntry)
    
    return $functionIPEntry
}

function calculate-WildCardURL
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $urlEntry,
        [Parameter(Mandatory = $true)]
        $urlToTest
    )

    $functionPeriod = "."
    $functionWildCard = "*"
    $functionSplitURLToTest = @()
    $functionSplitURLEntry =@()
    $functionSplit = @()
    $functionSplitURLEntryCount = 0
    $functionSplitURLToTestCount = 0
    
    out-logfile -string "The URL entry contains a wild card - rebuild the URL to test."

    out-logfile -string "URL to test split by period..."

    $functionSplitURLToTest = $urlToTest.split($functionPeriod)

    foreach ($member in $functionSplitURLToTest)
    {
        out-logfile -string $member
    }

    out-logfile -string "URL Entry split by period..."

    $functionSplitURLEntry = $urlEntry.split($functionPeriod)

    foreach ($member in $functionSplitURLEntry)
    {
        out-logfile -string $member
    }

    out-logfile -string "URL to test reverse split by period..."

    $functionSplitURLToTestReverse = $functionSplitURLToTest[-1..-($functionSplitURLToTest.Length)]
    
    foreach ($member in $functionSplitURLToTestReverse)
    {
        out-logfile -string $member
    }

    out-logfile -string "URL Entry reverse split by period..."

    $functionSplitURLEntryReverse = $functionSplitURLEntry[-1..-($functionSplitURLEntry.Length)]

    foreach ($member in $functionSplitURLEntryReverse)
    {
        out-logfile -string $member
    }

    $functionSplitURLEntryCount = $functionSplitURLEntryReverse.count
    out-logfile -string $functionSplitURLEntryCount.tostring()
    $functionSplitURLToTestCount = $functionSplitURLToTestReverse.count
    out-logfile -string $functionSplitURLToTestCount.tostring()

    if ($functionSplitURLToTestCount -gt $functionSplitURLEntryCount)
    {
        out-logfile -string "The URL to test is > the URL entry."

        for ($i = 0 ; $i -lt $functionSplitURLToTestReverse.count ; $i++)
        {
            if ($functionSplitURLEntryReverse[$i] -eq $functionWildCard)
            {
                out-logfile -string "Wild card character found - injecting and existing loop."
                $functionSplit += $functionWildCard
                $i = $functionSplitURLToTestReverse.count+1
            }
            else 
            {
                $functionSplit += $functionSplitURLToTestReverse[$i]
            }
        }
    }
    elseif ($functionSplitURLToTestCount -le $functionSplitURLEntryCount)
    {
        out-logfile -string "The URL to test is > the URL entry."

        for ($i = 0 ; $i -lt $functionSplitURLToTestReverse.count ; $i++)
        {
            if ($functionSplitURLEntryReverse[$i] -eq $functionWildCard)
            {
                out-logfile -string "Wild card character found - injecting and existing loop."
                $functionSplit += $functionWildCard
            }
            else 
            {
                $functionSplit += $functionSplitURLToTestReverse[$i]
            }
        }
    }

    foreach ($member in $functionSplit)
    {
        out-logfile -string $member
    }

    $functionSplitReverse = $functionSplit[-1..-($functionSplit.Length)]

    foreach ($member in $functionSplitReverse)
    {
        out-logfile -string $member
    }

    $functionTestURL=$functionSplitReverse[0]

    for ($i=1;$i -lt $functionSplitReverse.count ; $i++)
    {
        $functionTestURL = $functionTestURL + $functionPeriod
        $functionTestURL = $functionTestURL + $functionSplitReverse[$i]
    }

    out-logfile -string $functionTestURL

    return $functionTestURL
}

function test-URLSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $URLToTest,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionWildCard = "*"
    $functionPeriod = "."
    $functionSplitURLToTest = @()

    out-logfile -string "Entering test-URLSpace"

    foreach ($entry in $dataToTest)
    {
        Out-logfile -string ("Testing entry id: "+$entry.id)

        if ($entry.urls.count -gt 0)
        {
            out-logfile -string "URL count > 0"

            foreach ($urlEntry in $entry.URLs)
            {
                #If the URL entry is a wild card - rebuid the URL to test to contain the wild card.

                out-logfile -string "Determine if the url entry is a wild card URL."

                if ($urlEntry.contains($functionWildCard))
                {
                    $functionTestURL = calculate-WildCardURL -urlEntry $urlEntry -urlToTest $URLToTest
                    out-logfile -string $functionTestURL
                }
                else
                {
                    $functionTestURL = $URLToTest
                    out-logfile -string "The URL entry does not contain a wild card - use the URLToTest value."
                }

                if ($urlEntry -eq $functionTestURL)
                {
                    out-logfile -string "The url to test is contained within the entry.  Log the service."

                    $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $urlEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required

                    out-logfile -string $outputObject

                    $global:outputArray += $outputObject
                }
                else
                {
                out-logfile -string "The url to test is not contained within the entry - move on."
                }
            }
        }
        else 
        {
            out-logfile -string "url count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-urlSpace"
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

                 $functionNetwork = get-IPEntry -ipEntry $ipEntry

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

                    $outputObject = create-OutputChangebject -M365Instance $regionString -ChangeID $entry.ID -Disposition ($entry.Disposition+"-Add") -EndpointSetID $entry.endpointSetId -Version $entry.Version -ServiceAreaDisplayName $functionServiceAreaDisplayName -IPsAddredorRemoved $entry.add.ips -URLsAddedOrRemoved $entry.add.urls -IPInSubnetOrURL $ipEntry -PreviousCategory $entry.previous.Category -PreviousExpressRoute $entry.previous.expressRoute -PreviousServiceArea $entry.previous.serviceArea -PreviousRequire $entry.previous.required -PreviousTCPPort $entry.previous.tcpPorts

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

function test-URLChangeSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $changeDataToTest,
        [Parameter(Mandatory = $true)]
        $urlToTest,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionOriginalID = $null
    $functionServiceAreaDisplayName = $NULL
    $functionWildCard = "*"
    $functionPeriod = "."
    $functionSplitURLToTest = @()

    out-logfile -string "Entering test-URLChangeSpace"

    foreach ($entry in $changeDataToTest)
    {
        Out-logfile -string ("Testing change entry id: "+$entry.id)

        if ($entry.add.urls.count -gt 0)
        {
            out-logfile -string "URL count > 0"

            foreach ($urlEntry in $entry.add.urls)
            {
                out-logfile -string ("Testing entry url: "+$urlEntry)

                out-logfile -string "Determine if the url entry is a wild card URL."

                if ($urlEntry.contains($functionWildCard))
                {
                    $functionTestURL = calculate-WildCardURL -urlEntry $urlEntry -urlToTest $URLToTest
                    out-logfile -string $functionTestURL
                }
                else
                {
                    $functionTestURL = $URLToTest
                    out-logfile -string "The URL entry does not contain a wild card - use the URLToTest value."
                }

                if ($functionTestURL -eq $urlEntry)
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

                    $outputObject = create-OutputChangebject -M365Instance $regionString -ChangeID $entry.ID -Disposition ($entry.Disposition+"-Add") -EndpointSetID $entry.endpointSetId -Version $entry.Version -ServiceAreaDisplayName $functionServiceAreaDisplayName -IPsAddredorRemoved $entry.add.ips -URLsAddedOrRemoved $entry.add.urls -IPInSubnetOrURL $ipEntry -PreviousCategory $entry.previous.Category -PreviousExpressRoute $entry.previous.expressRoute -PreviousServiceArea $entry.previous.serviceArea -PreviousRequire $entry.previous.required -PreviousTCPPort $entry.previous.tcpPorts

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

                 $functionNetwork = get-IPEntry -ipEntry $ipEntry
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

                    $outputObject = create-OutputChangebject -M365Instance $regionString -ChangeID $entry.ID -Disposition ($entry.Disposition+"-Remove") -EndpointSetID $entry.endpointSetId -Version $entry.Version -ServiceAreaDisplayName $functionServiceAreaDisplayName -IPsAddredorRemoved $entry.remove.ips -URLsAddedOrRemoved $entry.remove.urls -IPInSubnetOrURL $ipEntry -PreviousCategory $entry.previous.Category -PreviousExpressRoute $entry.previous.expressRoute -PreviousServiceArea $entry.previous.serviceArea -PreviousRequire $entry.previous.required -PreviousTCPPort $entry.previous.tcpPorts

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

function test-URLRemoveSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $changeDataToTest,
        [Parameter(Mandatory = $true)]
        $urlToTest,
        [Parameter(Mandatory = $true)]
        $RegionString
    )

    $functionOriginalID = $null
    $functionServiceAreaDisplayName = $NULL
    $functionWildCard = "*"
    $functionPeriod = "."
    $functionSplitURLToTest = @()

    out-logfile -string "Entering test-URLChangeSpace"

    foreach ($entry in $changeDataToTest)
    {
        Out-logfile -string ("Testing change entry id: "+$entry.id)

        if ($entry.remove.urls.count -gt 0)
        {
            out-logfile -string "URL count > 0"

            foreach ($urlEntry in $entry.remove.urls)
            {
                out-logfile -string ("Testing entry url: "+$urlEntry)

                out-logfile -string "Determine if the url entry is a wild card URL."

                if ($urlEntry.contains($functionWildCard))
                {
                    $functionTestURL = calculate-WildCardURL -urlEntry $urlEntry -urlToTest $URLToTest
                    out-logfile -string $functionTestURL
                }
                else
                {
                    $functionTestURL = $URLToTest
                    out-logfile -string "The URL entry does not contain a wild card - use the URLToTest value."
                }

                if ($functionTestURL -eq $urlEntry)
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

                    $outputObject = create-OutputChangebject -M365Instance $regionString -ChangeID $entry.ID -Disposition ($entry.Disposition+"-Remove") -EndpointSetID $entry.endpointSetId -Version $entry.Version -ServiceAreaDisplayName $functionServiceAreaDisplayName -IPsAddredorRemoved $entry.remove.ips -URLsAddedOrRemoved $entry.remove.urls -IPInSubnetOrURL $urlEntry -PreviousCategory $entry.previous.Category -PreviousExpressRoute $entry.previous.expressRoute -PreviousServiceArea $entry.previous.serviceArea -PreviousRequire $entry.previous.required -PreviousTCPPort $entry.previous.tcpPorts

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

#=====================================================================================
#Begin main function body.
#=====================================================================================

#Define function variables.

$noIPSpecified = "0.0.0.0"
$noURLSpecified = "nodomain.local"
$urlSlashes = "//"
$urlSlash = "/"
$functionURL
$functionDomainName

if (($IPToTest -ne $noIPSpecified) -and ($URLToTest -ne $noURLSpecified))
{
    write-error "Specify either a URL or IP to test - do not specify both in the same command."
}

if ($IPAddressToTest -ne $noIPSpecified)
{
    write-host "IPAddress to test is not a null value - proceed with evaluation"

    if ($IPAddressToTest.contains("."))
    {
        $logFileName = $IPAddressToTest.replace(".","-")

        if (IsIPv4AddressValid -ip $ipAddressToTest)
        {
            write-host "IPv4 Address is in a valid format."
        }
        else 
        {
            write-error "IPv4 address specified and is not in a valid format."
        } 
    }
    else 
    {
        $logFileName = $IPAddressToTest.replace(":","-")

        if (IsIPv6AddressValid -ip $ipAddressToTest)
        {
            write-host "IPv6 address specified and is in proper format."
        }
        else 
        {
             write-error "IPv6 address specified and is not in a valid format."
        } 
    }
}
elseif ($URLToTest -ne $noURLSpecified)
{
    write-host "URL to test is not a null value - proceed with evaluation."

    write-host "Determine if the URL is specified as a domain name <or> web address."

    if ($urlToTest.contains($urlSlashes))
    {
        write-host "URL specified - break URL"

        $functionURL = $urlToTest.split($urlSlashes)

        foreach ($member in $functionURL)
        {
            write-host $member
        }

        write-host "Determine if more of the URL is specified."

        if ($functionURL[1].contains($urlSlash))
        {
            write-host "URL contains more information - break URL."

            $functionURL = $functionURL.split($urlSlash)

            foreach ($member in $functionURL)
            {
                write-host $member
            }

            write-host "Domain name represented in array location 1."

            write-host "Create the log directory based off domain name."
        }
        else 
        {
            write-host "No more of a URL is specified - domain in array position 1."
        }

        $functionDomainName = $functionURL[1]
        write-host $functionDomainName
    }
    else 
    {
        write-host "URL appears to only contain a domain name."

        $functionDomainName = $URLToTest
        $functionDomainName.replace(".","-")
    }

    $logFileName = $functionDomainName.replace(".","-")
}
else 
{
   write-host "Using a generic log file name."

   $logFileName = "Office365IPandURLTest"
}


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

Test-PowerShellVersion

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

if ($IPAddressToTest -ne $noIPSpecified)
{
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
        else 
        {
            out-logfile -string "Failed to determine the IP address location."
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
}
elseif ($urlToTest -ne $noURLSpecified)
{
    out-logfile -string "Begin testing URL spaces for presence of the specified IP address."

    test-URLSpace -dataToTest $allIPInformationWorldWide -urlToTest $functionDomainName -regionString $worldWideRegionString
    test-URLSpace -dataToTest $allIPInformationChina -urlToTest $functionDomainName -regionString $chinaRegionString
    test-URLSpace -dataToTest $allIPInfomrationUSGovGCCHigh -urlToTest $functionDomainName -regionString $gccHighRegionString
    test-URLSpace -dataToTest $allIPInformationUSGovDOD -urlToTest $functionDomainName -regionString $dodRegionString

    if ($global:outputArray.count -gt 0)
    {
        out-logfile -string "IPs found in Active Service - test for changes."
    
        test-URLChangeSpace -dataToTest $allIPInformationWorldWide -urlToTest $urlToTest -regionString $worldWideRegionString -changeDataToTest $allIPChangeInformationWorldWide
        test-URLChangeSpace -dataToTest $allIPInformationChina -urlToTest $urlToTest -regionString $chinaRegionString -changeDataToTest $allIPChangeInformationChina
        test-URLChangeSpace -dataToTest $allIPInfomrationUSGovGCCHigh -urlToTest $urlToTest -regionString $gccHighRegionString -changeDataToTest $allIPChangeInfomrationUSGovGCCHigh
        test-URLChangeSpace -dataToTest $allIPInformationUSGovDOD -urlToTest $urlToTest -regionString $dodRegionString -changeDataToTest $allIPChangeInformationUSGovDOD
    }
    else 
    {
        out-logfile -string "No IP addresses found -> no need to test for additions."
    }

    if ($global:outputArray.count -eq 0)
    {
        out-logfile -string "Since the global output count is equal to 0 -> testing to see if the IP address was removed."
    
        test-URLRemoveSpace -dataToTest $allIPInformationWorldWide -urlToTest $URLToTest -regionString $worldWideRegionString -changeDataToTest $allIPChangeInformationWorldWide
        test-URLRemoveSpace -dataToTest $allIPInformationChina -urlToTest $URLToTest -regionString $chinaRegionString -changeDataToTest $allIPChangeInformationChina
        test-URLRemoveSpace -dataToTest $allIPInfomrationUSGovGCCHigh -urlToTest $URLToTest -regionString $gccHighRegionString -changeDataToTest $allIPChangeInfomrationUSGovGCCHigh
        test-URLRemoveSpace -dataToTest $allIPInformationUSGovDOD -urlToTest $URLToTest -regionString $dodRegionString -changeDataToTest $allIPChangeInformationUSGovDOD
    }
    else 
    {
        out-logfile -string "The IP address was found - no need to test for removals."
    }

    if ($global:outputArray.count -gt 0)
    {
        out-logfile -string "*"
        out-logfile -string "**"
        out-logfile -string "******************************************************"
            
        write-host "URL entries present in the following Office 365 Services:"
    
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
            out-logfile -string "URLs was located in the following service description:"
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
            out-logfile -string "URL was located in the following version additions since 2018:"
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
        out-logfile -string ("The URL: "+$URLToTest+ " was NOT located in the following Office 365 Services.")
        out-logfile -string "******************************************************"

        out-logfile -string $global:outputRemoveArray.count.tostring()
    
        if ($global:outputRemoveArray.count -gt 0)
        {
            write-host "The following changes removed the URL:"
    
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
}