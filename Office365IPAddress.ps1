
<#PSScriptInfo

.VERSION 1.2.7

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
    [Parameter(Mandatory = $false)]
    [boolean]$includeAzureSearch = $false,
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
        [AllowEmptyString()]
        [AllowNull()]
        $TCPPorts,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $UDPPorts,
        [Parameter(Mandatory = $true)]
        $ExpressRoute,
        [Parameter(Mandatory = $true)]
        $Required,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Notes,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Category
    )

    out-logfile -string "Entering create output object..."
    
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
        Notes = $notes
        Category = $Category
    })

    out-logfile -string "Exiting create output object..."

    return $outputObject
}

function create-AzureutputObject
{
    Param
    (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $OverallChangeNumber,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Cloud,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Name,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $ID,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $ChangeNumber,
        [Parameter(Mandatory = $true)]
        $IPInSubnet,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Region,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $RegionID,
        [Parameter(Mandatory = $true)]
        $Platform,
        [Parameter(Mandatory = $true)]
        $SystemService,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $AddressPrefixes,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowNull()]
        $NetworkFeatures
    )

    out-logfile -string "Entering create output object..."
    
    $outputObject = new-Object psObject -property $([ordered]@{
        OverallChangeNumber = $OverallChangeNumber
        Cloud = $Cloud
        Name = $Name
        ID = $ID
        IPInSubnet = $IPInSubnet
        ChangeNumbe = $ChangeNumber
        Region = $Region
        RegionID = $RegionID
        Platform = $Platform
        SystemService = $SystemService
        AddressPrefixes = $AddressPrefixes
        NetworkFeatures = $NetworkFeatures
    })

    out-logfile -string "Exiting create output object..."

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

    out-logfile -string "Entering create output change object..."
    
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

    out-logfile -string "Exiting create output change object..."

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

    out-logfile -string "Entering get client guid..."

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

    out-logfile -string "Exiting get client guid..."

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
    $functionComma = ","

    out-logfile -string "Entering test-IPSpace"

    foreach ($entry in $dataToTest)
    {
        Out-logfile -string ("Testing entry id: "+$entry.id)

        if ($entry.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.ips)
            {
                $functionPortArray = @()

                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = get-IPEntry -ipEntry $ipEntry

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    if ($portToTest -eq "0")
                    {
                        out-logfile -string "The IP to test is contained within the entry.  Log the service."

                        $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                            
                        out-logfile -string $outputObject

                        $global:outputArray += $outputObject
                    }
                    else
                    {
                        out-logfile -string "Extracting TCP Ports from entry..."

                        if ($entry.tcpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.tcpPorts.split($functionComma)
                        }

                        out-logfile -string "Extracting UDP Ports from entry..."
                        
                        if ($entry.udpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.udpPorts.split($functionComma)
                        }

                        out-logfile -string "Selecting unique ports..."

                        $functionPortArray = $functionPortArray | select-object -Unique

                        out-logfile -string "Removing trailing or leading spaces from any port..."

                        for ($i = 0 ; $i -lt $functionPortArray.count ; $i++)
                        {
                            $functionPortArray[$i] = $functionPortArray[$i].replace(" ","")
                            out-logfile -string $functionPortArray[$i]
                        }

                        if ($functionPortArray.contains($portToTest))
                        {
                            out-logfile -string "The IP to test is contained within the entry.  Log the service."

                            $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                                
                            out-logfile -string $outputObject

                            $global:outputArray += $outputObject
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

function test-AzureIPSpace
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToTest,
        [Parameter(Mandatory = $true)]
        $IPAddress
    )

    $functionNetwork = $NULL
    $functionOverallChangeNumber = $dataToTest.changeNumber
    $functionCloud = $dataToTest.cloud

    out-logfile -string "Entering test-AzureIPSpace"

    out-logfile -string ("Function overall change number: "+$functionOverallChangeNumber.tostring()) 
    out-logfile -string ("Function cloud: "+$functionCloud)

    foreach ($entry in $dataToTest.values)
    {
        $functionName = $entry.name
        $functionID = $entry.id

        out-logfile -string ("Entry Name: "+$functionName)
        out-logfile -string ("Entry ID: "+$functionID)

        foreach ($property in $entry.properties)
        {
            $functionChangeNumber = $property.changeNumber.tostring()
            $functionRegion = $property.Region
            $functionRegionID = $property.RegionID
            $functionPlatform = $property.Platform
            $functionSystemService = $property.systemService
            $functionAddressPrefixes = $property.AddressPrefixes
            $functionNetworkFeatures = $property.networkFeatures

            if ($functionChangeNumber -ne $NULL)
            {
                out-logfile -string ("Entry change number: "+$functionChangeNumber)
            }
            else
            {
                $functionChangeNumber = ""
            }

            if ($functionRegion -ne $NULL)
            {
                out-logfile -string ("Entry region: "+$functionRegion)
            }
            else
            {
                $functionRegion = ""
            }

            if ($functionRegionID -ne $NULL)
            {
                out-logfile -string ("Entry region id: "+$functionRegionID)
            }
            else
            {
                $functionRegionID = ""
            }

            if ($functionPlatform -ne $NULL)
            {
                out-logfile -string ("Entry platform: "+$functionPlatform)
            }
            else
            {
                $functionPlatform = ""
            }

            if ($functionSystemService -ne $NULL)
            {
                out-logfile -string ("Entry system service: "+$functionSystemService)
            }
            else
            {
                $functionSystemService = ""
            }

            if ($functionAddressPrefixes -ne $NULL)
            {
                out-logfile -string ("Entry address prefixes: "+$functionAddressPrefixes)
            }
            else
            {
                $functionAddressPrefixes = ""
            }

            if ($functionNetworkFeatures -ne $NULL)
            {
                out-logfile -string ("Entry network features: "+$functionNetworkFeatures)     
            }
            else
            {
                $functionNetworkFeatures = ""
            }

            if ($property.addressPrefixes.count -gt 0)
            {
                out-logfile -string "There are IP addresses to test."

                foreach ($ipEntry in $property.addressPrefixes)
                {
                    out-logfile -string ("Testing IP entry: "+$ipEntry)

                    $functionNetwork = get-IPEntry -ipEntry $ipEntry

                    if ($functionNetwork.Contains($IPAddress))
                    {
                        out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                        $outputObject = create-AzureutputObject -OverallChangeNumber $functionOverallChangeNumber -Cloud $functionCloud -Name $functionName -id $functionID -changeNumber $functionChangeNumber -region $functionRegion -regionID $functionRegionID -platform $functionPlatform -systemService $functionSystemService -AddressPrefixes $functionAddressPrefixes -networkFeatures $functionNetworkFeatures -IPInSubnet $ipEntry

                        out-logfile -string $outputObject

                        $global:outputAzureArray += $outputObject
                    }
                }
            }
        }
    }

    <#

    foreach ($entry in $dataToTest)
    {
        Out-logfile -string ("Testing entry id: "+$entry.id)

        if ($entry.ips.count -gt 0)
        {
            out-logfile -string "IP count > 0"

            foreach ($ipEntry in $entry.ips)
            {
                $functionPortArray = @()

                out-logfile -string ("Testing entry IP: "+$ipEntry)

                 $functionNetwork = get-IPEntry -ipEntry $ipEntry

                 out-logfile -string ("BaseAddress: "+$functionNetwork.baseAddress+ " PrefixLength: "+$functionNetwork.PrefixLength)

                 if ($functionNetwork.Contains($IPAddress))
                 {
                    if ($portToTest -eq "0")
                    {
                        out-logfile -string "The IP to test is contained within the entry.  Log the service."

                        $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                            
                        out-logfile -string $outputObject

                        $global:outputArray += $outputObject
                    }
                    else
                    {
                        out-logfile -string "Extracting TCP Ports from entry..."

                        if ($entry.tcpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.tcpPorts.split($functionComma)
                        }

                        out-logfile -string "Extracting UDP Ports from entry..."
                        
                        if ($entry.udpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.udpPorts.split($functionComma)
                        }

                        out-logfile -string "Selecting unique ports..."

                        $functionPortArray = $functionPortArray | select-object -Unique

                        out-logfile -string "Removing trailing or leading spaces from any port..."

                        for ($i = 0 ; $i -lt $functionPortArray.count ; $i++)
                        {
                            $functionPortArray[$i] = $functionPortArray[$i].replace(" ","")
                            out-logfile -string $functionPortArray[$i]
                        }

                        if ($functionPortArray.contains($portToTest))
                        {
                            out-logfile -string "The IP to test is contained within the entry.  Log the service."

                            $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $ipEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                                
                            out-logfile -string $outputObject

                            $global:outputArray += $outputObject
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

    #>

    out-logfile -string "Exiting test-AzureIPSpace"
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
        out-logfile -string "The URL to test is < the URL entry."

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

    out-logfile -string "Returning URL to test..."
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
    $functionComma = ","
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
                $functionPortArray = @()
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
                    if ($portToTest -eq "0")
                    {
                        out-logfile -string "The URL to test is contained within the entry.  Log the service."

                        $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $urlEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                            
                        out-logfile -string $outputObject

                        $global:outputArray += $outputObject
                    }
                    else
                    {
                        out-logfile -string "Extracting TCP Ports from entry..."

                        if ($entry.tcpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.tcpPorts.split($functionComma)
                        }

                        out-logfile -string "Extracting UDP Ports from entry..."
                        
                        if ($entry.udpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.udpPorts.split($functionComma)
                        }

                        foreach ($member in $functionPortArray)
                        {
                            out-logfile -string $member
                        }

                        out-logfile -string "Selecting unique ports..."

                        $functionPortArray = $functionPortArray | select-object -Unique

                        for ($i = 0 ; $i -lt $functionPortArray.count ; $i++)
                        {
                            $functionPortArray[$i] = $functionPortArray[$i].replace(" ","")
                            out-logfile -string $functionPortArray[$i]
                        }

                        if ($functionPortArray.contains($portToTest))
                        {
                            out-logfile -string "The URL to test is contained within the entry.  Log the service."

                            $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $urlEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                                
                            out-logfile -string $outputObject

                            $global:outputArray += $outputObject
                        }
                    }
                }
                elseif ($urlEntry.contains($functionTestURL))
                {
                    if ($portToTest -eq "0")
                    {
                        out-logfile -string "The URL to test is contained within the entry.  Log the service."

                        $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $urlEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                            
                        out-logfile -string $outputObject

                        $global:outputArray += $outputObject
                    }
                    else
                    {
                        out-logfile -string "Extracting TCP Ports from entry..."

                        if ($entry.tcpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.tcpPorts.split($functionComma)
                        }

                        out-logfile -string "Extracting UDP Ports from entry..."
                        
                        if ($entry.udpPorts -ne $NULL)
                        {
                            $functionPortArray += $entry.udpPorts.split($functionComma)
                        }

                        foreach ($member in $functionPortArray)
                        {
                            out-logfile -string $member
                        }

                        out-logfile -string "Selecting unique ports..."

                        $functionPortArray = $functionPortArray | select-object -Unique

                        for ($i = 0 ; $i -lt $functionPortArray.count ; $i++)
                        {
                            $functionPortArray[$i] = $functionPortArray[$i].replace(" ","")
                            out-logfile -string $functionPortArray[$i]
                        }

                        if ($functionPortArray.contains($portToTest))
                        {
                            out-logfile -string "The URL to test is contained within the entry.  Log the service."

                            $outputObject = create-outputObject -m365Instance $regionString -id $entry.id -serviceArea $entry.serviceArea -serviceAreaDisplayName $entry.serviceareadisplayname -urls $entry.urls -ips $entry.ips -ipInSubnetorURL $urlEntry -tcpPorts $entry.tcpPorts -udpPorts $entry.udpPorts -expressRoute $entry.expressRoute -required $entry.required -notes $entry.notes -category $entry.category
                                
                            out-logfile -string $outputObject

                            $global:outputArray += $outputObject
                        }
                    }
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

    out-logfile -string "Entering test-IPChangeSpace"

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

    out-logfile -string "Exiting test-IPChangeSpace"
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
                    out-logfile -string "The URL to test is contained within the entry.  Log the service."

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
                    out-logfile -string "The URL to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "URL count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-URLChangeSpace"
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

    out-logfile -string "Entering test-IPRemoveSpace"

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

    out-logfile -string "Exiting test-IPRemoveSpace"
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

    out-logfile -string "Entering test-URLRemoveSpace"

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
                    out-logfile -string "The URL to test is contained within the entry.  Log the service."

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
                    out-logfile -string "The URL to test is not contained within the entry - move on."
                 }
            }
        }
        else 
        {
            out-logfile -string "URL count = 0 -> skipping"
        }
    }

    out-logfile -string "Exiting test-urlRemoveSpace"
}

function export-JSONInformation
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $dataToExport,
        [Parameter(Mandatory = $true)]
        $exportPath
    )

    out-logfile -string "Entering export-JSONInformation"

    try
    {
        $dataToExport | Export-Clixml $exportPath -errorAction STOP
    }
    catch {
        out-logfile -string "Unable to export xml data to the directory."
        out-logfile -string $_ -isError:$true
    }

    out-logfile -string "Exiting export-JSONInformation"
}

function get-IPLocationInformation
{
    Param
    (
        [Parameter(Mandatory = $true)]
        $ipAddress
    )

    $functionData = ""
    <#
    $global:ipLocationProvider = "https://api.country.is/"
    $global:ipLocationQuery = $global:ipLocationProvider+$ipAddress
    #>

    $global:ipLocationProvider = "ipinfo.io/"
    $global:ipLocationQuery = $global:ipLocationProvider+"/"+$ipAddress+"/json"

    out-logfile -string "Entering get-IPLocationInformation"

    try {
        #$functionData = invoke-WebRequest $global:ipLocationQuery
        $functionData = invoke-RestMethod $global:ipLocationQuery -errorAction STOP
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

Function Test-PowershellModule
{
    out-logfile -string "Entering Test-PowerShellModule"

    $functionModuleName = "PSWritehHTML"

    try {
        import-module "PSWriteHTML" -errorAction STOP
    }
    catch {
        out-logfile -string "Please install the PSWriteHTML Powershell Module using install-module PSWriteHTML" -isError:$true
    }

    out-logfile -string "Exiting Test-PowerShellModule"

}

Function get-AzureData
{
    param(
        [Parameter(Mandatory = $true)]
        $dataLocation
    )

    $functionData = $NULL

    out-logfile -string "Entering get-AzureData"

    out-logfile -string $dataLocation

    try {
        $functionData = Import-Clixml $dataLocation -errorAction STOP
    }
    catch {
        out-logfile -string "Unable to import the pre-gathered Azure IP information files."
        out-logfile -string "These files are stored in the AzureIPAddress folder created in the log file direactory."
        out-logfile -string "To enable this function run install-script AzureIPAddress"
        out-logfile -string "Run AzureIPAddress.ps1 -logFolderPath c:\Something where c:\something is the same log folder path used with this script."
        out-logfile -string $_ -isError:$TRUE
    }

    out-logfile -string "Entering get-AzureData"

    return $functionData
}

Function generate-HTMLData
{
    param(
        [Parameter(Mandatory = $true)]
        $IPorURLToTest
    )

    $functionHTMLSuffix = ".html"
    $functionLogSuffix = ".log"
    $functionHTMLFile = $global:logFile.replace($functionLogSuffix,$functionHTMLSuffix)
    $functionTitle = ("IP or URL Report: "+$IPorURLToTest)

    out-logfile -string "Entering generate-HTMLData"

    new-html -TitleText $IPorURLToTest -filePath $functionHTMLFile {
        new-htmlHeader {
            New-HTMLText -text $functionTitle -fontsize 24 -Color Black -Alignment center
        }
        New-HTMLMain {
            New-HTMLTableOption -DataStore JavaScript

            if ($global:ipLocation -ne "Failed")
            {
                #new-HTMLSection -HeaderText ("IP Location Lookup: "+$global:ipLocation.country) {
                new-HTMLSection -HeaderText ("IP Location Information") {
                    New-HTMLList{
                                new-htmlListItem -text ("IP Address: "+$global:ipLocation.ip) -fontSize 14
                                new-htmlListItem -text ("HostName: "+$global:ipLocation.hostname) -fontSize 14
                                new-htmlListItem -text ("City: "+$global:ipLocation.city) -fontSize 14
                                new-htmlListItem -text ("Region: "+$global:ipLocation.region) -fontSize 14
                                new-htmlListItem -text ("County: "+$global:ipLocation.country) -fontSize 14
                                new-htmlListItem -text ("Location: "+$global:ipLocation.loc) -fontSize 14
                                new-htmlListItem -text ("Organization: "+$global:ipLocation.org) -fontSize 14
                                new-htmlListItem -text ("Postal: "+$global:ipLocation.postal) -fontSize 14
                            }
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Blue"  -CanCollapse -BorderRadius 10px -collapsed
            }

            if (($global:outputArray.count -gt 0) -or ($global:outputChangeArray.count -gt 0) -or ($global:outputRemoveArray.count -gt 0) -or ($global:outputAzureArray.count -gt 0))
            {
                if ($global:outputArray.count -gt 0)
                {
                    new-htmlSection -HeaderText "IP or URL Entries in Office 365"{
                        new-htmlTable -dataTable $global:outputArray
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Blue"  -CanCollapse -BorderRadius 10px -collapsed
                }
                else{
                    new-htmlSection -HeaderText "No IP or URL Entries in Office 365"{
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red" -BorderRadius 10px 
                }

                if ($global:outputChangeArray.count -gt 0)
                {
                    new-htmlSection -HeaderText "IP or URL Change Entries in Office 365"{
                        new-htmlTable -dataTable $global:outputChangeArray
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Blue"  -CanCollapse -BorderRadius 10px -collapsed
                }
                else
                {
                    new-htmlSection -HeaderText "No IP or URL Change Entries in Office 365"{
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red" -BorderRadius 10px 
                }

                if ($global:outputRemoveArray.count -gt 0)
                {
                    new-htmlSection -HeaderText "IP or URL Remove Entries in Office 365"{
                        new-htmlTable -dataTable $global:outputRemoveArray
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Blue"  -CanCollapse -BorderRadius 10px -collapsed
                }
                else
                {
                    new-htmlSection -HeaderText "No IP or URL Remove Entries in Office 365"{
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red" -BorderRadius 10px 
                }

                if ($global:outputAzureArray.count -gt 0)
                {
                    new-htmlSection -HeaderText "IP Entries in Azure Services"{
                        new-htmlTable -dataTable $global:outputAzureArray
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Blue"  -CanCollapse -BorderRadius 10px -collapsed
                }
                else
                {
                    new-htmlSection -HeaderText "No IP Entries in Azure Services"{
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red" -BorderRadius 10px 
                }
            }
            else 
            {
                new-htmlSection -HeaderText "No URL or IP Information in Office 365 or Azure"{
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red" -BorderRadius 10px 
            }
        }
    } -online -ShowHTML

    out-logfile -string "Exiting generate-HTMLData"
}

Function get-AzureData
{
    param(
        [Parameter(Mandatory = $true)]
        $dataLocation
    )

    $functionData = $NULL

    out-logfile -string "Entering get-AzureData"

    out-logfile -string $dataLocation

    try {
        $functionData = Import-Clixml $dataLocation -errorAction STOP
    }
    catch {
        out-logfile -string "Unable to import the pre-gathered Azure IP information files."
        out-logfile -string "These files are stored in the AzureIPAddress folder created in the log file direactory."
        out-logfile -string "To enable this function run install-script AzureIPAddress"
        out-logfile -string "Run AzureIPAddress.ps1 -logFolderPath c:\Something where c:\something is the same log folder path used with this script."
        out-logfile -string $_ -isError:$TRUE
    }

    out-logfile -string "Entering get-AzureData"

    return $functionData
}

Function get-AzureIPInformation
{
    param(
        [Parameter(Mandatory = $true)]
        $logFolderPath
    )

    $azureLogName = "AzureIPAddress.log"
    $azureFolderName = "AzureIPAddress"
    $azureFilePath = $logFolderPath + "\" + $azureFolderName + "\" + $azureLogName
    $azureScriptName = "AzureIPAddress.ps1"
    $azureScriptString = " -logfolderPath "
    $processName = "Powershell.exe"
    $azureLog = $null

    try {
            #$job = Start-Job -ScriptBlock { AzureIPAddress.ps1 -logFolderPath $args[0] } -PSVersion 5.1 -ArgumentList $logFolderPath -errorAction Stop
            $jobString = $azureScriptName + $azureScriptString + $logFolderPath
            out-logfile -string $jobString
            Start-Process $processName $jobString -Wait -errorAction Stop
    }
    catch {
        Out-logfile -string "Unable to invoke AzureIPAddress.ps1 script."
        out-logfile -string $_ -isError:$true
    }

    <#

    out-logfile -string "AzureIPAddress.ps1 job invoked successfully."

    try {
            Wait-Job $job -ErrorAction Stop
    }
    catch {
        out-logfile -string "Unable to wait for job to complete."
        out-logfile -string $_ -isError:$true
    }

    out-logfile -string "Successfully waited for job to complete."

    #>

    try {
            $azureLog = Get-Content $azureFilePath -ErrorAction STOP
    }
    catch {
        out-logfile -string "Unable to capture the AzureIPAddress log file."
        out-logfile -string $_ -isError:$true
    }

    try {
        $azureLog | Out-File -FilePath $global:LogFile -Append -ErrorAction Stop
    }
    catch {
        out-logfile -string "Unable to append the AzureIPAddress log file to the current log."
        out-logfile -string $_
    }

    out-logfile -string "AzureIPAddress log appended successfully to current log."

    <#

    if ($job.State -eq 'Failed') {
        out-logfile -string "Unable to utilize the script AzureIPAddress.ps1 to download Azure IP address information for verification."
        out-logfile -string "Please run Install-Script AzureIPAddress to ensure the script is available."
        out-logfile -string "If the script is installed refer to the AzureIPAddress.log contained in the log file specified for this command."
        out-logfile -string "Unable to obtain AzureIPAddress information." -isError:$true
    } else {
        out-logfile -string "AzureIPAddress.ps1 invoked successfully."
    }

    try {
        Remove-Job $job -errorAction Stop
    }
    catch {
        out-logfile -string "Unable to remove job successfully."
        out-logfile -string $_ -isError:$true
    }

    out-logfile -string "Job removed successfully."
    #>
}

function test-AzureScript
{
     param(
        [Parameter(Mandatory = $true)]
        $logFolderPath
    )

    $requiredScriptVersion = "1.8"

    try {
        $version = Invoke-Command -ScriptBlock {AzureIPAddress.ps1 -logFolderPath $args[0] -versionTest $args[1] -errorAction Stop} -ArgumentList $logFolderPath,$true
    }
    catch {
        out-logfile -string "Error testing Azure script version."
        out-logfile -string "From Powershell 5 please run Install-Script AzureIPAddress or Update-Script AzureIPAddress if already installed."
        out-logfile -string $_ -isError:$true
    }

    if ($version -ne $requiredScriptVersion)
    {  
        out-logfile -string $version
        out-logfile -string ("AzureIPAddress must be version "+$requiredScriptVersion+" to utilize script.")
        out-logfile -string "Run Update-Script AzureIPAddress from Powershell5 to ensure up to date version." -isError:$TRUE
    }
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

if (($urlToTest -ne $noURLSpecified) -and ($includeAzureSearch -eq $True))
{
    write-error "Azure search does not support URL - specify only an IP address."
}

if (($portToTest -ne 0) -and ($includeAzureSearch -eq $TRUE))
{
    write-error "Azure search does not support port to test."
}

if (($IPAddressToTest -ne $noIPSpecified) -and ($URLToTest -ne $noURLSpecified))
{
    write-error "Specify either a URL or IP to test - do not specify both in the same command."
}

if ($IPAddressToTest -ne $noIPSpecified)
{
    write-host "IPAddress to test is not a null value - proceed with evaluation"

    if ($IPAddressToTest.contains("."))
    {
        $logFileName = $IPAddressToTest.replace(".","-")

        write-host $logFileName

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

        write-host $logFileName

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

        write-host "Testing to see if a wild card URL was specified."

        if ($functionDomainName[0] -eq "*")
        {
            write-host "A wild card URL was specified, replacing wild card."

            $functionDomainName=$functionDomainName.replace("*.","")
        }
        else 
        {
            write-host "A wild card URL was not specified."
        }

        write-host $functionDomainName
    }

    $logFileName = $functionDomainName.replace(".","-")
    write-host $logFileName
}
else 
{
   write-host "Using a generic log file name."

   $logFileName = "Office365IPandURLTest"
   write-host $logFileName
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

$azureDataStaticPath = "\AzureIPAddress"
$azurePublicCloudXML = $logFolderPath+$azureDataStaticPath+"\AzureIPAddress-Public.xml"
$azureGovernmentCloudXML = $logFolderPath+$azureDataStaticPath+"\AzureIPAddress-Government.xml"
$azurePublicData = $null
$azureGovernmentData = $null

$global:ipLocation = "Failed"

$global:outputArray = @()
$global:outputChangeArray=@()
$global:outputRemoveArray=@()
$global:outputAzureArray=@()

#Create the log file.

new-logfile -logFileName $logFileName -logFolderPath $logFolderPath

$logFileXMLString = $logFileName+".log"

$allIPInformationWorldWideXML = $global:logfile.replace($logFileXMLString,"WorldWide.xml")
$allIPInformationChinaXML = $global:logfile.replace($logFileXMLString,"China.xml")
$allIPInformationUSGovGCCHighXML = $global:logfile.replace($logFileXMLString,"GCCHigh.xml")
$allIPInformationUSGovDODXML = $global:logfile.replace($logFileXMLString,"DOD.xml")
$allIPChangeInformationWorldWideXML = $global:logfile.replace($logFileXMLString,"WorldWide-Change.xml")
$allIPChangeInformationChinaXML = $global:logfile.replace($logFileXMLString,"China-Change.xml")
$allIPChangeInformationUSGovGCCHighXML = $global:logfile.replace($logFileXMLString,"GCCHigh-Change.xml")
$allIPChangeInformationUSGoveDODXML = $global:logfile.replace($logFileXMLString,"DOD-Change.xml")
$allIPAzurePublic = $global:logfile.replace($logFileXMLString,"AzurePublic.xml")
$allIPAzureGovernment = $global:logfile.replace($logFileXMLString,"AzureGovernment.xml")

out-logfile -string "*********************************************************************************"
out-logfile -string "Start Office365IPAddress"
out-logfile -string "*********************************************************************************"

if ($includeAzureSearch -eq $TRUE)
{
    out-logfile -string "Test required azure script version..."

    Test-AzureScript -logFolderPath $logFolderPath
}

out-logfile -string $global:LogFile
out-logfile -string $logFileName
out-logfile -string $IPAddressToTest
out-logfile -string $URLToTest
out-logfile -string $portToTest

out-logfile -string $allIPInformationWorldWideXML
out-logfile -string $allIPInformationChinaXML
out-logfile -string $allIPInformationUSGovGCCHighXML
out-logfile -string $allIPInformationUSGovDODXML
out-logfile -string $allIPChangeInformationWorldWideXML
out-logfile -string $allIPChangeInformationChinaXML
out-logfile -string $allIPChangeInformationUSGovGCCHighXML
out-logfile -string $allIPChangeInformationUSGoveDODXML
out-logfile -string $allIPAzurePublic
out-logfile -string $allIPAzureGovernment
out-logfile -string $azurePublicCloudXML
out-logfile -string $azureGovernmentCloudXML

$outputXMLFile = $global:LogFile.replace(".log",".xml")
$outputChangeXMLFile = $global:LogFile.replace(".log","-Adds.xml")
$outputRemoveXMLFile = $global:LogFile.replace(".log","-Removes.xml")
$outputAzureXMLFile = $global:LogFile.replace(".log","-Azure.xml")

out-logfile -string $outputXMLFile
out-logfile -string $outputChangeXMLFile
out-logfile -string $outputRemoveXMLFile
out-logfile -string $outputAzureXMLFile

#Start logging

out-logfile -string "Testing powershell version..."

Test-PowerShellVersion

Test-PowershellModule

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

out-logfile -string "Export all gathered json information to the log directory for further review."

export-JSONInformation -dataToExport $allIPInformationWorldWide -exportPath $allIPInformationWorldWideXML
export-JSONInformation -dataToExport $allIPInformationChina -exportPath $allIPInformationChinaXML
export-JSONInformation -dataToExport $allIPInfomrationUSGovGCCHigh -exportPath $allIPInformationUSGovGCCHighXML
export-JSONInformation -dataToExport $allIPInformationUSGovDOD -exportPath $allIPInformationUSGovDODXML
export-JSONInformation -dataToExport $allIPChangeInformationWorldWide -exportPath $allIPChangeInformationWorldWideXML
export-JSONInformation -dataToExport $allIPChangeInformationChina -exportPath $allIPChangeInformationChinaXML
export-JSONInformation -dataToExport $allIPChangeInfomrationUSGovGCCHigh -exportPath $allIPChangeInformationUSGovGCCHighXML
export-JSONInformation -dataToExport $allIPChangeInformationUSGovDOD -exportPath $allIPChangeInformationUSGoveDODXML

out-logfile -string "Determine if it is necessary to gather Azure IP information and parse the data."

if ($includeAzureSearch -eq $TRUE)
{
    out-logfile -string "Obtain the Azure IP address information."

    out-logfile -string "***A new process window will open to service obtaining Azure IP Information in 5 seconds.***"
    out-logfile -string "***Do not change focus from this Window, the process will exist automatically.***"

    start-sleep -s 5

    Get-AzureIPInformation -logFolderPath $logFolderPath

    out-logfile -string "Azure IP information is included in the query."
    
    $azurePublicData = get-azureData -dataLocation $azurePublicCloudXML
    $azureGovernmentData = get-AzureData -dataLocation $azureGovernmentCloudXML
    export-JSONInformation -dataToExport $azurePublicData -exportPath $allIPAzurePublic
    export-JSONInformation -dataToExport $azureGovernmentData -exportPath $allIPAzureGovernment
}

if ($IPAddressToTest -ne $noIPSpecified)
{
    out-logfile -string "Begin testing IP spaces for presence of the specified IP address."

    test-IPSpace -dataToTest $allIPInformationWorldWide -IPAddress $IPAddressToTest -regionString $worldWideRegionString -portToTest $portToTest
    test-IPSpace -dataToTest $allIPInformationChina -IPAddress $IPAddressToTest -regionString $chinaRegionString -portToTest $portToTest
    test-IPSpace -dataToTest $allIPInfomrationUSGovGCCHigh -IPAddress $IPAddressToTest -regionString $gccHighRegionString -portToTest $portToTest
    test-IPSpace -dataToTest $allIPInformationUSGovDOD -IPAddress $IPAddressToTest -regionString $dodRegionString -portToTest $portToTest
    
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

    if ($includeAzureSearch -eq $TRUE)
    {
        out-logfile -string "Begin searching for IP address in azure space."

        test-AzureIPSpace -IPAddress $ipAddressToTest -dataToTest $azurePublicData
        test-AzureIPSpace -IPAddress $ipAddressToTest -dataToTest $azureGovernmentData
    }

    out-logfile -string "If specified gather the IP location for the IP specified."

    if (($allowQueryIPLocationInformationFromThirdParty -eq $TRUE) -and ($IPAddressToTest -ne $noIPSpecified))
    {
        $global:ipLocation = get-IPLocationInformation -ipAddress $ipAddressToTest

        out-logfile -string $global:ipLocation

        <#

        if ($global:ipLocation -ne "Failed")
        {
            out-logfile -string "Converting IP location JSON."
            $global:ipLocation = get-jsonData -data $global:ipLocation
        }

        #>
    }

    if ($global:outputArray.count -gt 0)
    {        
        foreach ($member in $global:outputArray)
        {
            out-logfile -string $member
            $member
        }

        export-JSONInformation -dataToExport $global:outputArray -exportPath $outputXMLFile
        
        if ($global:outputChangeArray.count -gt 0)
        {
            write-host "IP entries present in the changes file:"
    
            foreach ($member in $global:outputChangeArray)
            {
                out-logfile -string $member
                $member
            }

            export-JSONInformation -dataToExport $global:outputChangeArray -exportPath $outputChangeXMLFile        
        }
    }

    if ($global:outputRemoveArray.count -gt 0)
    {
        foreach ($member in $global:outputRemoveArray)
        {
            out-logfile -string $member
            $member
        }

        export-JSONInformation -dataToExport $global:outputRemoveArray -exportPath $outputRemoveXMLFile
    }

    if ($global:outputAzureArray.count -gt 0)
    {
        foreach ($member in $global:outputAzureArray)
        {
            out-logfile -string $member
            $member
        }

        export-JSONInformation -dataToExport $global:outputAzureArray -exportPath $outputAzureXMLFile
    }

    generate-HTMLData -IPorURLToTest $IPAddressToTest
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
        foreach ($entry in $global:outputArray)
        {
            out-logfile -string $entry
            $entry
        }

        export-JSONInformation -dataToExport $global:outputArray -exportPath $outputXMLFile

        if ($global:outputChangeArray.count -gt 0)
        {    
            foreach ($entry in $global:outputChangeArray)
            {
                out-logfile -string $entry
                $entry
            }

            export-JSONInformation -dataToExport $global:outputChangeArray -exportPath $outputChangeXMLFile
        }
    }

    if ($global:outputRemoveArray.count -gt 0)
    {
        foreach ($entry in $global:outputRemoveArray)
        {
            out-logfile -string $entry
            $entry
        }

        export-JSONInformation -dataToExport $global:outputRemoveArray -exportPath $outputRemoveXMLFile
    }

    generate-HTMLData -IPorURLToTest $URLToTest
}