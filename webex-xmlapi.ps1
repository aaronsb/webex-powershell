







#region utilities
function Invoke-WebExAPICall
{param($XML,[switch]$debug)
        $WebExSiteName = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext.SiteName
        # WebEx XML API URL 
        $URL = "https://$WebExSiteName.webex.com/WBXService/XMLService"

        # Send WebRequest and save to URLResponse
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $URLResponse = Invoke-WebRequest -Uri $URL -Method Post -ContentType 'text/xml' -TimeoutSec 120 -Body $XML
            Write-Verbose "Webex: Getting information for user $user_WebexID"
        } catch {
            Write-Error "Webex: Unable to send XML request"
        }
        
        $xmlData = ([xml]($URLresponse.get_content())).message.body.bodyContent
        $xmlMeta = ([xml]($URLresponse.get_content())).message.header.response
        
        if ($xmlMeta.result -ne "SUCCESS")
        {
            if ($xmlMeta)
            {
                Write-Error -Category ObjectNotFound -Message ($xmlMeta.result + ": " + $xmlMeta.reason)
                return [PSCustomObject]@{"xmlData"=$xmlData;"xmlMeta"=$xmlMeta;"URLResponse"=$URLResponse}
            }
            else
            {
                Throw "No metadata returned from xml service endpoint. No other error data is available."
            }
            if ($debug)
            {
                return [PSCustomObject]@{"xmlData"=$xmlData;"xmlMeta"=$xmlMeta;"URLResponse"=$URLResponse}
            }
            
        }
        else
        {
            Write-Verbose "WebEx Service Success"
            return [PSCustomObject]@{"xmlData"=$xmlData;"xmlMeta"=$xmlMeta;"URLResponse"=$URLResponse}
        }    
}



function Get-WebExUserXML
{param([Parameter(mandatory=$true)]$UserID)

    $SecurityContextSource = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext
    $XML = [xml](Get-Content .\xml\get-webexuser.xml)
    $securityContext = $XML.message.header.securityContext
    $bodyContent = $xml.message.body.bodyContent
    
    $bodyContent.webExId = $UserID
    $securityContext.webExID = ($SecurityContextSource.ServiceAccountName)
    $securityContext.password = ($SecurityContextSource.ServiceAccountPassword)
    $securityContext.siteID = ($SecurityContextSource.SiteID)
    $securityContext.partnerID = ($SecurityContextSource.PartnerID)
    return $xml
}

function Remove-WebExUserXML
{param([Parameter(mandatory=$true)]$UserID)

    $SecurityContextSource = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext
    $XML = [xml](Get-Content .\xml\remove-webexuser.xml)
    $securityContext = $XML.message.header.securityContext
    $bodyContent = $xml.message.body.bodyContent
    
    $bodyContent.webExId = $UserID
    $securityContext.webExID = ($SecurityContextSource.ServiceAccountName)
    $securityContext.password = ($SecurityContextSource.ServiceAccountPassword)
    $securityContext.siteID = ($SecurityContextSource.SiteID)
    $securityContext.partnerID = ($SecurityContextSource.PartnerID)
    return $xml
}


function New-WebExUserXML
{param($ADObject,$SAMAccountName)
    $Properties = @("GivenName","Surname","Title","Description","Company","SAMAccountName","StreetAddress","City","State","PostalCode","EmailAddress")
    if ($SAMAccountName)
    {
        $ADUser = Get-ADUser $SAMAccountName -Properties $Properties
    }

    if ($ADObject)
    {
        $ADUser = $ADObject
        #so some sort of error checking here. no time right now.
        #compare-object $ADObject.PropertyNames $Properties
    }

    if ((!$SAMAccountName) -and (!$ADObject))
    {
        throw "Need ADObject or SAMAccountName to configure XML template."
    }
    $ErrorActionPreference = "SilentlyContinue"
    

    $SecurityContextSource = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext
    $XML = [xml](Get-Content .\xml\new-webexuser.xml)
    $securityContext = $XML.message.header.securityContext
    $bodyContent = $xml.message.body.bodyContent

    #todo implement trap so errors don't silently fail on unknown object entities
    $bodyContent.firstName = $ADUser.GivenName
    $bodyContent.lastName = $ADUser.Surname
    $bodyContent.title = $ADUser.Title
    $bodyContent.description = $ADUser.Description
    $bodyContent.company = $ADUser.Company
    $bodyContent.webExId = $ADUser.SAMAccountName
    $bodyContent.address.address1 = ($ADUser.StreetAddress -split [Environment]::NewLine)[0].ToString()
    $bodyContent.address.address2 = ($ADUser.StreetAddress -split [Environment]::NewLine)[1].ToString()
    $bodyContent.address.city = $ADUser.City
    $bodyContent.address.state = $ADUser.State
    $bodyContent.address.zipCode = $ADUser.PostalCode
    $bodyContent.email = $ADUser.EmailAddress
    $bodyContent.password = $bodyContent.password
    $bodyContent.personalUrl = $ADUser.SAMAccountName
    $bodyContent.personalMeetingRoom.title = ($ADUser.$GivenName + " " + $ADUser.$Surname + "'s Personal Room")
    $bodyContent.personalMeetingRoom.personalMeetingRoomURL = ("https://"+ $bodyContent.company + ".webex.com/meet/" + $ADUser.SAMAccountName)
    $bodyContent.personalMeetingRoom.sipURL = ($ADUser.SAMAccountName + "@" + $bodyContent.company + ".webex.com")
    $bodyContent.personalMeetingRoom.hostPIN = (Get-Random -Minimum 0 -Maximum 9999).ToString('0000')
    $securityContext.webExID = ($SecurityContextSource.ServiceAccountName).ToString()
    $securityContext.password = ($SecurityContextSource.ServiceAccountPassword).ToString()
    $securityContext.siteID = ($SecurityContextSource.SiteID).ToString()
    $securityContext.partnerID = ($SecurityContextSource.PartnerID).ToString()
    
    return $xml
}

function Get-WebExServiceXML
{
    $SecurityContextSource = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext
    $XML = [xml](Get-Content .\xml\get-webexservice.xml)
    $securityContext = $XML.message.header.securityContext
    
    $securityContext.webExID = ($SecurityContextSource.ServiceAccountName).ToString()
    $securityContext.password = ($SecurityContextSource.ServiceAccountPassword).ToString()
    $securityContext.siteID = ($SecurityContextSource.SiteID).ToString()
    $securityContext.partnerID = ($SecurityContextSource.PartnerID).ToString()
    return $xml
}

function Get-WebExUserSummaryXML
{param($startFrom=1,$maximumNum=10,$listMethod="AND",$orderBy="UID",$orderAD="ASC",$dateStart,$dateEnd)
    
    $dateNow = Get-Date
    
    if (!$dateStart)
    {
        $dateStart = $dateNow.AddYears(-10)
    }
    
    if (!$dateEnd)
    {
      $dateEnd = $dateNow  
    }
    $SecurityContextSource = ([xml](Get-Content .\xml\webex-auth.xml)).SecurityContext
    $xml = [xml](Get-Content .\xml\get-webexusersummary.xml)
    
    $securityContext = $XML.message.header.securityContext
    $listControl = $xml.message.body.bodyContent.listControl
    $order = $xml.message.body.bodyContent.order
    $datascope = $xml.message.body.bodyContent.datascope

    if ([datetime]$dateStart -gt [datetime]$dateEnd)
    {
        write-error "Start and End dates are invalid. You can't start before you finish."
        break
    }

    $dateStart = ([datetime]$dateStart).ToString("MM/dd/yyyy HH:mm:ss")
    $dateEnd = ([datetime]$dateEnd).ToString("MM/dd/yyyy HH:mm:ss")
    
    $securityContext.webExID = ($SecurityContextSource.ServiceAccountName)
    $securityContext.password = ($SecurityContextSource.ServiceAccountPassword)
    $securityContext.siteName = ($SecurityContextSource.SiteName)
        
    $listControl.startFrom = $startFrom.toString()
    $listControl.maximumNum = $maximumNum.toString()
    $listControl.listMethod = $listMethod.toString()
    
    $order.orderBy = $orderBy.toString()
    $order.orderAD = $orderAd.toString()

    $datascope.regDateStart = $dateStart.toString()
    $datascope.regDateEnd = $dateEnd.toString()
    return $xml
}

function WriteXmlToScreen ([xml]$xml)
{
    $StringWriter = New-Object System.IO.StringWriter;
    $XmlWriter = New-Object System.Xml.XmlTextWriter $StringWriter;
    $XmlWriter.Formatting = "indented";
    $xml.WriteTo($XmlWriter);
    $XmlWriter.Flush();
    $StringWriter.Flush();
    Write-Output $StringWriter.ToString();
}


function ConvertFrom-Xml {
    <#
    .SYNOPSIS
        Converts XML object to PSObject representation for further ConvertTo-Json transformation
    .EXAMPLE
        # JSON->XML
        $xml = ConvertTo-Xml (get-content 1.json | ConvertFrom-Json) -Depth 4 -NoTypeInformation -as String
    .EXAMPLE
        # XML->JSON
        ConvertFrom-Xml ([xml]($xml)).Objects.Object | ConvertTo-Json
    #>
        param([System.Xml.XmlElement]$Object)
    
        if (($Object -ne $null) -and ($Object.Property -ne $null)) {
            $PSObject = New-Object PSObject
    
            foreach ($Property in @($Object.Property)) {
                if ($Property.Property.Name -like 'Property') {
                    $PSObject | Add-Member NoteProperty $Property.Name ($Property.Property | % {ConvertFrom-Xml $_})
                } else {
                    if ($Property.'#text' -ne $null) {
                        $PSObject | Add-Member NoteProperty $Property.Name $Property.'#text'
                    } else {
                        if ($Property.Name -ne $null) {
                            $PSObject | Add-Member NoteProperty $Property.Name (ConvertFrom-Xml $Property)
                        }
                    }
                } 
            }   
            $PSObject
        }
    }




function Get-Excuse
{
    $excusePath = ("$home\excuses.txt")
    if ((Test-Path $excusePath) -eq $false)
    {
        Write-Verbose "Researching Excuses"
        (Invoke-WebRequest -uri "http://pages.cs.wisc.edu/~ballard/bofh/excuses").content > $excusePath
    }
    $excuses = (Get-Content $excusePath).split(([environment]::NewLine))
    $ID = (Get-Random $excuses.count)
    $message = $excuses[$ID]
    Write-Error $message -ErrorId $ID
}

## send-eventmessage
## aaron bockelie
# a wrapper for setting up an event. Simplifies creating events, since it creates the event source automagically if it doesn't exist.


