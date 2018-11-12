

#load the xmlapi

. .\webex-xmlapi.ps1

function Get-WebExADAudit
{param([switch]$gridview)
    function Priv_GetDiff
    {
        $AD = [xml](gc .\xml\ADConfig.xml)
        $ldapFilter = ($ad.Config.searchQuery -replace "`n","" -replace "`r","" -replace "`t","" -replace " ","")
        $searchBase = ($ad.Config.searchBase)        
        $ADUsers = Get-ADUser -LDAPFilter $ldapFilter -SearchBase $searchBase -Properties GivenName,Surname,Title,Description,Company,SAMAccountName,StreetAddress,City,State,PostalCode,EmailAddress
        $WebExUsers = Get-WebExUserSummary
        $DiffObject = Compare-Object $adusers.samaccountname ($WebExUsers.xmlData.user | %{$_.webexid}) -IncludeEqual
        ForEach ($Comparison in $DiffObject)
        {
            $WebExUser = ($WebExUsers.xmldata.user | ?{$_.webExId -eq $Comparison.InputObject.ToString()})
            switch ($Comparison.SideIndicator) {
                "==" {[pscustomobject]@{"AccountName" = $Comparison.InputObject;"ProvisionStatus" = "Syncronized";"Status" = $WebExUser.active}}
                "<=" {[pscustomobject]@{"AccountName"= $Comparison.InputObject;"ProvisionStatus" = "NotSyncronized";"Status" = "null"}}
                "=>" {[pscustomobject]@{"AccountName" = $Comparison.InputObject;"ProvisionStatus" = "UnknownSyncronized";"Status" = $WebExUser.active}}
                default {write-error "Unknown comparison"; break}
            }
            $WebExUser = $null
        }
    }

    $result = Priv_GetDiff
    if ($gridview)
    {$result | Out-Gridview}
    else
    {
        return $result
    }
}



function Add-WebExUser
{param([Parameter(mandatory=$true)]$UserID)
    
    $XML = New-WebExUserXML -SAMAccountName $UserID
    Write-Host ("Creating new user " + $xml.message.body.bodyContent.firstName + " " + $xml.message.body.bodyContent.lastName)
    Write-Host ($xml.message.body.bodyContent.title)
    Write-Host ($xml.message.body.bodyContent.email)
    Write-Host " "
    Write-Host "Press `"Enter`" to continue or `"Ctrl-C`" to cancel"
    do
    {
        $key = [Console]::ReadKey("noecho")
    }
    while($key.Key -ne "Enter")
    Write-Host "Creating User"
    $result = Invoke-WebExAPICall $XML
    $result.xmlMeta
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

function Get-WebexUser
{param([Parameter(mandatory=$true)]$UserID,[switch]$OutputXML)
    if (!$OutputXML)
    {
        Invoke-WebExAPICall (Get-WebExUserXML -UserID $UserID)
    }
    else 
    {
        WriteXmlToScreen (Invoke-WebExAPICall (Get-WebExUserXML -UserID $UserID)).URLResponse
    }
    
}

function Remove-WebexUser
{param([Parameter(mandatory=$true)]$UserID)
    
    Invoke-WebExAPICall (Remove-WebExUserXML -UserID $UserID)
}




function Get-WebExService
{param([switch]$OutputXML)
    if (!$OutputXML)
    {
        Invoke-WebExAPICall (Get-WebExServiceXML)
    }
    else
    {
        WriteXmlToScreen (Invoke-WebExAPICall (Get-WebExServiceXML))
    }
}

function Get-WebExUserSummary
{
    Invoke-WebExAPICall (Get-WebExUserSummaryXML -dateStart "1/1/2010" -dateEnd (Get-Date) -maximumNum 200)
}

