[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Set-StrictMode -Version Latest

[string]$BaseUri = "https://api.trello.com/1"
[string]$PSTrelloXml = "$env:LOCALAPPDATA\PSTrelloXml.xml"
[PSCustomObject]$Global:TrelloAccount = $null

function PSTrello 
{

}

###############
### ACCOUNTS
###############

Function Add-TrelloAccount
{
    param 
    (
        [Parameter(Mandatory)]
        $Name,
        [Parameter(Mandatory)]
        $AppKey,
        [Parameter(Mandatory)]
        $Token
    )

    #build new account object
    $Account =  [PSCustomObject]@{
        Name=$Name
        AppKey=$AppKey
        Token=$Token
        String="key=$AppKey&token=$Token"} 


    $ImportedAccounts = @()

    if ((Get-TrelloAccount) -ne $null)
    {
        $ImportedAccounts += (Import-Clixml -Path $PSTrelloXml)
    }

    $ImportedAccounts += $Account
    $ImportedAccounts | Export-Clixml -Path $PSTrelloXml
}

Function Get-TrelloAccount
{
    [cmdletbinding(DefaultParameterSetName="All")]
    param (
        
        [Parameter(ParameterSetName='Name')]
        [string]$Name,

        [Parameter(ParameterSetName='Active')]
        [switch]$Active,

        [Parameter(ParameterSetName='All')]
        [switch]$All
    )

    if ([System.Io.file]::Exists($PSTrelloXml))
    {
        try {
            if ($PSCmdlet.ParameterSetName -eq "All") 
            { return Import-Clixml -Path $PSTrelloXml }
            if ($PSCmdlet.ParameterSetName -eq "Name")
            { return Import-Clixml -Path $PSTrelloXml | ? {$_.Name -eq $Name} | select -First 1 }
            if ($PSCmdlet.ParameterSetName -eq "Active")
            {if ($TrelloAccount -ne $null) {return $TrelloAccount }}
             else {Write-Error -Message "Set Trello account first with Set-TrelloAccount"}
        }
        catch { 
            return $null 
        }
    }

    return $null
}

Function Set-TrelloAccount
{
    param (
        [Parameter(ParameterSetName='Name',Mandatory=$true)]
        [string]$Name
    )

    if ((Get-TrelloAccount -Name $Name) -ne $null)
    {
        $Global:TrelloAccount = Get-TrelloAccount -Name $Name
        Write-Host("Trello account set as $Name")
    }
    else
    {
        Write-Host("$Name was not found as an available account. Use Add-TrelloAccount to add your account") -ForegroundColor Red
    }
}


###############
### BOARDS
###############

Function Get-TrelloBoard
{
    [cmdletbinding(DefaultParameterSetName="All")]
    param (
        [Parameter(ParameterSetName='Name',Mandatory=$true)]
        [string]$Name,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [int]$Index = 0,

        [Parameter(ParameterSetName='All')]
        [switch]$All
    )

    if ($TrelloAccount -eq $null)
    { return write-error -Message "Set Trello account first with Set-TrelloAccount"}

    try
    {
        if ($PSCmdlet.ParameterSetName -eq "All")
        { return (invoke-restmethod -uri "$BaseUri/members/me/boards?$($TrelloAccount.String)") }

        return ((invoke-restmethod -uri "$BaseUri/members/me/boards?$($TrelloAccount.String)") | ? {$_.name -eq "$Name"} | select -Index $Index)
    }
    catch {Write-Error $_.Exception.Message }

}

###############
### LISTS
###############

Function Get-TrelloList
{
    [cmdletbinding(DefaultParameterSetName="All")]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$Board,

        [Parameter(ParameterSetName='Name',Mandatory=$true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Name,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [int]$Index = 0,

        [Parameter(ParameterSetName='All')]
        [switch]$All
    )

    if ($PSCmdlet.ParameterSetName -eq "All")
    { return (invoke-restmethod -uri "$BaseUri/boards/$($Board.id)/lists?$($TrelloAccount.String)") }

    return ((invoke-restmethod -uri "$BaseUri/boards/$($Board.id)/lists?$($TrelloAccount.String)") | ? {$_.name -eq "$Name"} | select -Index $Index)
}

Function Remove-TrelloList
{
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$List
    )

    $body = @{
        closed=$true
    } | ConvertTo-Json

    return (invoke-restmethod -uri "$BaseUri/lists/$($list.id)?$($TrelloAccount.String)" -Body $body  -Method Put -ContentType 'application/json')

}

Function New-TrelloList {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [object]$Board,
        [Parameter(Mandatory)]
        [object]$Name,
        [ValidateSet("top","bottom")]
        [object]$Position
    )

    $body = @{
        name=$Name
        pos=$Position
    } | ConvertTo-Json

    return (invoke-restmethod -uri "$BaseUri/boards/$($board.id)/lists?$($TrelloAccount.String)" -Body $body  -Method Post -ContentType 'application/json')

}

###############
### CARD
###############

Function Get-TrelloCard
{
    [cmdletbinding(DefaultParameterSetName="All")]
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$List,

        [Parameter(ParameterSetName='Name',Mandatory=$true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Name,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [int]$Index = 0,

        [Parameter(ParameterSetName='All')]
        [switch]$All
    )

    #if ($Name -eq $null) {return $null}

    try {
        if ($PSCmdlet.ParameterSetName -eq "All") { return (invoke-restmethod -uri "$BaseUri/lists/$($List.id)/cards?$($TrelloAccount.String)") }
        else {return ((invoke-restmethod -uri "$BaseUri/lists/$($List.id)/cards?$($TrelloAccount.String)") | ? {$_.name -eq "$Name"} | select -Index $Index)}
    } catch {Write-Error $_.Exception.Message }
}

Function Update-TrelloCard 
{
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$Card,
        [Parameter(Mandatory=$true)]
        $JsonObject
    )

    return (invoke-restmethod -uri "$BaseUri/cards/$($Card.id)?$($TrelloAccount.String)" -Method Put -Body $JsonObject -ContentType 'application/json')
}

Function New-TrelloCard
{
    param (
        [Parameter(ParameterSetName='JsonObject',Mandatory=$true)]
        $JsonObject,

        [Parameter(ParameterSetName='Name',Mandatory=$true,ValueFromPipeline)]
        [object]$List,
        [Parameter(ParameterSetName='Name',Mandatory=$true)]
        [string]$Name,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [string]$Description,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [string]$Labels,
        [Parameter(ParameterSetName='Name',Mandatory=$false)]
        [string]$Position
    )

    if ($PSCmdlet.ParameterSetName -eq "JsonObject") {
        return (invoke-restmethod -uri "$BaseUri/cards?$($TrelloAccount.String)" -Method Post -Body $JsonObject -ContentType 'application/json')
    }

    if ($PSCmdlet.ParameterSetName -eq "Name") {
        $JsonObject = @{
            name=$Name
            desc=$Description
            pos=$Position
            idList=$List.id
            idLabels=$Labels
        } | ConvertTo-Json
        write-host($JsonObject)
        return (invoke-restmethod -uri "$BaseUri/cards?$($TrelloAccount.String)" -Method Post -Body $JsonObject -ContentType 'application/json')
    }

}

Function Remove-TrelloCard
{
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$Card
    )

    return (invoke-restmethod -uri "$BaseUri/cards/$($Card.id)?$($TrelloAccount.String)" -Method Delete -ContentType 'application/json')

}

function Add-TrelloCardAttachment 
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory, ValueFromPipeline)]
		[ValidateNotNullOrEmpty()]
		[object]$Card,

		[ValidateNotNullOrEmpty()]
	    [string]$Name,

		[Parameter(ParameterSetName='File',Mandatory)]
		[ValidateNotNullOrEmpty()]
		[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
		[string]$FilePath,
        
        [Parameter(ParameterSetName='URL',Mandatory)]
		[string]$URL

	)
	begin {
		$ErrorActionPreference = 'Stop'
	}
	process {
            try {
                if ($PSCmdlet.ParameterSetName -eq "File")
                {
                    $URI = "$BaseUri/cards/$($card.id)/attachments?name=$Name&$($TrelloAccount.String)"
                    return (curl.exe -s -F file=@$filepath $URI) | ConvertFrom-Json
                }
                if ($PSCmdlet.ParameterSetName -eq "URL")
                {
                    $URL = [System.Web.HttpUtility]::UrlEncode($URL)
                    $URI = "$BaseUri/cards/$($card.id)/attachments?url=$URL&name=$Name&$($TrelloAccount.String)"
                    return (invoke-restmethod -Uri $URI -Method Post)
                }
            }
            catch {
                Write-Error $_.Exception.Message
            }
	}
}


###############
### LABELS
###############

function Get-TrelloLabel
{
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$Board,
        [Parameter(Mandatory=$false)]
        [object]$Fields="all",
        [Parameter(Mandatory=$false)]
        [int]$Limit=100


    )

    try { return (invoke-restmethod -uri "$BaseUri/boards/$($Board.id)/labels?fields=$Fields&limit=$Limit&$($TrelloAccount.String)") }
    catch { return null }
}

function New-TrelloLabel
{
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline)]
        [object]$Board,
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [Parameter(Mandatory=$true)]
        [string]$Color
    )

    try { return (invoke-restmethod -uri "$BaseUri/boards/$($Board.id)/labels?name=$name&color=$color&$($TrelloAccount.String)" -Method Post) }
    catch { return null }
}

function Remove-TrelloLabel
{
        param (
            [Parameter(Mandatory)]
            [ValidateSet("Board","Card")]
            [string[]]$RemoveFrom,
            [Parameter(Mandatory,ValueFromPipeline)]
            [object]$Object
        )

        if ($ObjectType -eq "Board")
        {
            try { return (invoke-restmethod -uri "$BaseUri/labels/$($Label.id)?$($TrelloAccount.String)" -Method Delete) }
            catch { return null }
        }

        if ($ObjectType -eq "Card")
        {
            foreach ($label in $card.labels)
            {
                $label

            }
        }
}
