#Requires -Version 5.1
#Requires -Modules ImportExcel, Toolbox.IPAM
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, Toolbox.FileAndFolder

<#
.SYNOPSIS
    Add fixed IP addresses to IPAM.

.DESCRIPTION
    The IPAM fixed IP addresses table will be populated based on an Excel file. Each row must
    have the two mandatory fields 'name', ipv4addr' and 'match_client'.

    ADD A NEW ADDRESS RESERVATION
    ------------------------------
    When a new address reservation needs to be made in IPAM we first check the following:
        - Is the IP/hostname online?
        - Is the IP/hostname known in DNS?
    if one of these cases is true an informative error message is written to the row of that specific
    address and the address reservation is not made in IPAM. This will allow the user to verify before
    effectively adding the new address reservation:

    Ex. OverWrite : NULL or FALSE
        Status    : Error
        Action    :
        Error     : ipv4addr '10.10.10.1' is online, please verify it is not in use.

    If these tests can be ignored the property 'OverWrite' can be set to TRUE and the new address
    reservation will be added without testing for the cases above:

    Ex. OverWrite : TRUE
        Status    : Added
        Action    : Added new reservations
        Error     :

    When all tests passed and there is no error the address will simply be added:

    Ex. OverWrite : NULL or FALSE
        Status    : Added
        Action    : Added new reservations
        Error     :

    UPDATE AN EXISTING ADDRESS RESERVATION
    --------------------------------------
    By checking if the requested ipv4addr, mac, name and/or ddns_hostname address is known by any
    fixed address in IPAM, it can be determined if the address reservation has already been made
    for this specific row.

    When the reservation is already made in IPAM its properties will simply be *cross checked with
    the requested properties in the row. In case any of the properties is not matched it will be
    reported and no changes will be made:

    Ex. OverWrite       : NULL or FALSE
        Status          : Incorrect
        Action          :
        Error           :
        IncorrectFields : ddns_hostname, comment

    if these fields need to be corrected, the property 'OverWrite' can be set to TRUE and the incorrect
    fields in IPAM will be updated:

    Ex. OverWrite       : TRUE
        Status          : Updated
        Action          : Updated incorrect fields
        IncorrectFields : ddns_hostname, comment

    * properties in the row with value NULL will be ignored and not cross checked with IPAM.


    The field 'match_client' is mandatory because it decides on how to assign a fixed address to a client.
    Possible values for 'match_client' are:
    - MAC_ADDRESS : The fixed IP address is leased to the matching MAC address.
    - RESERVED    : The fixed IP address is reserved for later use with a MAC address that only has zeros.
    - CLIENT_ID   : The fixed IP address is leased to the matching DHCP client identifier.Note that
                    the “dhcp_client_identifier” field must be set in this case.
    - CIRCUIT_ID  : The fixed IP address is leased to the DHCP client with a matching circuit ID. Note that
                    the “agent_circuit_id” field must be set in this case.
    - REMOTE_ID   : The fixed IP address is leased to the DHCP client with a matching remote ID. Note that
                    the “agent_remote_id” field must be set in this case.

.LINK
    https://ipam/wapidoc
    https://ipam/wapidoc/objects/fixedaddress.html?highlight=match_client#fixedaddress.match_client
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Parameter(Mandatory)]
    [String]$Environment,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\IPAM\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Function Compare-ConfigHC {
        [OutputType([HashTable])]
        Param (
            [Parameter(Mandatory)]
            [PSCustomObject]$ReferenceObject,
            [Parameter(Mandatory)]
            [PSCustomObject]$DifferenceObject,
            [Parameter(Mandatory)]
            [String[]]$Property
        )

        $Result = @{ }

        @($ReferenceObject.PSObject.Properties).where( {
                ($Property.Contains($_.Name)) -and
                ($_.Value -ne $DifferenceObject.($_.Name)) -and
                ($null -ne $_.Value)
            }).Foreach( {
                $Result[$_.Name] = $_.Value
            })

        if ($Result.Count -ne 0) {
            $Result
        }
    }

    Try {
        $Error.Clear()
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Test ImportFile present
        if (-not (Test-Path -Path $ImportFile -PathType Leaf)) {
            throw "Import file '$ImportFile' not found"
        }
        #endregion

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import file properties
        $FixedAddressWorksheetProperties = @{
            comment                = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            ddns_hostname          = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            disable                = @{
                Type        = 'Boolean'
                Mandatory   = $false
                ApiProperty = $true
            }
            device_type            = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            device_vendor          = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            device_location        = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            enable_ddns            = @{
                Type        = 'Boolean'
                Mandatory   = $false
                ApiProperty = $true
            }
            ipv4addr               = @{
                Type        = 'String'
                Mandatory   = $true
                ApiProperty = $true
            }
            mac                    = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            match_client           = @{
                Type        = 'String'
                Mandatory   = $true
                ApiProperty = $true
            }
            name                   = @{
                Type        = 'String'
                Mandatory   = $true
                ApiProperty = $true
            }
            agent_circuit_id       = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            agent_remote_id        = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            dhcp_client_identifier = @{
                Type        = 'String'
                Mandatory   = $false
                ApiProperty = $true
            }
            # Non API properties:
            overwrite              = @{
                Type        = 'Boolean'
                Mandatory   = $false
                ApiProperty = $false
            }
        }

        $RemoveWorksheetProperties = @{
            ddns_hostname = @{
                Type        = 'String'
                ApiProperty = $true
            }
            ipv4addr      = @{
                Type        = 'String'
                ApiProperty = $true
            }
            mac           = @{
                Type        = 'String'
                ApiProperty = $true
            }
            name          = @{
                Type        = 'String'
                ApiProperty = $true
            }
        }
        #endregion

        #region Get API properties
        if (-not ($ApiPropertyList = ($FixedAddressWorksheetProperties.GetEnumerator().where( { $_.Value.ApiProperty })).Name)) {
            throw 'The API property list cannot be empty.'
        }
        #endregion

        $RemoveWorksheetErrors = $FixedAddressWorksheetErrors = $ConflictingWorksheetErrors = @()

        $FixedAddressWorksheet = @(Import-Excel -Path $ImportFile -WorksheetName 'FixedAddress' |
            Remove-ImportExcelHeaderProblemOnEmptySheetHC |
            Select-Object -Property @{N = 'Status'; E = { $null } }, *,
            @{N = 'Action'; E = { , @() } },
            @{N = 'Error'; E = { $null } },
            @{N = 'IncorrectFields'; E = { $null } } -ExcludeProperty Status, Action, Error, IncorrectFields )

        $RemoveWorksheet = @(Import-Excel -Path $ImportFile -WorksheetName 'Remove' |
            Remove-ImportExcelHeaderProblemOnEmptySheetHC |
            Select-Object -Property @{N = 'Status'; E = { $null } }, *,
            @{N = 'Action'; E = { , @() } },
            @{N = 'Error'; E = { $null } } -ExcludeProperty Status, Action, Error )

        #region Test Worksheet Remove
        foreach ($R in $RemoveWorksheet) {
            #region Remove leading and trailing spaces
            @($RemoveWorksheetProperties.GetEnumerator().where( {
                        ($_.Value.Type -eq 'String') -and
                        ($R.PSObject.Properties.Name -contains $_.Name  )
                    })).ForEach( {
                    $R.($_.Name) = if (
                        ($R.($_.Name)) -and ($tmp = $R.($_.Name).Trim())) {
                        $tmp
                    }
                    else {
                        $null
                    }
                })
            #endregion

            #region Test valid ipv4addr
            if (
                ($R.ipv4addr) -and
                (-not (($R.ipv4addr -match '^(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])$'))
                )
            ) {
                $RemoveWorksheetErrors += "ipv4addr '$($R.ipv4addr)' is not a valid IP address."
            }
            #endregion

            #region ddns_hostname cannot contain spaces or dots
            if ($R.ddns_hostname -match '\s|\.') {
                $RemoveWorksheetErrors += "The field ddns_hostname '$($F.ddns_hostname)' cannot contain spaces or dots. Fully qualified domain names are not needed."
            }
            #endregion

            #region Test valid mac
            if ($R.mac) {
                if ($R.mac -notmatch '^((([a-zA-z0-9]{2}[-:]){5}([a-zA-z0-9]{2}))$|^(([a-zA-z0-9]{2}:){5}([a-zA-z0-9]{2})))$') {
                    $RemoveWorksheetErrors += "mac '$($R.mac)' is not a valid MAC address."
                    Continue
                }

                $R.mac = $R.mac.ToLower()

                $R.mac = $R.mac.Replace('-', ':')

                if ($R.mac -eq '00:00:00:00:00:00') {
                    $RemoveWorksheetErrors += "mac '$($R.mac)' is not allowed."
                    Continue
                }
            }
            #endregion
        }

        #region Test duplicate name
        $RemoveWorksheet.where( { $_.name }) |
        Group-Object name | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $RemoveWorksheetErrors += "The name '$($_.Name)' is used multiple times and needs to be unique."
        }
        #endregion

        #region Test duplicate ddns_hostname
        $RemoveWorksheet.where( { $_.ddns_hostname }) |
        Group-Object ddns_hostname | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $RemoveWorksheetErrors += "The ddns_hostname '$($_.Name)' is used multiple times and needs to be unique."
        }
        #endregion

        #region Test duplicate mac
        $RemoveWorksheet.where( { ($_.mac) -and ($_.mac -ne '00:00:00:00:00:00') }) |
        Group-Object mac | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $RemoveWorksheetErrors += "The mac '$($_.Name)' is used multiple times and needs to be unique."
        }
        #endregion

        #region Test duplicate ipv4addr
        $RemoveWorksheet.where( { ($_.ipv4addr) -and ($_.ipv4addr -notmatch '^func:') }) |
        Group-Object ipv4addr | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $RemoveWorksheetErrors += "The ipv4addr '$($_.Name)' is used multiple times and needs to be unique."
        }
        #endregion
        #endregion

        #region Test Worksheet FixedAddress
        foreach ($F in $FixedAddressWorksheet) {
            #region Remove leading and trailing spaces
            @($FixedAddressWorksheetProperties.GetEnumerator().where( {
                        ($_.Value.Type -eq 'String') -and
                        ($F.PSObject.Properties.Name -contains $_.Name  )
                    })).ForEach( {
                    $F.($_.Name) = if (
                        ($F.($_.Name)) -and ($tmp = $F.($_.Name).Trim())) {
                        $tmp
                    }
                    else {
                        $null
                    }
                })
            #endregion

            #region Test mandatory properties
            @($FixedAddressWorksheetProperties.GetEnumerator().where( { $_.Value.Mandatory }).Name).Where( { -not ($F.$_) }).ForEach( {
                    $FixedAddressWorksheetErrors += "The mandatory property '$_' is missing for IPv4addr '$($F.ipv4addr)'."
                    Continue
                })
            #endregion

            #region Test valid ipv4addr
            if (
                -not (
                    ($F.ipv4addr -match '^(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])$') -or
                    ($F.ipv4addr -match '^func:')
                )
            ) {
                $FixedAddressWorksheetErrors += "ipv4addr '$($F.ipv4addr)' is not a valid IP address."
            }
            #endregion

            #region Test boolean properties
            @($FixedAddressWorksheetProperties.GetEnumerator().where( { $_.Value.Type -eq 'Boolean' }).Name).where( {
                    ($F.$_) -and (-not ($F.$_ -is [Boolean]))
                }).ForEach( {
                    $FixedAddressWorksheetErrors += "The boolean property '$_' with value '$($F.$_)' for ipv4addr '$($F.ipv4addr)' is not valid. Only TRUE, FALSE or NULL are supported"
                })
            #endregion

            #region Test ddns_hostname

            #region ddns_hostname is mandatory when enable_ddns is true
            if ($F.enable_ddns -and (-not $F.ddns_hostname)) {
                $FixedAddressWorksheetErrors += "The field 'ddns_hostname' is mandatory when 'enable_ddns' is set to true for ipv4addr '$($F.ipv4addr)'."
            }
            #endregion

            #region ddns_hostname cannot contain spaces or dots
            if ($F.ddns_hostname -match '\s|\.') {
                $FixedAddressWorksheetErrors += "The field ddns_hostname '$($F.ddns_hostname)' for ipv4addr '$($F.ipv4addr)' cannot contain spaces or dots. Fully qualified domain names are not needed."
            }
            #endregion

            #endregion

            #region Test name cannot contain spaces or dots
            if ($F.name -match '\s|\.') {
                $FixedAddressWorksheetErrors += "The field name '$($F.name)' for ipv4addr '$($F.ipv4addr)' cannot contain spaces or dots."
            }
            #endregion

            #region Test match_client
            $F.match_client = $F.match_client.ToUpper()

            $FixedAddressWorksheetErrors += switch ($F.match_client) {
                'MAC_ADDRESS' {
                    if (-not $F.mac) {
                        "The field 'mac' is missing for ipv4addr '$($F.ipv4addr)'. When 'match_client' is set to 'MAC_ADDRESS' the field 'mac' is mandatory"
                    }
                    break
                }
                'CIRCUIT_ID' {
                    if (-not $F.agent_circuit_id) {
                        "The field 'agent_circuit_id' is missing for ipv4addr '$($F.ipv4addr)'. When 'match_client' is set to 'CIRCUIT_ID' the field 'agent_circuit_id' is mandatory"
                    }
                    break
                }
                'CLIENT_ID' {
                    if (-not $F.dhcp_client_identifier) {
                        "The field 'dhcp_client_identifier' is missing for ipv4addr '$($F.ipv4addr)'. When 'match_client' is set to 'CLIENT_ID' the field 'dhcp_client_identifier' is mandatory"
                    }
                    break
                }
                'REMOTE_ID' {
                    if (-not $F.agent_remote_id) {
                        "The field 'agent_remote_id' is missing for ipv4addr '$($F.ipv4addr)'. When 'match_client' is set to 'REMOTE_ID' the field 'agent_remote_id' is mandatory"
                    }
                    break
                }
                'RESERVED' {
                    if (($F.mac) -and ($F.mac -ne '00:00:00:00:00:00')) {
                        "The field 'mac' cannot be set to '$($F.mac)' when 'match_client' is set to 'RESERVED' for ipv4addr '$($F.ipv4addr)'. When 'match_client' is set to 'RESERVED' the mac address will be set automatically by the API to '00:00:00:00:00:00', so please leave the field 'mac' blanc or set it to '00:00:00:00:00:00'. Or simply change the 'match_client' setting to 'MAC_ADDRESS' in case you want to use this mac address."
                    }
                    break
                }
                Default {
                    "The field 'match_client' contains the unsupported value '$_' for ipv4addr '$($F.ipv4addr)'. Only the following values are supported: 'MAC_ADDRESS','RESERVED', 'CLIENT_ID', 'CIRCUIT_ID' or 'REMOTE_ID'."
                }
            }
            #endregion

            #region Test valid mac
            if ($F.mac) {
                if (($F.mac -notmatch '^((([a-zA-z0-9]{2}[-:]){5}([a-zA-z0-9]{2}))$|^(([a-zA-z0-9]{2}:){5}([a-zA-z0-9]{2})))$')) {
                    $FixedAddressWorksheetErrors += "mac '$($F.mac)' is not a valid MAC address."
                    Continue
                }

                $F.mac = $F.mac.ToLower()
                $F.mac = $F.mac.Replace('-', ':')
            }
            #endregion

            #region Test when ipv4addr starts with 'func:' a valid mac is required
            if (
                ($F.ipv4addr -match '^func:') -and
                ((-not $F.mac) -or ($F.match_client -ne 'MAC_ADDRESS') -or ($F.mac -eq '00:00:00:00:00:00'))
            ) {
                $FixedAddressWorksheetErrors += "ipv4addr '$($F.ipv4addr)' with mac '$($F.mac)' and match_client '$($F.match_client)' is not valid. When ipv4addr starts with 'func:' the field match_client needs to be set to 'MAC_ADDRESS' and the mac address cannot be '00:00:00:00:00:00'."
            }
            #endregion
        }

        #region Test duplicate name
        $FixedAddressWorksheet.where( { $_.name }) |
        Group-Object name | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $FixedAddressWorksheetErrors += "The name '$($_.Name)' is used multiple times for different ipv4addr '$($_.Group.ipv4addr -join ', ')'. The name needs to be unique."
        }
        #endregion

        #region Test duplicate ddns_hostname
        $FixedAddressWorksheet.where( { $_.ddns_hostname }) |
        Group-Object ddns_hostname | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $FixedAddressWorksheetErrors += "The ddns_hostname '$($_.Name)' is used multiple times for different ipv4addr '$($_.Group.ipv4addr -join ', ')'. The ddns_hostname needs to be unique."
        }
        #endregion

        #region Test duplicate mac
        $FixedAddressWorksheet.where( { ($_.mac) -and ($_.mac -ne '00:00:00:00:00:00') }) |
        Group-Object mac | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $FixedAddressWorksheetErrors += "The mac '$($_.Name)' is used multiple times for different ipv4addr '$($_.Group.ipv4addr -join ', ')'. The mac address needs to be unique."
        }
        #endregion

        #region Test duplicate ipv4addr
        $FixedAddressWorksheet.where( { ($_.ipv4addr) -and ($_.ipv4addr -notmatch '^func:') }) |
        Group-Object ipv4addr | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $FixedAddressWorksheetErrors += "The ipv4addr '$($_.Name)' is used multiple times. The IP address needs to be unique."
        }
        #endregion
        #endregion

        #region Test duplicate records between worksheet 'FixedAddress' and 'Remove'
        foreach ($N in @('name', 'ddns_hostname')) {
            @($FixedAddressWorksheet + $RemoveWorksheet).Where( { $_.$N }) | Group-Object -Property $N |
            Where-Object { $_.count -ge 2 } | ForEach-Object {
                $ConflictingWorksheetErrors += "$N with value '$($_.Name)'"
            }
        }

        @($FixedAddressWorksheet + $RemoveWorksheet).where( { ($_.mac) -and ($_.mac -ne '00:00:00:00:00:00') }) |
        Group-Object mac | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $ConflictingWorksheetErrors += "$N with value '$($_.Name)'"
        }

        @($FixedAddressWorksheet + $RemoveWorksheet).where( { ($_.ipv4addr) -and ($_.ipv4addr -notmatch '^func:') }) |
        Group-Object ipv4addr | Where-Object { $_.count -ge 2 } | ForEach-Object {
            $ConflictingWorksheetErrors += "ipv4addr with value '$($_.Name)'"
        }
        #
        #endregion

        #region Send mail to users on incorrect input
        if ($FixedAddressWorksheetErrors -or $RemoveWorksheetErrors -or $ConflictingWorksheetErrors) {
            $MailParams = @{
                LogFolder = $LogParams.LogFolder
                Header    = $ScriptName
                Save      = $LogFile + ' - Mail.html'
                To        = $MailTo
                Bcc       = $ScriptAdmin
                Message   = "<p>Incorrect data found in the Excel import file.</p>"
                Subject   = 'FAILURE - Incorrect input'
                Priority  = 'High'
            }

            if ($FixedAddressWorksheetErrors) {
                $MailParams.Message += ("<p>Worksheet '<b>FixedAddress</b>':</p>" + ($FixedAddressWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'FixedAddress':`n`n- $(($FixedAddressWorksheetErrors | Select-Object -First 10) -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage
                Write-Warning $WarningMessage
            }
            if ($RemoveWorksheetErrors) {
                $MailParams.Message += ("<p>Worksheet '<b>Remove</b>':</p>" + ($RemoveWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'Remove':`n`n- $($RemoveWorksheetErrors -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage
                Write-Warning $WarningMessage
            }
            if ($ConflictingWorksheetErrors) {
                $MailParams.Message += ("<p>Worksheet '<b>FixedAddress</b>' and '<b>Remove</b>' contain duplicate records:</p>" + ($ConflictingWorksheetErrors | ConvertTo-HtmlListHC))

                $WarningMessage = "Worksheet 'FixedAddress' and 'Remove' contain duplicate records:`n`n- $($ConflictingWorksheetErrors -join "`n")"
                Write-EventLog @EventErrorParams -Message $WarningMessage
                Write-Warning $WarningMessage
            }
            Send-MailHC @MailParams

            Write-EventLog @EventEndParams; Exit
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $IpamParams = @{
            Environment = $Environment
            ErrorAction = 'Stop'
        }

        $IpamWorksheet = @()

        $DNSResolveProps = @(
            @{N = 'Hostname'; E = {
                    if ($_.NameHost) {
                        $_.NameHost -replace ".$env:USERDNSDOMAIN"
                    }
                    else {
                        $_.Name -replace ".$env:USERDNSDOMAIN"
                    }
                }
            },
            @{N = 'IP'; E = {
                    if ($_.IPAddress) {
                        $_.IPAddress -replace '.in-addr.arpa.'
                    }
                    else {
                        $_.Name -replace '.in-addr.arpa.'
                    }
                }
            },
            'name', 'ddns_hostname'
        )

        $IpamFixedAddressList = @(Get-IpamFixedAddressHC @IpamParams)

        foreach ($F in $FixedAddressWorksheet) {
            $AddAddressReservation = $TestOnline = $Incorrect = $false

            $IpamBody = Copy-ObjectHC $F

            #region Remove properties that the API can't handle
            $IpamBody.PSObject.Properties.Name.Where( { $ApiPropertyList -notcontains $_ }).ForEach( {
                    $IpamBody.PSObject.Properties.Remove($_)
                })
            #endregion

            #region Set match_client and mac to match the API return values
            if ((($IpamBody.match_client -eq 'MAC_ADDRESS') -and ($IpamBody.mac -eq '00:00:00:00:00:00')) -or
                ($IpamBody.match_client -eq 'RESERVED')) {
                $IpamBody.match_client = $IpamBody.match_client = 'RESERVED'
                $IpamBody.mac = '00:00:00:00:00:00'
            }
            #endregion

            Try {
                #region Create filter to find existing address reservations
                $ipv4addrFunction = $F.ipv4addr -match '^func:'

                $Filter = @()

                if ($ipv4addrFunction) {
                    $CompareProperties = $IpamBody.PSObject.Properties.Name.where( { $_ -ne 'ipv4addr' })
                }
                else {
                    $Filter += "`$_.ipv4addr -eq `$F.ipv4addr"
                    $CompareProperties = $IpamBody.PSObject.Properties.Name
                }

                if (($F.mac) -and ($F.mac -ne '00:00:00:00:00:00')) {
                    $Filter += "`$_.mac -eq `$F.mac"
                }

                if ($F.ddns_hostname) {
                    $Filter += "`$_.ddns_hostname -eq `$F.ddns_hostname"
                    $Filter += "`$_.ddns_hostname -eq `"`$(`$F.ddns_hostname)`.`$env:USERDOMAIN`""
                    $Filter += "`$_.ddns_hostname -eq `"`$(`$F.ddns_hostname)`.`$env:USERDNSDOMAIN`""
                }

                if ($F.name) {
                    $Filter += "`$_.name -eq `$F.name"
                    $Filter += "`$_.name -eq `"`$(`$F.name)`.`$env:USERDOMAIN`""
                    $Filter += "`$_.name -eq `"`$(`$F.name)`.`$env:USERDNSDOMAIN`""
                }

                $WhereClause = [ScriptBlock]::Create($Filter -join ' -or ')
                #endregion

                if ($IpamFixedAddress = $IpamFixedAddressList.Where($WhereClause)) {
                    if ($IpamFixedAddress.count -ge 2) {
                        $IpamWorksheet += $IpamFixedAddress

                        $Problem = $IpamFixedAddress |
                        Select-Object @{N = 'Result'; E = { "$($_.name) - $($_.ddns_hostname) - $($_.ipv4addr) - $($_.mac) " } }

                        $F.Status = 'Error'
                        $F.Error = "mac, ip, name or ddns_hostname already known in IPAM with $($IpamFixedAddress.count) other fixed address reservations '$($Problem.Result -join ', ')'."
                        Continue
                    }

                    $CompParams = @{
                        ReferenceObject  = $IpamBody
                        DifferenceObject = $IpamFixedAddress
                        Property         = $CompareProperties
                    }
                    $Incorrect = Compare-ConfigHC @CompParams

                    if ($Incorrect) {
                        $IpamWorksheet += $IpamFixedAddress
                        $F.Status = 'Incorrect'
                        $F.IncorrectFields = $Incorrect.Keys -join ', '

                        Write-EventLog @EventWarnParams -Message "ipv4addr '$($F.ipv4addr)' is incorrect, incorrect fields '$($F.IncorrectFields)'"

                        if ($F.OverWrite) {
                            Remove-IpamObjectHC @IpamParams -ReferenceObject $IpamFixedAddress -NoServiceRestart
                            $F.Action += 'Removed address reservation'
                            $F.Status = 'Removed'

                            Write-EventLog @EventOutParams -Message "Removed reservation for ipv4addr '$($F.ipv4addr)' for incorrect fields '$($F.IncorrectFields)'"

                            $AddAddressReservation = $true
                            $TestOnline = $false
                        }
                        else {
                            $TestOnline = $true
                        }
                    }
                    else {
                        $F.Status = 'Ok'
                        # Write-EventLog @EventVerboseParams -Message "ipv4addr '$($F.ipv4addr)' is correct"
                    }
                }
                else {
                    $AddAddressReservation = $true

                    if (-not $F.OverWrite) {
                        $TestOnline = $true
                    }
                }

                if ($AddAddressReservation) {
                    if ($TestOnline) {
                        #region Test if ipv4addr and ddns_hostname is online
                        $ConnectionParams = @{
                            ComputerName = @()
                            Count        = 1
                            ErrorAction  = 'Ignore'
                        }

                        if (-not $ipv4addrFunction) {
                            $ConnectionParams.ComputerName += $F.ipv4addr
                        }

                        if ($F.enable_ddns) {
                            $ConnectionParams.ComputerName += $F.ddns_hostname
                        }

                        if ($ConnectionParams.ComputerName) {
                            $Problem = $null

                            if ($Connection = @(Test-Connection @ConnectionParams)) {
                                $Problem += "Address online '$($Connection.foreach( { $_.Address }) -join ', ')'. "
                            }

                            if ($ResolveDNS = $ConnectionParams.ComputerName.foreach( {
                                        Resolve-DnsName $_ -EA Ignore | Select-Object $DNSResolveProps |
                                        Select-Object @{N = 'Result'; E = { "$($_.Hostname) > $($_.IP)" } }
                                    })) {
                                $Problem += "Address in DNS '$($ResolveDNS.Result -join ', ')'."
                            }

                            if ($Problem) {
                                throw $Problem
                            }
                        }
                        #endregion
                    }

                    New-IpamFixedAddressHC @IpamParams -Body $IpamBody -NoServiceRestart
                    $F.Action += 'Added new reservation'
                    $F.Status = 'Added'

                    if ($Incorrect) {
                        $F.Status = 'Updated'
                    }

                    Write-EventLog @EventOutParams -Message "Added new reservation for ipv4addr '$($F.ipv4addr)'"
                }
            }
            Catch {
                Write-Warning $_
                $F.Status = 'Error'
                $F.Error = $_

                Write-EventLog @EventErrorParams -Message "Error adding address reservation for ipv4addr '$($F.ipv4addr)'`r`n-$($F.Error)"


                $Error.Remove($Error[0])

                if (($F.mac -ne '00:00:00:00:00:00') -and
                    ($IpamFixedAddress = Get-IpamFixedAddressHC @IpamParams -Filter "mac=$($F.mac)")) {
                    $IpamWorksheet += $IpamFixedAddress
                }
                Continue
            }
        }

        if ($FixedAddressWorksheet.Action) {
            Restart-IpamServiceHC @IpamParams
        }

        $IpamFixedAddressList = @(Get-IpamFixedAddressHC @IpamParams)

        foreach ($R in $RemoveWorksheet) {
            Try {
                $RowText = "name '$($R.name)' ddns_hostname '$($R.ddns_hostname)' ipv4addr '$($R.ipv4addr)' mac '$($R.mac)'"

                #region Create filter to find existing address reservations
                $Filter = @()

                if ($R.ipv4addr) {
                    $Filter += "`$_.ipv4addr -eq `$R.ipv4addr"
                }

                if ($R.mac) {
                    $Filter += "`$_.mac -eq `$R.mac"
                }

                if ($R.ddns_hostname) {
                    $Filter += "`$_.ddns_hostname -eq `$R.ddns_hostname"
                    $Filter += "`$_.ddns_hostname -eq `"`$(`$R.ddns_hostname)`.`$env:USERDOMAIN`""
                    $Filter += "`$_.ddns_hostname -eq `"`$(`$R.ddns_hostname)`.`$env:USERDNSDOMAIN`""
                }

                if ($R.name) {
                    $Filter += "`$_.name -eq `$R.name"
                    $Filter += "`$_.name -eq `"`$(`$R.name)`.`$env:USERDOMAIN`""
                    $Filter += "`$_.name -eq `"`$(`$R.name)`.`$env:USERDNSDOMAIN`""
                }

                if (-not $Filter) {
                    throw 'The filter to find IPAM address reservations cannot be empty.'
                }

                $WhereClause = [ScriptBlock]::Create($Filter -join ' -or ')
                #endregion

                if ($IpamFixedAddress = $IpamFixedAddressList.Where($WhereClause)) {
                    $IpamWorksheet += $IpamFixedAddress

                    if ($IpamFixedAddress.count -ge 2) {
                        $Problem = $IpamFixedAddress |
                        Select-Object @{N = 'Result'; E = { "$($_.name) - $($_.ddns_hostname) - $($_.ipv4addr) - $($_.mac)" } }

                        $R.Status = 'Error'
                        $R.Error = "mac, ip, name or ddns_hostname known in IPAM with $($IpamFixedAddress.count) other fixed address reservations '$($Problem.Result -join ', ')'."
                        Continue
                    }

                    Remove-IpamObjectHC @IpamParams -ReferenceObject $IpamFixedAddress -NoServiceRestart
                    $R.Status = 'Removed'
                    $R.Action += 'Removed address reservation'

                    Write-EventLog @EventOutParams -Message "Removed reservation for $RowText"
                }
                else {
                    $R.Status = 'Not found'
                }
            }
            Catch {
                Write-Warning $_
                $R.Status = 'Error'
                $R.Error = $_

                Write-EventLog @EventErrorParams -Message "Error removing address reservation for $RowText`r`n-$($R.Error)"

                $Error.Remove($Error[0])
            }
        }

        if ($RemoveWorksheet.Action) {
            Restart-IpamServiceHC @IpamParams
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        $Intro = $FixedAddressSummaryTable = $RemoveSummaryTable = $ExcelWorksheetDescription = $Subject = $null

        $MailParams = @{
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
            To        = $MailTo
            Bcc       = $ScriptAdmin
        }

        #region Export to Excel
        $ExcelParams = @{
            Path               = $LogFile + '.xlsx'
            NoNumberConversion = ($FixedAddressWorksheetProperties.GetEnumerator().where( { $_.Value.Type -eq 'String' })).Name
            AutoSize           = $true
            FreezeTopRow       = $true
        }

        if (Test-Path -Path $ExcelParams.Path) {
            Write-Warning "Excel file '$($ExcelParams.Path)' exists already and will be removed"
            Remove-Item -LiteralPath $ExcelParams.Path -Force
        }

        if ($FixedAddressWorksheet) {
            $FixedAddressWorksheet | Select-Object *, @{N = 'Action'; E = { $_.Action -join ', ' } } -ExcludeProperty Action |
            Export-Excel @ExcelParams -WorksheetName 'FixedAddress' -TableName 'FixedAddresses'

            $MailParams.Attachments = $ExcelParams.Path
        }

        if ($RemoveWorksheet) {
            $RemoveWorksheet | Select-Object *, @{N = 'Action'; E = { $_.Action -join ', ' } } -ExcludeProperty Action |
            Export-Excel @ExcelParams -WorksheetName 'Remove' -TableName 'Remove'

            $MailParams.Attachments = $ExcelParams.Path
        }

        if ($IpamWorksheet) {
            $IpamWorksheet | Export-Excel @ExcelParams -WorksheetName 'IPAM' -TableName 'IPAM'

            $MailParams.Attachments = $ExcelParams.Path
        }

        if ($Error) {
            $Error.Exception.Message.ForEach( {
                    Write-EventLog @EventErrorParams -Message $_
                })

            $Error.Exception.Message | Select-Object @{N = 'Error message'; E = { $_ } } |
            Export-Excel @ExcelParams -WorksheetName 'Errors' -TableName 'Errors'

            $MailParams.Attachments = $ExcelParams.Path
        }
        #endregion

        #region Add summary table FixedAddress
        $FixedAddressCountTotal = $FixedAddressWorksheet.Count
        $FixedAddressCountError = @($FixedAddressWorksheet.Where( { $_.Status -eq 'Error' })).Count
        $FixedAddressCountOk = @($FixedAddressWorksheet.Where( { $_.Status -eq 'Ok' })).Count
        $FixedAddressCountIncorrect = @($FixedAddressWorksheet.Where( { $_.Status -eq 'Incorrect' })).Count
        $FixedAddressCountAdded = @($FixedAddressWorksheet.Where( { $_.Status -eq 'Added' })).Count
        $FixedAddressCountUpdated = @($FixedAddressWorksheet.Where( { $_.Status -eq 'Updated' })).Count

        if ($FixedAddressWorksheet) {
            $FixedAddressSummaryTable = "
            <p><i>Worksheet 'FixedAddress':</i></p>
            <table>
                <tr><th>Quantity</th><th>Status</th></tr>
                $(if ($FixedAddressCountOk) {"<tr><td style=``"text-align: center``">$FixedAddressCountOk</td><td>OK</td></tr>"})
                $(if ($FixedAddressCountAdded) {"<tr><td style=``"text-align: center``">$FixedAddressCountAdded</td><td>Added</td></tr>"})
                $(if ($FixedAddressCountUpdated) {"<tr><td style=``"text-align: center``">$FixedAddressCountUpdated</td><td>Updated</td></tr>"})
                $(if ($FixedAddressCountIncorrect) {"<tr><td style=``"text-align: center``">$FixedAddressCountIncorrect</td><td>Incorrect</td></tr>"})
                $(if ($FixedAddressCountError) {"<tr><td style=``"text-align: center``">$FixedAddressCountError</td><td>Error</td></tr>"})
                <tr><td style=`"text-align: center`"><b>$FixedAddressCountTotal</b></td><b>Total</b></tr>
            </table>"
        }
        #endregion

        #region Add summary table Remove
        $RemoveCountTotal = $RemoveWorksheet.Count
        $RemoveCountError = @($RemoveWorksheet.Where( { $_.Status -eq 'Error' })).Count
        $RemoveCountNotFound = @($RemoveWorksheet.Where( { $_.Status -eq 'Not found' })).Count
        $RemoveCountRemoved = @($RemoveWorksheet.Where( { $_.Status -eq 'Removed' })).Count

        if ($RemoveWorksheet) {
            $RemoveSummaryTable = "
            <p><i>Worksheet 'Remove':</i></p>
            <table>
                <tr><th>Quantity</th><th>Status</th></tr>
                $(if ($RemoveCountRemoved) {"<tr><td style=``"text-align: center``">$RemoveCountRemoved</td><td>Removed</td></tr>"})
                $(if ($RemoveCountNotFound) {"<tr><td style=``"text-align: center``">$RemoveCountNotFound</td><td>Not found</td></tr>"})
                $(if ($RemoveCountError) {"<tr><td style=``"text-align: center``">$RemoveCountError</td><td>Error</td></tr>"})
                <tr><td style=`"text-align: center`"><b>$RemoveCountTotal</b></td><b>Total</b></tr>
            </table>"
        }
        #endregion

        #region Format mail message, subject and priority
        if ($Error) {
            $Subject = "FAILURE - $FixedAddressCountTotal fixed addresses, $($Error.Count) errors"

            $Intro = "Failed to add/update fixed addresses due to <b>$($Error.Count) errors</b> that were encountered during execution."

            $ExcelWorksheetDescription = "<p><i>* Please verify the worksheet 'Errors' in attachment</i></p>"
        }
        elseif ($FixedAddressCountError -or $RemoveCountError) {
            $Subject = "FAILURE - $FixedAddressCountTotal fixed addresses, {0} errors" -f
            $($FixedAddressCountError + $RemoveCountError)

            if ($FixedAddressCountError -and $RemoveCountError) {
                $Intro = "Failed to add/update <b>$FixedAddressCountError fixed addresses</b> amd failed to remove <b>$RemoveCountError fixed addresses</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheets 'FixedAddress' and the worksheet 'Remove' in attachment.</i></p>"
            }
            elseif ($FixedAddressCountError) {
                $Intro = "Failed to add/update <b>$FixedAddressCountError fixed addresses</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheet 'FixedAddress' in attachment.</i></p>"
            }
            elseif ($RemoveCountError) {
                $Intro = "Failed to remove <b>$RemoveCountError fixed addresses</b>."
                $ExcelWorksheetDescription = "<p><i>* Please verify the 'Error' column in the worksheet 'Remove' in attachment.</i></p>"
            }

        }
        elseif ($FixedAddressCountIncorrect) {
            $Subject = "$FixedAddressCountTotal fixed addresses, $FixedAddressCountIncorrect incorrect"

            $Intro = "We found <b>$FixedAddressCountIncorrect incorrect</b> {0}." -f
            $(if ($FixedAddressCountIncorrect -eq 1) { 'fixed address' }else { 'fixed addresses' })

            $ExcelWorksheetDescription = "<p><i>* Please verify the column 'IncorrectFields' in the worksheet 'FixedAddress'. The worksheet 'IPAM' contains the details of the incorrect fixed address as it is known in IPAM at the time of execution.</i></p>" -f
            $(if ($FixedAddressCountIncorrect -eq 1) { 'fixed address' }else { 'fixed addresses' })
        }
        elseif ($FixedAddressCountTotal -eq $FixedAddressCountOk) {
            $Subject = "$FixedAddressCountTotal fixed addresses, all correct"

            $Intro = "All fixed addresses are correct, no changes done."

            $ExcelWorksheetDescription = "<p><i>* Please find the overview in attachment.</i></p>"
        }
        else {
            $Subject = "$FixedAddressCountTotal fixed addresses"

            if ($FixedAddressCountAdded -and $FixedAddressCountUpdated) {
                $Intro = "Successfully <b>added $FixedAddressCountAdded</b> and <b>updated $FixedAddressCountUpdated</b> fixed addresses."
                $Subject = "$FixedAddressCountTotal fixed addresses, $FixedAddressCountAdded added, $FixedAddressCountUpdated updated"
            }
            if ($FixedAddressCountAdded -and (-not $FixedAddressCountUpdated)) {
                $Intro = "Successfully <b>added $FixedAddressCountAdded</b> {0}" -f
                $(if ($FixedAddressCountAdded -eq 1) { 'fixed address.' }else { 'fixed addresses.' })

                $Subject = "$FixedAddressCountTotal fixed addresses, $FixedAddressCountAdded added"
            }
            if ((-not $FixedAddressCountAdded) -and $FixedAddressCountUpdated) {
                $Intro = "Successfully <b>updated $FixedAddressCountUpdated</b> {0}" -f
                $(if ($FixedAddressCountUpdated -eq 1) { 'fixed address.' }else { 'fixed addresses.' })

                $Subject = "$FixedAddressCountTotal fixed addresses, $FixedAddressCountUpdated updated"
            }

            $ExcelWorksheetDescription = "<p><i>* Please find in attachment an overview. All changes can be found in the field 'Action' of the worksheet 'FixedAddresses'. $(if($FixedAddressCountUpdated) {"The worksheet 'IPAM' contains the details in IPAM before we updated the object."})</i></p>"
        }

        $MailParams.Subject = $Subject
        $MailParams.Priority = if ($Error -or $FixedAddressCountError -or $RemoveCountError) { 'High' } else { 'Normal' }
        $MailParams.Message = "
                $Intro
                $FixedAddressSummaryTable
                $RemoveSummaryTable
                $ExcelWorksheetDescription
        "
        #endregion

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message  "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}