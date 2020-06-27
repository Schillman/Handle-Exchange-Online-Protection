function Use-SnowRestApi {
    <#
.Synopsis
   Service Now Functions for Create, Update, and Get Incidents, Changes and Requests
.DESCRIPTION
   This PowerShell function provides a single way of interacting with the ServiceNow REST API, normally performed with a series of cmdlet's.
   We utilize the switch statement on the parameter '-Method' to understand what statement to use, then wrapping it in, using the Invoke-RestMethod cmdlet.

    Valid METHOD's values are GET, PATCH & POST..

.EXAMPLE
   -Method 'GET'
     Get a specified item '{sys_id}' from {TableName}...
   $Result = Use-SnowRestApi -SnowCredentialsFile '.\SNOWCredentials.xml' -Method 'GET' -SnowUrl 'https://YourInstance.service-now.com/' -ApiCall '/api/now/table/{tableName}/{sys_id}'

.EXAMPLE
   -Method 'PATCH'
     Adds a comment to specified {sys_id} located under {TableName}
   $Result = Use-SnowRestApi -SnowCredentialsFile '.\SNOWCredentials.xml' -Method 'PATCH' -SnowUrl 'https://YourInstance.service-now.com/' -ApiCall '/api/now/table/{tableName}/{sys_id}' -JsonQuery @{ comments = "Additional Coments" }

.EXAMPLE
   -Method 'POST'
     Inserts one record in the specified table. Multiple record insertion is not supported by this method.
   $Result = Use-SnowRestApi -SnowCredentialsFile '.\SNOWCredentials.xml' -Method 'PATCH' -SnowUrl 'https://YourInstance.service-now.com/' -ApiCall '/api/now/table/{tableName}/{sys_id}' -JsonQuery @{ short_description = 'Unable to connect to office wifi'; assignment_group = '287ebd7da9fe198100f92cc8d1d2154e'; urgency  = '2'; impact = '2'}

    #>
    [CmdletBinding()]
    Param
    (
        # CredentialFile for Snow Access to the above URL .\Credfile.xml
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $SnowCredentialsFile,

        # Specifies the method used for the web request. Valid values are Default, Delete, Get, Head, Merge, Options, Patch, Post, Put, and Trace.
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, Position = 1)]
        [ArgumentCompleter({@('DELETE', 'GET', 'HEAD', 'MERGE', 'OPTIONS', 'PATCH', 'POST', 'PUT', 'TRACE')})]
        $Method,

        # "https://YourInstance.service-now.com/"
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 2)]
        $SnowUrl,

        # "/api/now/table/{tableName}/{sys_id}"
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 3)]
        $ApiCall,

        <# 
        @{
        comments = "Additional Coments" # Add the comments you want to add to the Change. 
        }
        #>
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 4)]
        $JsonQuery
    )

    Begin {
        #region if doesn't exist, prompting and save credentials into an xml-file
        if (!(Test-Path $SnowCredentialsFile)) {
            $SNOWCredentials = Get-Credential -Message "Enter SNOW login credentials"
            $SNOWCredentials | Export-Clixml $SnowCredentialsFile -Force
        }
        #endregion

        #region Importing credentials
        $SNOWCredentials = Import-Clixml $SnowCredentialsFile
        $SNOWUsername = $SNOWCredentials.UserName
        $SNOWPassword = $SNOWCredentials.GetNetworkCredential().Password
        #endregion

        #region Building Authentication Header & setting content type
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::tls12
        $HeaderAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $SNOWUsername, $SNOWPassword)))
        $Global:Header = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $Global:Header.Add('Authorization', ('Basic {0}' -f $HeaderAuth))
        $Global:Header.Add('Accept', 'application/json')
        $Type = "application/json"
        #endregion

        $Url = $SNOWURL + $ApiCall

        if ($JsonQuery) {
            $Body = $JsonQuery | ConvertTo-Json
        }
    }
    Process {
        $error.Clear()
        switch ($Method) {
            {$_ -eq 'GET'} {
                Try {
                    $JsonResponse = Invoke-RestMethod -Method $Method -Uri $Url -Headers $Header -ContentType $Type -TimeoutSec 25
                    $Response = $JsonResponse.result
                } Catch {
                    Return $_
                }
            }

            {$_ -eq 'PATCH'} {
                if (!$Body) {
                    Return "JsonQuery missing. Please populate parameter '-JsonQuery'"
                }
                Try {
                    $JsonResponse = Invoke-RestMethod -Method $Method -Uri $Url -Headers $Header -Body $Body -ContentType $Type -TimeoutSec 25
                    $Response = $JsonResponse.result
                    if (!$error) {
                        $Response = 'Success'
                    } else {$Response = "Failed: $($error)"}

                } Catch {
                    Return $_
                }
            }

            {$_ -eq 'POST'} {
                if (!$Body) {
                    Return "JsonQuery missing. Please populate parameter '-JsonQuery'"
                }
                Try {
                    $JsonResponse = Invoke-RestMethod -Method $Method -Uri $Url -Headers $Header -Body $Body -ContentType $Type -TimeoutSec 25
                    $Response = $JsonResponse.result
                } Catch {
                    Return $_
                }
            }

            {$_ -eq 'PUT'} {
            if (!$Body) {
                    Return "JsonQuery missing. Please populate parameter '-JsonQuery'"
                }
                Try {
                    $JsonResponse = Invoke-RestMethod -Method $Method -Uri $Url -Headers $Header -Body $Body -ContentType $Type -TimeoutSec 25
                    $Response = $JsonResponse.result
                } Catch {
                    Return $_
                }
            }

            Default {Return 'Not a vaild Method'}
        }
    }
     End {
        if (!$error) {
        Return $Response
        } else { Write-Host $Error[0].Exception.Message$Error[0].ScriptStackTrace }
    }
}
