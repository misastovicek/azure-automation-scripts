Disable-AzContextAutosave –Scope Process

$ExpirationDays = $null
try {
    $ExpirationDays = Get-AutomationVariable -Name SPExpirationDays -ErrorAction Stop | Invoke-Expression
}
catch {

    $ExpirationDays = 0 .. 5 + 7, 10, 15, 20, 30
}

$TeamsWebHookUri = Get-AutomationVariable -Name TeamsWebhook

$connection = Get-AutomationConnection -Name AzureRunAsConnection
while (!($connectionResult) -and ($logonAttempt -le 10)) {
    $LogonAttempt++
    $connectionResult = Connect-AzAccount `
        -ServicePrincipal `
        -Tenant $connection.TenantID `
        -ApplicationId $connection.ApplicationID `
        -CertificateThumbprint $connection.CertificateThumbprint

    Start-Sleep -Seconds 30
}

try {
    $AzureADApps = Get-AzADApplication -ErrorAction Stop
}
catch {
    Write-Error "Can't get the Azure AD applications"
    exit 1
}

$AppsWithCredentials = foreach($App in $AzureADApps) {
    try {
        $Credentials = Get-AzADAppCredential -ApplicationId $App.ApplicationId -ErrorAction Stop
    }
    catch {
        Write-Error "Can't get credentials for Azure AD app '$app.DisplayName' (ID: $app.applicationId)"
        continue
    }
    if (!$Credentials.EndDate) {
        continue
    }

    [PSCustomObject]@{
        'ApplicationId' = $App.ApplicationId;
        'DisplayName'   = $App.DisplayName;
        'Keys'          = foreach ($Credential in $Credentials) {
                [PSCustomObject]@{
                'KeyType' = $Credentials.Type
                'KeyId'   = $Credential.KeyId
                'EndDate' = $Credential.EndDate
            }
        }
    }
}

foreach ($App in $AppsWithCredentials) {
    foreach ($ExpirationDay in $ExpirationDays) {
        foreach ($AppKey in $App.Keys) {
            if ((Get-Date $AppKey.EndDate).Date.AddDays(-$ExpirationDay) -eq (Get-Date).Date) {
                $TeamsMessageJSONBody = [PSCustomObject][Ordered]@{
                    "@type"      = "MessageCard"
                    "@context"   = "http://schema.org/extensions"
                    "summary"    = "Service Principal Expiration Warning!"
                    "themeColor" = '0078D7'
                    "sections"   = @(
                        @{
                            "activityTitle" = "Service Principal Exires in $ExpirationDay days!"
                            "facts"         = @(
                                @{
                                    "name"  = "Application Name"
                                    "value" = $App.DisplayName
                                },
                                @{
                                    "name"  = "Application ID"
                                    "value" = $App.ApplicationId
                                },
                                @{
                                    "name"  = "Key Type"
                                    "value" = $AppKey.KeyType
                                },
                                @{
                                    "name" = "Key ID"
                                    "value" = $AppKey.KeyId
                                },
                                @{
                                    "name"  = "Expires at"
                                    "value" = $AppKey.EndDate
                                }
                            )
                        }
                    )
                }
                $TeamMessageBody = ConvertTo-Json $TeamsMessageJSONBody -Depth 10
                $TeamsWebhookParameters = @{
                    "URI"         = $TeamsWebHookUri
                    "Method"      = 'POST'
                    "Body"        = $TeamMessageBody
                    "ContentType" = 'application/json'
                }
                Invoke-RestMethod -Uri $TeamsWebhookParameters.URI -Method $TeamsWebhookParameters.Method `
                    -ContentType $TeamsWebhookParameters.ContentType -Body $TeamsWebhookParameters.Body
            }
        }
    }
}