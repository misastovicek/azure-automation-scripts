[CmdletBinding()]
param (
    [Parameter()]
    [object]
    $WebhookData
    )
    
$TeamsWebhookURI = Get-AutomationVariable -Name TeamsWebhook

if ($WebhookData) {
    $WebhookBody = $WebhookData.RequestBody | ConvertFrom-Json
    $schemaId = $WebhookBody.schemaId

    if ($schemaId -eq "azureMonitorCommonAlertSchema") {
        # This is the common Metric Alert schema (released March 2019)
        $Essentials = [object] $WebhookBody.data.essentials
        $AlertContext = [object] $WebhookBody.data.alertContext

        $AlertRuleName = $Essentials.alertRule
        $AlertCondition = $Essentials.monitorCondition
        $AlertDescription = $Essentials.description
        
        $AlertResultsFacts = [System.Collections.ArrayList]@()
        $AlertResultsLinksFacts = [System.Collections.ArrayList]@()
        $AlertColumnSets = [System.Collections.ArrayList]@()
        
        if ($Essentials.signalType -eq "Log"){
            if ($Essentials.monitoringService -eq "Log Alerts V2") {
                
                $AlertContext.condition.allOf | ForEach-Object {
                    $AlertResultsLinksFacts.Add(
                        @{
                            'title' = 'Query (as link):'
                            'value' = "[" + $_.searchQuery + "](" + $_.linkToFilteredSearchResultsUI + ")"
                        }
                        )
                    } | Out-Null
                }
                elseif ($Essentials.monitoringService -eq "Log Analytics" -or $Essentials.monitoringService -eq "Application Insights") {
                    $AlertResultsLinksFacts.Add(
                        @{
                            'title' = 'Query Results:'
                            'value' = '[Link to Query Results in Azure Portal](' + $AlertContext.linkToFilteredSearchResultsUI + ')'
                        }
                    )
                    if ($AlertContext.SearchResults.tables) {
                        $TableColumns = $AlertContext.SearchResults.tables[0].columns
                        $TableRows = $AlertContext.SearchResults.tables[0].rows
                        
                        for ($row = 0; $row -lt ($TableRows.length,5 | Measure-Object -Minimum).Minimum; $row++){
                            $i = 0
                            $TeamsRow = @{
                                "type" = "ColumnSet"
                                "columns" = [System.Collections.ArrayList]@()
                            }
                            for ($j = 0; $j -lt $TableColumns.length; $j++) {
                                $TeamsRow.columns.Add(
                                    @{
                                        "type" = "Column"
                                        "items" = [System.Collections.ArrayList]@()
                                    }
                                )
                            }
                            foreach ($item in $TableRows[$row]) {
                                if ($i -ge $TableColumns.length){
                                    throw "Column number $i doesn't exist in the received Table!"
                                }
                                ($TeamsRow.columns[$i].items).Add(
                                    @{
                                        "type" = "TextBlock"
                                        "text" = $TableRows[$row][$i]
                                        "wrap" = $true
                                    }
                                )
                                $i++
                            }
                            $AlertColumnSets.Add($TeamsRow)
                        }
                    }
                }
            }
        elseif ($Essentials.signalType -eq "Metric") {
            $AlertContext.condition.allOf | ForEach-Object {
                $AlertResultsFacts.Add(
                    @{
                        'title' = $_.metricName
                        'value' = $_.metricValue
                    }
                )
            } | Out-Null
        }
        else {
            Write-Error "The signal type - " + $Essentials.signalType + " - !"
        }
    }
    elseif ($schemaId -eq "AzureMonitorMetricAlert") {
        # This is the near-real-time Metric Alert schema
        $AlertContext = [object] ($WebhookBody.data).context
    }
    elseif ($null -eq $schemaId) {
        # This is the original Metric Alert schema
        $AlertContext = [object] $WebhookBody.context
    }
    else {
        # Schema not supported
        Write-Error "The alert data schema - $schemaId - is not supported."
    }

    $Facts = [System.Collections.ArrayList]@()

    if ($AlertResultsLinksFacts) {
        $Facts.Add(
            @{
                "type" = "FactSet"
                "facts" = $AlertResultsLinksFacts 
            }
        ) | Out-Null
    }

    if ($AlertResultsFacts) {
        $Facts.Add(
            @{
                "type" = "FactSet"
                "facts" = $AlertResultsFacts 
            }
        ) | Out-Null
    }
    $AlertColor = if ($AlertCondition -eq "Fired") {"Attention"} else {"Good"}
    $TeamsCardSchema = @{
        '$schema' = "http://adaptivecards.io/schemas/adaptive-card.json"
        "type" = "AdaptiveCard"
        "version" = "1.2"
        "body" = [System.Collections.ArrayList]@(
            @{
                "type" = "TextBlock"
                "text" = "**[$AlertCondition]** - **$AlertRuleName**"
                "size" = "medium"
                "weight" = "bolder"
                "color" = $AlertColor
            }
            @{
                "type" = "TextBlock"
                "text" = "$AlertDescription"
                "size" = "default"
                "weight" = "default"
                "wrap" = $true
            }
            $Facts
        )
    }

    if ($AlertColumnSets.length -gt 0) {
        foreach ($column in $AlertColumnSets) {
            ($TeamsCardSchema.body).Add($column)
        }
    }

    $TeamsCardSchema | ConvertTo-Json -Depth 50

    $TeamsWebhookPayload = @{
        "type" = "message"
        "attachments" = @(
            @{
                "contentType" = "application/vnd.microsoft.card.adaptive"
                "contentUrl" = $null
                "content" = $TeamsCardSchema
            }
        )
    } | ConvertTo-Json -Depth 50

    Invoke-WebRequest -Method Post -ContentType 'Application/Json' -Uri $TeamsWebhookURI -Body $TeamsWebhookPayload -UseBasicParsing
}
else {
    Write-Output "No data passed to the script"
}