# Scripts for Azure Automation Account

This directory contains scripts which can be used with Azure Automation Account

## SPExpirationChecker.ps1

Purpose of this script is to look for credentials of the Service Principals in the Azure Active Directory and alert to MS Teams when they are about to expire.

The script reads two variables from the Azure Automation Account:
* **TeamsWebhook** - This is the Webhook URL created in MS Teams.
* **SPExpirationDays (OPTIONAL)** - This is a string value of dayes which should be checked for expiration.
For example: `1 .. 5 + 10, 15, 30` will be loaded as a list into a script variable which would be otherwise defined as `$MyVar = @(1, 2, 3, 4, 5, 10, 15, 30)`. If not defined, the default is `0 .. 5 + 7, 10, 15, 20, 30`.

## TeamsAlert.ps1

This script can be used for alerting from Azure Monitor to MS Teams using Azure Automation Account. You may proceed with Azure Logic Apps if you wish, but I found this solution as a better one since I don't really like the Teams integration in the Logic Apps.

The script reads one variable from the Azure Automation Account:
* **TeamsWebhook** - This is the Webhook URL created in MS Teams.

>**WARNING:** The script works properly only with the Azure **Common Alert Schema**!
