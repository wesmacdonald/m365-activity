# m365-activity

This repository contains two (2) PowerShell based Azure Functions (v3):

HttpTriggerProcessActivity

TimerTriggerProcessReports

The HttpTrigger is an Azure Function whose endpoint has a subscription to the Office 365 Management Activity API.  The TimerTrigger connects to the Microsoft Graph Reports API and pulls down the data daily.

Both of these functions output the data into CosmosDB (which is not currently configured)

The Reports C# project is an example of pulling data from the Usage Analytics API (used by Power BI). 
