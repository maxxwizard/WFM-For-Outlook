# WFM for Outlook
Add-in for Outlook that syncs your WFM schedule into your calendar, with optional filtering of segments.

## Prerequisites
* [.NET Framework 4.5](http://www.microsoft.com/en-us/download/details.aspx?id=42643)
* Outlook 2013 / Outlook 2016
* Corpnet connectivity to resolve //wfm

## Quick-start guide
1. Exit Outlook.
2. Install the add-in by executing `\\mahuynh-z220\wfmForOutlook\WFM for Outlook.vsto`. Future updates are automatic with ClickOnce technology.
3. Launch Outlook and locate the new **WFM for Outlook** tab in the ribbon.
4. Configure **Meeting Options** to determine how synced segments appear on your calendar.
5. Click the **Sync Now** button to do an immediate pull from WFM.

## Technical details
The default settings are:
* **Exclusive** sync mode meaning all segments will be synced unless they are excluded inside the **Segment Filter**
* Research, Shift, and Meal segments are excluded
* Sync happens every 8 hours
* Each sync pulls the next 14 days of your schedule
* New meetings are prepended with "WFM: " and tagged with a yellow category

The sync logic is very basic:
1. Delete all future WFM events from your calendar.
2. Recreate them based on your WFM schedule.

## Installation issues
* If you get error 0x8007007E, try installing the [Visual Studio Tools for Office Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=48217).

## Runtime issues
* Each time WFM for Outlook does a sync, it maintains a sync log at `%APPDATA%\WFM For Outlook\sync.log`. Look in here to potentially spot what the issue is. Reach out to mahuynh@microsoft.com and attach your sync log for in-depth issues.
