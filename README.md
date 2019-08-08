
This repo shows how to create a [PowerBI](https://powerbi.microsoft.com/en-us/) Custom Connector that will authenticate a user with Forms Based Authentication and then query [SharePoint Search REST API](https://docs.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview).  With this approach, we are able to query SharePoint lists across a site collection and then return results that are security trimmed, thus easily providing Row-level security in our Power BI reports.

## SharePoint Search with ContentClass
In order to query all items across site collection of a particular list, you use the contentclass and path.  For more information, review article [here](https://blogs.msdn.microsoft.com/mvpawardprogram/2015/02/16/sharepoint-power-searching-using-contentclass/).  For example, here is query to get all list items for my list:

```
[SiteCollectionURL]/_api/search/query?querytext='Path:[SiteCollectionURL]+contentclass:STS_ListItem_GenericList'&rowlimit=500&rowsperpage=500&TrimDuplicates=false&&selectproperties='[comma seperated list of properies'
```
