// Use this file to write queries to test your data connector
let
    result = SP2013Extension.Contents("http://sourcepbidemo.eastus2.cloudapp.azure.com", "'contentclass:STS_LISTITEM_DocumentLibrary'&rowlimit=500&rowsperpage=500&TrimDuplicates=false&selectproperties='Title,Author'")
in
    result