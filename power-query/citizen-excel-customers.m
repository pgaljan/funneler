let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="custList"]}[Content]{0}[Column1],
    
    // Include CustomerName field
    restUrl = siteUrl & "_api/web/lists/getbytitle('" & listName & "')/items?$select=ID,Title,CustomerName,PrimaryContact,PrimaryContactTitle,AlternateContact,AlternateContactTitle,AlternateContact2,AlternateContact2Title,Website,Modified,Created",
    
    Source = Json.Document(Web.Contents(restUrl, [
        Headers=[
            #"Accept" = "application/json;odata=verbose"
        ]
    ])),
    
    Results = Source[d][results],
    
    #"Converted to Table" = Table.FromList(Results, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
        {"ID", "Title", "CustomerName", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}, 
        {"Id", "Title", "Customer Name", "PrimaryContact", "PrimaryContactTitle", "AlternateContact", "AlternateContactTitle", "AlternateContact2", "AlternateContact2Title", "Website", "Modified", "Created"}),
    
    #"Expanded Website" = Table.ExpandRecordColumn(#"Expanded Column1", "Website", {"Url"}, {"Url"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Website",{"Title"})
in
    #"Removed Columns"