let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="contactList"]}[Content]{0}[Column1],
    
    // Include contact fields and expand the Customer lookup
    restUrl = siteUrl & "_api/web/lists/getbytitle('" & listName & "')/items?$select=ID,Title,ContactName,JobTitle,OfficePhone,MobilePhone,Email,CustomerId/ID,CustomerId/CustomerName,Modified,Created&$expand=CustomerId",
    
    Source = Json.Document(Web.Contents(restUrl, [
        Headers=[
            #"Accept" = "application/json;odata=verbose"
        ]
    ])),
    
    Results = Source[d][results],
    
    #"Converted to Table" = Table.FromList(Results, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
        {"ID", "Title", "ContactName", "JobTitle", "OfficePhone", "MobilePhone", "Email", "CustomerId", "Modified", "Created"}, 
        {"Id", "Title", "Contact Name", "Job Title", "Office Phone", "Mobile Phone", "Email", "CustomerId", "Modified", "Created"}),
    
    #"Expanded CustomerId" = Table.ExpandRecordColumn(#"Expanded Column1", "CustomerId", {"ID", "CustomerName"}, {"Customer ID", "Customer Name"}),
    
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded CustomerId",{
        {"Modified", type datetime},
        {"Created", type datetime}
    }),
    
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Title"})
in
    #"Removed Columns"