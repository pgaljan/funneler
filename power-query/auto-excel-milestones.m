let
    siteUrl = Excel.CurrentWorkbook(){[Name="siteUrl"]}[Content]{0}[Column1],
    listName = Excel.CurrentWorkbook(){[Name="milestoneList"]}[Content]{0}[Column1],
    
    // Include milestone fields and expand the Opportunity lookup
    restUrl = siteUrl & "_api/web/lists/getbytitle('" & listName & "')/items?$select=ID,Title,MilestoneName,Owner,MilestoneDate,MilestoneStatus,OpportunityId/ID,OpportunityId/OpportunityName,Modified,Created&$expand=OpportunityId",
    
    Source = Json.Document(Web.Contents(restUrl, [
        Headers=[
            #"Accept" = "application/json;odata=verbose"
        ]
    ])),
    
    Results = Source[d][results],
    
    #"Converted to Table" = Table.FromList(Results, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", 
        {"ID", "Title", "MilestoneName", "Owner", "MilestoneDate", "MilestoneStatus", "OpportunityId", "Modified", "Created"}, 
        {"Id", "Title", "Milestone Name", "Owner", "Milestone Date", "Milestone Status", "OpportunityId", "Modified", "Created"}),
    
    #"Expanded OpportunityId" = Table.ExpandRecordColumn(#"Expanded Column1", "OpportunityId", {"ID", "OpportunityName"}, {"Opportunity ID", "Opportunity Name"}),
    
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded OpportunityId",{
        {"Milestone Date", type datetime},
        {"Modified", type datetime},
        {"Created", type datetime}
    }),
    
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Title"})
in
    #"Removed Columns"