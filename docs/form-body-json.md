### Form Body JSON
*I've been unsuccessful getting the body JSON to be applied via PnP because the `CustomFormatterBody` isn't allowed in PnP powershell.  Here's the JSON body to group the fields appropriately and consistently:

## Opportunity
```json
{
    "sections": [
        {
            "displayname": "Basics",
            "fields": [
                "Opportunity Name",
                "CustomerId",
                "Opportunity Owner",
                "Expected Close Date"
            ]
        },
        {
            "displayname": "Status",
            "fields": [
                "Status",
                "Stage",
                "Next Milestone",
                "Next Deadline or Milestone"
            ]
        },
        {
            "displayname": "Financial Details",
            "fields":[    
                "Opportunity Value",
                "Win Probability",
                "Comment Log"
        ]
        }
    ]
}
```

## Customer
```json
{
    "sections": [
        {
            "displayname": "Customer Details",
            "fields": [
                "Customer Name",
                "Website",
                "NAICS code",
                "Status"
            ]
        },
        {
            "displayname": "Primary Contact",
            "fields": [
                "Primary Contact",
                "Primary Contact Title"
            ]
        },
        {
            "displayname": "Alternate 1 Contact",
            "fields": [
                "Alternate Contact",
                "Alternate Contact Title"
            ]
        },
        {
            "displayname": "Alternate 2 Contact",
            "fields": [
                "Title",
                "Alternate Contact 2",
                "Alternate Contact 2 Title"
            ]
        }
    ]
}
```