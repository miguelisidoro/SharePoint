#Parameters
$SiteURL = "https://createdevpt.sharepoint.com/sites/DevSite"
$ListName = "User Information"
$SourceColumn = "HireDate" #Internal Name of the Fields
$DestinationColumn = "WorkAnniversary"
 
#Connect to PnP Online
Connect-PnPOnline -Url $SiteURL -Interactive
  
#Get all items from List
$ListItems = Get-PnPListItem -List $Listname  -PageSize 500
 
#Copy Values from one column to another
ForEach ($Item in $ListItems) 
{
    Set-PnPListItem -List $Listname -Identity $Item.Id -Values @{$DestinationColumn = $Item[$SourceColumn]}
}