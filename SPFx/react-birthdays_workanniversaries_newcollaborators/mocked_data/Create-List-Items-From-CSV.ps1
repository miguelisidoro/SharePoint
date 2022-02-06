#REF: https://www.sharepointdiary.com/2015/09/import-csv-file-to-sharepoint-list-using-powershell.html

##Variables for Processing
$TenantRootURL = "https://createdevpt.sharepoint.com/sites/DevSite/"
$sharePointList = "User Information"
$location = Get-Location
$CSVPath = "$location\MOCK_DATA_Final.csv"
 
function CreateSharePointListItems()
{
    try
    {
        Connect-PnPOnline $TenantRootURL -UseWebLogin

        #Get he CSV file contents
        $CSVData = Import-CSV -Path $CSVPath -ErrorAction Stop

        #Iterate through each Row in the CSV and import data to SharePoint Online List
        ForEach ($Row in $CSVData)
        {
            $title = $Row.FirstName + " " + $Row.LastName
            $jobTitle = $Row.JobTitle
            $birthDate = $Row.BirthDate
            $hireDate = $Row.HireDate
            $email = $Row.Email

            Write-Host "Creating $title in List $sharePointList" -foregroundcolor Yellow
            Add-PnPListItem -List $sharePointList -Values @{"Title" = $title; "JobTitle" = $jobTitle; "BirthDate" = $birthDate; "HireDate" = $hireDate; "EMail" = $email }
            Write-Host "$title in List $sharePointList sucessfully created" -foregroundcolor Green
        }
    }
    catch
    {
        Write-Host "Error creating SharePoint List Items: $($_.Exception.Message)" -foregroundcolor Red
    }
}

CreateSharePointListItems