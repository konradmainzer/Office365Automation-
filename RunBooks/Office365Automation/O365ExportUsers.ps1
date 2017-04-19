# POWERBI CONFIGURATION, SCHEMA
$UserTable = @{
    name = "Users"; 
    columns = @(
        @{ name = "User_Id"; dataType = "string" }
       ,@{ name = "Title"; dataType = "String" }
       ,@{ name = "First_Name"; dataType = "String" }
       ,@{ name = "Last_Name"; dataType = "String" }
       ,@{ name = "Display_Name"; dataType = "string" }
       ,@{ name = "Department"; dataType = "string" }
       ,@{ name = "is_Licensed"; dataType = "boolean" }
       ,@{ name = "Mobile_Phone"; dataType = "string" }
       ,@{ name = "Office_Phone"; dataType = "string" }
       ,@{ name = "Phone"; dataType = "string" }
       ,@{ name = "Password_Never_Expires"; dataType = "boolean" }
       ,@{ name = "Street"; dataType = "string" }
       ,@{ name = "City"; dataType = "string" }
       ,@{ name = "State"; dataType = "string" }
       ,@{ name = "Postal_Code"; dataType = "string" }
       ,@{ name = "Country"; dataType = "string" }
       ,@{ name = "License"; dataType = "string" }
    ) 
}

$datasetName = "Office 365 Users"
$dataSetSchema = @{
    name = $datasetName; 
    tables = @(
        $UserTable
    )
} 

#
# Helper Functions
#

#captures debug statements globally
function Debug {
    param($text="")
    
    Write-Output $text
}

#Connect to Office 365
Debug -text ("Connecting to Office 365") -action "write" 
$o365Cred = Get-AutomationPSCredential -Name "AutomateO365Cred"
Connect-MsolService -Credential $o365Cred
Debug -text ("Connected Successfully to Office 365") -action "write"

#Connect to Power BI
Debug -text ("Connecting to Power BI") -action "write"
$powerBICred = Get-AutomationPSCredential -Name 'PowerBICred'
$authToken = Get-PBIAuthToken -ClientId "0f3ee942-2f4e-460c-9708-6c056827ebf9" -Credential $powerBICred	
$group = Get-PBIGroup -authToken $authToken -name "Office365"
Set-PBIGroup -id $group.id
Debug -text ("Connected Successfully to Power BI") -action "write"

Debug -text ("Retreiving All Users from Office 365") -action "write"	
$AllUsers = Get-MsolUser -All
Debug -text ("Successfully Retreived All Users from Office 365") -action "write"

# POWERBI CONFIGURATION AND UPLOAD
# If cannot find the DataSet create a new one with this schema. Test the existence of the dataset
Debug -text ("Working on {0}" -f "Checking DataSet") -action "write"
if (!(Test-PBIDataSet -authToken $authToken -name $datasetName) -and $AllUsers.Count -gt 0)
{  
    #create the dataset
    Debug -text ("Working on {0}" -f "Creating DataSet") -action "write"
    $pbiDataSet = New-PBIDataSet -authToken $authToken -dataSet $dataSetSchema #-verbose
}
else
{
    $pbiDataSet = Get-PBIDataSet -authToken $authToken -name $datasetName

    #clear the dataset and update the schema if necessary
    Debug -text ("Working on {0}" -f "Clearing UserTable DataSet") -action "write"
    Clear-PBITableRows -authToken $authToken -dataSetName $datasetName -tableName $UserTable.Name -verbose

    Debug -text ("Working on {0}" -f "Updating UserTable DataSet Schema") -action "write"
    Update-PBITableSchema -authToken $authToken -dataSetId $pbiDataSet.id -table $UserTable -verbose
}

$userPBIOut = @()

foreach ($User in $AllUsers)
{      
    if ($User.Licenses.Count -le 0)
    {
        $userObj = @{}        
        $userObj.User_Id = $User.UserPrincipalName
    	$userObj.Title = $User.Title
    	$userObj.First_Name = $User.FirstName
    	$userObj.Last_Name = $User.LastName
    	$userObj.Display_Name = $User.DisplayName
    	$userObj.Department = $User.Department	
    	$userObj.is_Licensed = $User.IsLicensed
    	$userObj.Mobile_Phone = $User.MobilePhone	
    	$userObj.Office_Phone = $User.OfficePhone	
    	$userObj.Phone = $User.Phone	
    	$userObj.Password_Never_Expires = $User.PasswordNeverExpires
    	$userObj.Street = $User.StreetAddress	
    	$userObj.City = $User.City
    	$userObj.State = $User.State
    	$userObj.Postal_Code = $User.PostalCode
    	$userObj.Country = $User.Country        
		$userObj.License = [string]::Empty
		
		$userPBIOut += New-Object –TypeName PSObject –Prop $userObj        
    }
    else
    {   
    	foreach ($License in $User.Licenses)
    	{
            $userObj = @{}
            $userObj.User_Id = $User.UserPrincipalName
        	$userObj.Title = $User.Title
        	$userObj.First_Name = $User.FirstName
        	$userObj.Last_Name = $User.LastName
        	$userObj.Display_Name = $User.DisplayName
        	$userObj.Department = $User.Department	
        	$userObj.is_Licensed = $User.IsLicensed
        	$userObj.Mobile_Phone = $User.MobilePhone	
        	$userObj.Office_Phone = $User.OfficePhone	
        	$userObj.Phone = $User.Phone	
        	$userObj.Password_Never_Expires = $User.PasswordNeverExpires
        	$userObj.Street = $User.StreetAddress	
        	$userObj.City = $User.City
        	$userObj.State = $User.State
        	$userObj.Postal_Code = $User.PostalCode
        	$userObj.Country = $User.Country
 		                 
            switch -Exact ($License.AccountSkuId.TrimStart("ElementMaterialsTechno:").ToUpper())
                {
                    "CAL_SERVICES" {$userObj.License = "ECAL Services (EOA, EOP, DLP)"; break} 
                    "CRMPLAN2" {$userObj.License = "Microsoft Dynamics CRM Online Basic"; break} 
                    "CRMSTANDARD" {$userObj.License = "Microsoft Dynamics CRM Online Professional"; break} 
                    "POWER_BI_STANDARD" {$userObj.License = "Power BI (free)"; break} 
                    "POWERAPPS_INDIVIDUAL_USER" {$userObj.License = "Microsoft PowerApps and Logic flows"; break} 
                    "RIGHTSMANAGEMENT" {$userObj.License = "Azure Rights Management Premium"; break} 
                    default {$userObj.License = $License.AccountSkuId.TrimStart("ElementMaterialsTechno:")}                    
                }   
   		
            $userPBIOut += New-Object –TypeName PSObject –Prop $userObj
    	}
    }          
}

# POWERBI CONFIGURATION AND UPLOAD
Debug -text ("Denormalized Office 365 Users") -action "write"    
Debug -text ("Writing {0} [{1}]" -f "UserTable DataSet Values", $userPBIOut.Count) -action "write"
$userPBIOut | Out-PowerBI -AuthToken $authToken -dataSetName $datasetName -tableName $UserTable.Name -verbose
Debug -text ("Completed") -action "write"